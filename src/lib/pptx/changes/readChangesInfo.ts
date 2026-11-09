import JSZip from "jszip";
import { XMLParser } from "fast-xml-parser";
import type {
    ChangeMetadata,
    PresentationChangeRecord,
} from "./changesTypes";

const parser = new XMLParser({
    ignoreAttributes: false,
    attributeNamePrefix: "@_",
});

interface XmlObject {
    [key: string]: any;
}

function asArray<T>(value: T | T[] | undefined): T[] {
    if (value === undefined) return [];
    return Array.isArray(value) ? value : [value];
}

function normaliseZipPath(path: string): string {
    const parts: string[] = [];

    for (const part of path.replace(/^\//, "").split("/")) {
        if (part === "..") {
            parts.pop();
        } else if (part !== "." && part !== "") {
            parts.push(part);
        }
    }

    return parts.join("/");
}

function resolveRelationshipPath(target: string): string {
    if (target.startsWith("/")) {
        return normaliseZipPath(target);
    }

    return normaliseZipPath(`ppt/${target}`);
}

function getXmlChild(
    node: XmlObject,
    localName: string,
): XmlObject | undefined {
    const key = Object.keys(node).find(
        (name) => name.split(":").pop() === localName,
    );

    return key ? node[key] as XmlObject : undefined;
}

function readChangeMetadata(node: XmlObject): ChangeMetadata {
    const data = getXmlChild(node, "chgData") ?? {};

    return {
        id: data["@_id"],
        actionId: data["@_actId"],
        timestamp: data["@_dt"],
        deviceId: data["@_clId"],
        version: data["@_v"],
        author: {
            name: data["@_name"],
            email: data["@_email"],
            userId: data["@_userId"],
            providerId: data["@_providerId"],
        },
    };
}

function mergeChangeMetadata(
    parent: ChangeMetadata | undefined,
    child: ChangeMetadata,
): ChangeMetadata {
    return {
        id: child.id ?? parent?.id,
        actionId: child.actionId ?? parent?.actionId,
        timestamp: child.timestamp ?? parent?.timestamp,
        deviceId: child.deviceId ?? parent?.deviceId,
        version: child.version ?? parent?.version,
        author: {
            name: child.author.name ?? parent?.author.name,
            email: child.author.email ?? parent?.author.email,
            userId: child.author.userId ?? parent?.author.userId,
            providerId: child.author.providerId ?? parent?.author.providerId,
        },
    };
}

function readMarkerId(
    node: XmlObject,
    listName: string,
    markerName: string,
    attributeName: string,
): string | undefined {
    const markerList = getXmlChild(node, listName);
    if (!markerList) return undefined;

    const marker = asArray<XmlObject>(
        getXmlChild(markerList, markerName),
    )[0];

    const id = marker?.[`@_${attributeName}`];

    return typeof id === "string" ? id : undefined;
}

function extractChangeRecords(
    node: XmlObject,
    level: PresentationChangeRecord["level"],
    inheritedMetadata?: ChangeMetadata,
    inheritedMarkers: {
        slideId?: string;
        shapeId?: string;
        graphicFrameId?: string;
    } = {},
): PresentationChangeRecord[] {
    const records: PresentationChangeRecord[] = [];

    const ownMetadata = readChangeMetadata(node);
    const metadata = mergeChangeMetadata(
        inheritedMetadata,
        ownMetadata,
    );

    const markers = {
        slideId:
        readMarkerId(node, "sldMkLst", "sldMk", "sldId") ??
        inheritedMarkers.slideId,
        shapeId:
        readMarkerId(node, "spMkLst", "spMk", "id") ??
        inheritedMarkers.shapeId,
        graphicFrameId:
        readMarkerId(
            node,
            "graphicFrameMkLst",
            "graphicFrameMk",
            "id",
        ) ?? inheritedMarkers.graphicFrameId,
    };

    const operation = node["@_chg"];

    if (typeof operation === "string") {
        records.push({
            operation,
            level,
            metadata,
            ...markers,
        });
    }

    const childLevels: Array<{
        name: string;
        level: PresentationChangeRecord["level"];
    }> = [
        { name: "docChg", level: "document" },
        { name: "sldChg", level: "slide" },
        { name: "spChg", level: "shape" },
        { name: "graphicFrameChg", level: "graphicFrame" },
    ];

    for (const childLevel of childLevels) {
        const child = getXmlChild(node, childLevel.name);

        for (const item of asArray<XmlObject>(child)) {
            records.push(
                ...extractChangeRecords(
                    item,
                    childLevel.level,
                    metadata,
                    markers,
                ),
            );
        }
    }

    return records;
}

export async function readChangesInfo(
    fileBuffer: Buffer,
): Promise<PresentationChangeRecord[] | null> {
    const zip = await JSZip.loadAsync(fileBuffer);

    const relationshipsFile = zip.file(
        "ppt/_rels/presentation.xml.rels",
    );

    if (!relationshipsFile) {
        throw new Error(
            "Could not find the presentation relationships file.",
        );
    }

    const relationshipsXml = await relationshipsFile.async("string");
    const relationships = parser.parse(relationshipsXml) as XmlObject;

    const relationshipRoot = relationships["Relationships"];
    const relationshipItems = asArray<XmlObject>(
        relationshipRoot?.["Relationship"],
    );

    const changesInfoRelationship = relationshipItems.find(
        (relationship) =>
        typeof relationship["@_Type"] === "string" &&
        relationship["@_Type"].endsWith("/changesInfo"),
    );

    // Some presentations may not contain revision history.
    if (!changesInfoRelationship) {
        return null;
    }

    const target = changesInfoRelationship["@_Target"];

    if (typeof target !== "string") {
        throw new Error("The changesInfo relationship has no target.");
    }

    const changesInfoPath = resolveRelationshipPath(target);
    const changesInfoFile = zip.file(changesInfoPath);

    if (!changesInfoFile) {
        throw new Error(
            `Could not find the changesInfo file: ${changesInfoPath}`,
        );
    }

    const changesInfoXml = await changesInfoFile.async("string");
    const changesInfo = parser.parse(changesInfoXml) as XmlObject;

    const rootKey = Object.keys(changesInfo).find(
        (key) => key.split(":").pop() === "chgInfo",
    );

    const root = rootKey
    ? (changesInfo[rootKey] as XmlObject)
    : undefined;

    if (!root) {
        return [];
    }

    const documentChangeLists = asArray<XmlObject>(
        getXmlChild(root, "docChgLst"),
    );

    if (
        documentChangeLists.length === 0 &&
        getXmlChild(root, "docChg")
    ) {
        documentChangeLists.push(root);
    }

    const records: PresentationChangeRecord[] = [];

    for (const documentChangeList of documentChangeLists) {
        const listMetadata = readChangeMetadata(documentChangeList);

        const documentChanges = getXmlChild(
            documentChangeList,
            "docChg",
        );

        for (const item of asArray<XmlObject>(documentChanges)) {
            records.push(
                ...extractChangeRecords(
                    item,
                    "document",
                    listMetadata,
                ),
            );
        }
    }

    return records;
}
