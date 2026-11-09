
import { readFile } from "node:fs/promises";
import JSZip from "jszip";
import { XMLParser } from "fast-xml-parser";

import type {
    PresentationContent,
    SlideContent,
    TableContent,
    TextShape,
} from "./types";

type XmlObject = Record<string, any>;

const parser = new XMLParser({
    ignoreAttributes: false,
    attributeNamePrefix: "@_",
});

function asArray<T>(value: T | T[] | undefined): T[] {
    if (value === undefined) return [];
    return Array.isArray(value) ? value : [value];
}

// Collects all PowerPoint text runs (a:t) in their original order.

function extractText(node: any): string {
    if (node === null || node === undefined) return "";

    if (typeof node === "string" || typeof node === "number") {
        return String(node);
    }

    if (Array.isArray(node)) {
        return node.map(extractText).join("");
    }

    if (typeof node === "object") {
        // Only extract actual PowerPoint text, never formatting attributes.
        if (node["a:t"] !== undefined) {
            return extractText(node["a:t"]);
        }

        return Object.entries(node)
        .filter(([key]) => !key.startsWith("@_"))
        .map(([, value]) => extractText(value))
        .join("");
    }

    return "";
}


// PowerPoint stores text in paragraphs. Keeping them separate helps
// us detect paragraph-level changes later.
function extractParagraphs(textBody: any): string[] {
    if (!textBody) return [];

    const paragraphs = asArray<XmlObject>(textBody["a:p"]);

    // Keep empty paragraphs so indexes remain aligned with PowerPoint.
    return paragraphs.map((paragraph) =>
    extractText(paragraph).trim()
    );
}

function getShapeProperties(shape: XmlObject): XmlObject | undefined {
    return shape["p:nvSpPr"]?.["p:cNvPr"];
}

function getFrameProperties(frame: XmlObject): XmlObject | undefined {
    return frame["p:nvGraphicFramePr"]?.["p:cNvPr"];
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
    // Absolute package paths already start at the ZIP archive root.
    if (target.startsWith("/")) {
        return normaliseZipPath(target);
    }

    // Presentation relationships are relative to the ppt/ directory.
    return normaliseZipPath(`ppt/${target}`);
}

function getFileName(filePath: string): string {
    return filePath.split(/[\\/]/).pop() || "presentation.pptx";
}

export async function readPresentation(
    source: string | Buffer,
    fileName?: string,
): Promise<PresentationContent> {
    const fileBuffer =
    typeof source === "string"
    ? await readFile(source)
    : source;

    const resolvedFileName =
    fileName ??
    (typeof source === "string"
    ? getFileName(source)
    : "presentation.pptx");

    const zip = await JSZip.loadAsync(fileBuffer);

    const presentationFile = zip.file("ppt/presentation.xml");
    const relationshipsFile = zip.file("ppt/_rels/presentation.xml.rels");

    if (!presentationFile || !relationshipsFile) {
        throw new Error("The file is not a valid PowerPoint presentation.");
    }

    const presentationXml = await presentationFile.async("string");
    const relationshipsXml = await relationshipsFile.async("string");

    const presentation = parser.parse(presentationXml) as XmlObject;
    const relationships = parser.parse(relationshipsXml) as XmlObject;

    const relationshipRoot = relationships["Relationships"] as XmlObject;
    const relationshipItems = asArray<XmlObject>(
        relationshipRoot?.["Relationship"],
    );

    const slidePaths = new Map<string, string>();

    for (const relationship of relationshipItems) {
        const type = relationship["@_Type"];
        const id = relationship["@_Id"];
        const target = relationship["@_Target"];

        if (
            typeof type === "string" &&
            type.endsWith("/slide") &&
            typeof id === "string" &&
            typeof target === "string"
        ) {
            slidePaths.set(id, resolveRelationshipPath(target));
        }
    }

    const presentationRoot = presentation["p:presentation"] as XmlObject;
    const slideList = presentationRoot?.["p:sldIdLst"] as XmlObject;

    const slideIds = asArray<XmlObject>(slideList?.["p:sldId"]);

    const slides: SlideContent[] = [];

    for (let index = 0; index < slideIds.length; index++) {
        const relationshipId = slideIds[index]["@_r:id"];
        const slidePath = slidePaths.get(relationshipId);

        if (!slidePath) {
            throw new Error(
                `Could not find the file path for slide ${index + 1}.`,
            );
        }

        const slideFile = zip.file(slidePath);

        if (!slideFile) {
            throw new Error(
                `Could not open slide ${index + 1}: ${slidePath}`,
            );
        }

        const slideXml = await slideFile.async("string");
        const slideDocument = parser.parse(slideXml) as XmlObject;

        const slideRoot = slideDocument["p:sld"] as XmlObject;
        const commonSlideData = slideRoot?.["p:cSld"] as XmlObject;
        const shapeTree = commonSlideData?.["p:spTree"] as XmlObject;

        if (!shapeTree) {
            throw new Error(`Slide ${index + 1} has no shape tree.`);
        }

        // Extract text-containing shapes.
        const rawShapes = asArray<XmlObject>(shapeTree["p:sp"]);

        const shapes: TextShape[] = rawShapes.map((shape, shapeIndex) => {
            const properties = getShapeProperties(shape);
            const paragraphs = extractParagraphs(shape["p:txBody"]);

            return {
                id: String(properties?.["@_id"] ?? `shape-${shapeIndex + 1}`),
                                                  name: String(properties?.["@_name"] ?? `Shape ${shapeIndex + 1}`),
                                                  text: paragraphs.join("\n"),
                                                  paragraphs,
            };
        }).filter((shape) => shape.text.length > 0);

        // Extract PowerPoint tables.
        const rawFrames = asArray<XmlObject>(
            shapeTree["p:graphicFrame"],
        );

        const tables: TableContent[] = [];

        for (let frameIndex = 0; frameIndex < rawFrames.length; frameIndex++) {
            const frame = rawFrames[frameIndex];
            const graphicData = frame["a:graphic"]?.["a:graphicData"];
            const table = graphicData?.["a:tbl"];

            if (!table) continue;

            const properties = getFrameProperties(frame);
            const rawRows = asArray<XmlObject>(table["a:tr"]);

            const rows = rawRows.map((row) => {
                const cells = asArray<XmlObject>(row["a:tc"]);

                return cells.map((cell) =>
                extractParagraphs(cell["a:txBody"]).join("\n"),
                );
            });

            tables.push({
                id: String(properties?.["@_id"] ?? `table-${frameIndex + 1}`),
                        name: String(properties?.["@_name"] ?? `Table ${frameIndex + 1}`),
                        rows,
            });
        }

        slides.push({
            index: index + 1,
            slideId: String(slideIds[index]["@_id"]),
                    shapes,
                    tables,
        });
    }

    return {
        fileName: resolvedFileName,
        slides,
    };
}
