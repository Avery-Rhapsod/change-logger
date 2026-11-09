
import fs from "node:fs/promises";
import path from "node:path";
import pptxgen from "pptxgenjs";
import JSZip from "jszip";
import { XMLParser } from "fast-xml-parser";

const OUTPUT_DIR = path.resolve("test-files");

const PRESENTATION_NS =
"http://schemas.microsoft.com/office/powerpoint/2013/main/command";

const DRAWING_NS =
"http://schemas.microsoft.com/office/drawing/2013/main/command";

const RELATIONSHIP_TYPE =
"http://schemas.microsoft.com/office/2016/11/relationships/changesInfo";

const CONTENT_TYPE =
"application/vnd.ms-powerpoint.changesinfo+xml";

const parser = new XMLParser({
    ignoreAttributes: false,
    attributeNamePrefix: "@_",
});

type XmlObject = Record<string, any>;

const sampleChanges = [
    {
        title: "Project Aurora",
        subtitle: "Objective and scope",
        original: "Objective: prepare the first production demonstrator.",
        modified:
        "Additional scope: validate audit logging and operational handover.",
        author: "Saul Goodman (Example Legal / Test Agency)",
        userId: "S::100001@example.test::SAUL001",
        timestamp: "2026-10-21T09:15:00Z",
    },
{
    title: "Budget overview",
    subtitle: "Approved funding",
    original: "Approved budget: €85,000",
    modified: "Approved budget: €92,000",
    author: "Dallas Payday (Example Finance / Test Agency)",
    userId: "S::100002@example.test::DALLAS002",
    timestamp: "2026-10-22T11:30:00Z",
},
{
    title: "Delivery risks",
    subtitle: "Current risk assessment",
    original: "Risk status: supplier delivery remains on track.",
    modified: "Risk status: supplier delay now requires escalation.",
    author: "Mr. Scrooge (Example Operations / Test Agency)",
    userId: "S::100003@example.test::SCROOGE003",
    timestamp: "2026-10-23T14:45:00Z",
},
];

function localName(name: string): string {
    return name.split(":").pop() ?? name;
}

function child(node: XmlObject | undefined, name: string): any {
    if (!node) return undefined;

    const key = Object.keys(node).find(
        (item) => localName(item) === name,
    );

    return key ? node[key] : undefined;
}

function attribute(node: XmlObject, name: string): string | undefined {
    return node[`@_${name}`];
}

function array<T>(value: T | T[] | undefined): T[] {
    if (value === undefined) return [];
    return Array.isArray(value) ? value : [value];
}

function findRoot(node: XmlObject, name: string): XmlObject | undefined {
    for (const [key, value] of Object.entries(node)) {
        if (localName(key) === name && typeof value === "object") {
            return value as XmlObject;
        }
    }

    return undefined;
}

function collectText(node: any): string[] {
    if (node === null || node === undefined) return [];

    if (Array.isArray(node)) {
        return node.flatMap(collectText);
    }

    if (typeof node !== "object") return [];

    const result: string[] = [];

    for (const [key, value] of Object.entries(node)) {
        if (key.startsWith("@_")) continue;

        if (localName(key) === "t") {
            if (typeof value === "string") {
                result.push(value);
            } else {
                result.push(...collectText(value));
            }
        } else {
            result.push(...collectText(value));
        }
    }

    return result;
}

function getShapeText(shape: XmlObject): string {
    const textBody = child(shape, "txBody");
    return collectText(textBody).join("").trim();
}

function normaliseText(value: string): string {
    return value.replace(/\s+/g, " ").trim();
}

function normalisePath(value: string): string {
    const parts: string[] = [];

    for (const part of value.replace(/^\//, "").split("/")) {
        if (part === "..") {
            parts.pop();
        } else if (part && part !== ".") {
            parts.push(part);
        }
    }

    return parts.join("/");
}

function makeDeck(modified: boolean): pptxgen {
    const pptx = new pptxgen();

    pptx.layout = "LAYOUT_WIDE";
    pptx.author = "Change Logger synthetic test";
    pptx.company = "Change Logger";
    pptx.subject = "Synthetic author attribution test";
    pptx.title = modified
    ? "Change Logger V3 Modified"
    : "Change Logger V3 Baseline";

    for (let index = 0; index < sampleChanges.length; index++) {
        const data = sampleChanges[index];
        const slide = pptx.addSlide();

        slide.background = { color: "F5F7FA" };

        slide.addText("PROJECT AURORA / DELIVERY REVIEW", {
            x: 0.6,
            y: 0.4,
            w: 8,
            h: 0.3,
            fontFace: "Aptos",
            fontSize: 10,
            bold: true,
            color: "526174",
            margin: 0,
        });

        slide.addText(data.title, {
            x: 0.7,
            y: 1.2,
            w: 11,
            h: 0.7,
            fontFace: "Aptos Display",
            fontSize: 30,
            bold: true,
            color: "172B4D",
            margin: 0,
        });

        slide.addText(data.subtitle, {
            x: 0.7,
            y: 2.1,
            w: 10,
            h: 0.4,
            fontFace: "Aptos",
            fontSize: 16,
            color: "526174",
            margin: 0,
        });

        slide.addText(
            modified ? data.modified : data.original,
            {
                x: 0.9,
                y: 3.1,
                w: 10.5,
                h: 1.2,
                fontFace: "Aptos",
                fontSize: 21,
                color: "172B4D",
                margin: 0.15,
                valign: "mid",
                fill: { color: "FFFFFF" },
                line: { color: "D5DDE8", width: 1 },
            },
        );

        slide.addText(`${index + 1} / 3`, {
            x: 11.7,
            y: 6.8,
            w: 0.8,
            h: 0.3,
            fontSize: 10,
            color: "526174",
            align: "right",
            margin: 0,
        });
    }

    return pptx;
}

async function locateChangedShapes(zip: JSZip) {
    const presentationFile = zip.file("ppt/presentation.xml");
    const relationshipsFile = zip.file(
        "ppt/_rels/presentation.xml.rels",
    );

    if (!presentationFile || !relationshipsFile) {
        throw new Error("The PPTX is missing its presentation files.");
    }

    const presentationXml = parser.parse(
        await presentationFile.async("string"),
    ) as XmlObject;

    const relationshipsXml = parser.parse(
        await relationshipsFile.async("string"),
    ) as XmlObject;

    const presentationRoot = findRoot(presentationXml, "presentation");
    const slideList = child(presentationRoot, "sldIdLst");
    const slideEntries = array<XmlObject>(
        child(slideList, "sldId"),
    );

    const relationshipsRoot = findRoot(
        relationshipsXml,
        "Relationships",
    );

    const relationships = array<XmlObject>(
        child(relationshipsRoot, "Relationship"),
    );

    const relationshipMap = new Map<string, string>();

    for (const relationship of relationships) {
        const id = attribute(relationship, "Id");
        const target = attribute(relationship, "Target");

        if (id && target) relationshipMap.set(id, target);
    }

    const found = [];

    for (let index = 0; index < sampleChanges.length; index++) {
        const data = sampleChanges[index];
        const slideEntry = slideEntries[index];

        if (!slideEntry) {
            throw new Error(`Could not find slide ${index + 1}.`);
        }

        const slideId = attribute(slideEntry, "id");
        const relationshipId = attribute(slideEntry, "r:id");

        if (!slideId || !relationshipId) {
            throw new Error(`Slide ${index + 1} has missing identifiers.`);
        }

        const target = relationshipMap.get(relationshipId);

        if (!target) {
            throw new Error(
                `Could not resolve the relationship for slide ${index + 1}.`,
            );
        }

        const slidePath = normalisePath(
            target.startsWith("/") ? target : `ppt/${target}`,
        );

        const slideFile = zip.file(slidePath);

        if (!slideFile) {
            throw new Error(`Could not find ${slidePath}.`);
        }

        const slideXml = parser.parse(
            await slideFile.async("string"),
        ) as XmlObject;

        const slideRoot = findRoot(slideXml, "sld");
        const commonSlideData = child(slideRoot, "cSld");
        const shapeTree = child(commonSlideData, "spTree");

        if (!shapeTree) {
            throw new Error(`Could not find the shape tree on slide ${index + 1}.`);
        }

        const shapes = array<XmlObject>(child(shapeTree, "sp"));

        const matchingShape = shapes.find(
            (shape) =>
            normaliseText(getShapeText(shape)) ===
            normaliseText(data.modified),
        );

        if (!matchingShape) {
            const availableText = shapes
            .map(getShapeText)
            .filter(Boolean);

            throw new Error(
                `Could not match the edited text on slide ${index + 1}.\n` +
                `Found text: ${JSON.stringify(availableText)}`,
            );
        }

        const nonVisual = child(matchingShape, "nvSpPr");
        const properties = child(nonVisual, "cNvPr");
        const shapeId = attribute(properties, "id");

        if (!shapeId) {
            throw new Error(`Could not read the shape ID on slide ${index + 1}.`);
        }

        found.push({
            slideId,
            shapeId,
            data,
        });
    }

    return found;
}

function escapeXml(value: string): string {
    return value
    .replace(/&/g, "&amp;")
    .replace(/"/g, "&quot;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;");
}

function createChangesInfo(
    entries: Awaited<ReturnType<typeof locateChangedShapes>>,
): string {
    const lists = entries.map(({ slideId, shapeId, data }, index) => {
        const id = String(index + 1).padStart(12, "0");

        return `
        <pc:docChgLst>
        <ac:chgData
        name="${escapeXml(data.author)}"
        userId="${escapeXml(data.userId)}"
        dt="${data.timestamp}"
        id="{00000000-0000-4000-8000-${id}}"
        actId="${index + 1}"
        v="${index + 1}"
        />
        <pc:docChg chg="mod">
        <pc:docMkLst>
        <pc:docMk/>
        </pc:docMkLst>
        <pc:sldChg chg="mod">
        <pc:sldMkLst>
        <pc:sldMk sldId="${slideId}"/>
        </pc:sldMkLst>
        <ac:spChg chg="mod">
        <ac:spMkLst>
        <ac:spMk id="${shapeId}"/>
        </ac:spMkLst>
        </ac:spChg>
        </pc:sldChg>
        </pc:docChg>
        </pc:docChgLst>`;
    });

    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
    <pc:chgInfo
    xmlns:pc="${PRESENTATION_NS}"
    xmlns:ac="${DRAWING_NS}">
    ${lists.join("\n")}
    </pc:chgInfo>`;
}

async function injectChangesInfo(filePath: string): Promise<void> {
    const buffer = await fs.readFile(filePath);
    const zip = await JSZip.loadAsync(buffer);

    const entries = await locateChangedShapes(zip);
    const xml = createChangesInfo(entries);

    zip.file("ppt/changesInfo1.xml", xml);

    const relsFile = zip.file("ppt/_rels/presentation.xml.rels");

    if (!relsFile) {
        throw new Error("Could not find presentation relationships.");
    }

    const relsXml = await relsFile.async("string");
    const relsParsed = parser.parse(relsXml) as XmlObject;
    const relsRoot = findRoot(relsParsed, "Relationships");

    const existingRels = array<XmlObject>(
        child(relsRoot, "Relationship"),
    );

    const usedIds = new Set(
        existingRels.map((item) => attribute(item, "Id")),
    );

    let relationshipId = "rIdChangeInfo";
    let suffix = 2;

    while (usedIds.has(relationshipId)) {
        relationshipId = `rIdChangeInfo${suffix++}`;
    }

    const updatedRels = relsXml.replace(
        "</Relationships>",
        `<Relationship Id="${relationshipId}" Type="${RELATIONSHIP_TYPE}" Target="changesInfo1.xml"/></Relationships>`,
    );

    zip.file("ppt/_rels/presentation.xml.rels", updatedRels);

    const contentTypesFile = zip.file("[Content_Types].xml");

    if (!contentTypesFile) {
        throw new Error("Could not find [Content_Types].xml.");
    }

    const contentTypesXml = await contentTypesFile.async("string");

    const updatedContentTypes = contentTypesXml.replace(
        "</Types>",
        `<Override PartName="/ppt/changesInfo1.xml" ContentType="${CONTENT_TYPE}"/></Types>`,
    );

    zip.file("[Content_Types].xml", updatedContentTypes);

    const output = await zip.generateAsync({
        type: "nodebuffer",
        compression: "DEFLATE",
        compressionOptions: { level: 6 },
    });

    await fs.writeFile(filePath, output);

    // Verify that the part was written into the finished PPTX.
    const verificationZip = await JSZip.loadAsync(output);
    const verificationFile = verificationZip.file("ppt/changesInfo1.xml");

    if (!verificationFile) {
        throw new Error("Verification failed: changesInfo1.xml is missing.");
    }

    console.log("Added and verified changesInfo1.xml:");

    for (const entry of entries) {
        console.log(
            `  Slide ID ${entry.slideId}, shape ${entry.shapeId}: ${entry.data.author}`,
        );
    }
}

async function main(): Promise<void> {
    await fs.mkdir(OUTPUT_DIR, { recursive: true });

    const baselinePath = path.join(
        OUTPUT_DIR,
        "sample-v3.0.pptx",
    );

    const modifiedPath = path.join(
        OUTPUT_DIR,
        "sample-v3.0-modified.pptx",
    );

    const baseline = makeDeck(false);
    const modified = makeDeck(true);

    await baseline.writeFile({ fileName: baselinePath });
    await modified.writeFile({ fileName: modifiedPath });

    await injectChangesInfo(modifiedPath);

    console.log("");
    console.log("V3 sample files created:");
    console.log(`  ${baselinePath}`);
    console.log(`  ${modifiedPath}`);
}

main().catch((error) => {
    console.error("V3 generation failed:", error);
    process.exitCode = 1;
});
