
import { readFile } from "node:fs/promises";
import JSZip from "jszip";
import { XMLParser } from "fast-xml-parser";

type XmlObject = Record<string, any>;

const parser = new XMLParser({
    ignoreAttributes: false,
    attributeNamePrefix: "@_",
});

function asArray<T>(value: T | T[] | undefined): T[] {
    if (value === undefined) return [];
    return Array.isArray(value) ? value : [value];
}

function getText(node: any): string {
    if (!node) return "";

    if (typeof node === "string") return node;

    if (Array.isArray(node)) {
        return node.map(getText).filter(Boolean).join("");
    }

    if (typeof node === "object") {
        if (node["a:t"] !== undefined) {
            return String(node["a:t"]);
        }

        // Each PowerPoint paragraph becomes a separate line.
        if (node["a:p"] !== undefined) {
            return asArray(node["a:p"])
            .map(getText)
            .filter(Boolean)
            .join("\n");
        }

        return ["a:r", "a:fld", "a:br"]
        .filter((key) => node[key] !== undefined)
        .map((key) => getText(node[key]))
        .filter(Boolean)
        .join("");
    }

    return "";
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

async function main(): Promise<void> {
    const filePath = process.argv[2];

    if (!filePath) {
        console.error("Usage: npm run inspect:pptx -- test-files/sample.pptx");
        process.exitCode = 1;
        return;
    }

    const fileBuffer = await readFile(filePath);
    const zip = await JSZip.loadAsync(fileBuffer);

    const presentationXml = await zip.file("ppt/presentation.xml")?.async("string");
    const relationshipsXml = await zip.file("ppt/_rels/presentation.xml.rels")?.async("string");

    if (!presentationXml || !relationshipsXml) {
        throw new Error("This file does not appear to be a valid .pptx presentation.");
    }

    const presentation = parser.parse(presentationXml) as XmlObject;
    const relationships = parser.parse(relationshipsXml) as XmlObject;

    const relRoot = relationships["Relationships"] as XmlObject;
    const relItems = asArray<XmlObject>(relRoot["Relationship"]);

    const slideRelationships = new Map<string, string>();

    for (const relationship of relItems) {
        if (relationship["@_Type"]?.endsWith("/slide")) {
            slideRelationships.set(
                relationship["@_Id"],
                normaliseZipPath(`ppt/${relationship["@_Target"]}`)
            );
        }
    }

    const presentationRoot = presentation["p:presentation"] as XmlObject;
    const slideList = presentationRoot["p:sldIdLst"] as XmlObject;
    const slideIds = asArray<XmlObject>(slideList?.["p:sldId"]);

    console.log("\n==================================");
    console.log("       POWERPOINT INSPECTION");
    console.log("==================================");
    console.log(`File: ${filePath}`);
    console.log(`Slides found: ${slideIds.length}`);

    for (let index = 0; index < slideIds.length; index++) {
        const slideId = slideIds[index];
        console.log(`  PowerPoint slide ID: ${slideId["@_id"]}`);&
        const relationshipId = slideId["@_r:id"];
        const slidePath = slideRelationships.get(relationshipId);

        console.log(`\n---------- SLIDE ${index + 1} ----------`);

        if (!slidePath) {
            console.log("Could not resolve this slide's file path.");
            continue;
        }

        const slideFile = zip.file(slidePath);

        if (!slideFile) {
            console.log(`Slide file not found: ${slidePath}`);
            continue;
        }

        const slideXml = await slideFile.async("string");
        const slide = parser.parse(slideXml) as XmlObject;

        const slideRoot = slide["p:sld"] as XmlObject;
        const commonSlideData = slideRoot["p:cSld"] as XmlObject;
        const shapeTree = commonSlideData["p:spTree"] as XmlObject;

        const shapes = asArray<XmlObject>(shapeTree["p:sp"]);
        const graphicFrames = asArray<XmlObject>(shapeTree["p:graphicFrame"]);

        console.log(`Text shapes: ${shapes.length}`);
        console.log(`Graphic frames: ${graphicFrames.length}`);

        for (let shapeIndex = 0; shapeIndex < shapes.length; shapeIndex++) {
            const shape = shapes[shapeIndex];
            const properties = shape["p:nvSpPr"] as XmlObject;
            const nonVisual = properties?.["p:cNvPr"] as XmlObject;
            const name = nonVisual?.["@_name"] ?? `Shape ${shapeIndex + 1}`;
            const shapeId = nonVisual?.["@_id"];
            console.log(`  ID: ${shapeId}`);
            const text = getText(shape["p:txBody"]);

            if (text) {
                console.log(`\n[${name}]`);
                console.log(text);
            }
        }

        for (const frame of graphicFrames) {
            const frameProperties = frame["p:nvGraphicFramePr"] as XmlObject;
            const frameNonVisual = frameProperties?.["p:cNvPr"] as XmlObject;
            console.log(`  Graphic frame ID: ${frameNonVisual?.["@_id"]}`);
            const graphicData = frame["a:graphic"] as XmlObject;
            const graphic = graphicData?.["a:graphicData"] as XmlObject;
            const table = graphic?.["a:tbl"] as XmlObject;

            if (!table) continue;

            console.log("\n[TABLE]");

            const rows = asArray<XmlObject>(table["a:tr"]);

            for (const row of rows) {
                const cells = asArray<XmlObject>(row["a:tc"]);
                const values = cells.map((cell) => getText(cell["a:txBody"]));
                console.log(values.join(" | "));
            }
        }
    }

    console.log("\n==================================");
    console.log("       INSPECTION COMPLETE");
    console.log("==================================\n");
}

main().catch((error: unknown) => {
    console.error("\nInspection failed:");
    console.error(error);
    process.exitCode = 1;
});
