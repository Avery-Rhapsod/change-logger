
import { readFile, writeFile } from "node:fs/promises";
import JSZip from "jszip";

async function main(): Promise<void> {
    const originalPath = "test-files/sample.pptx";
    const outputPath = "test-files/sample-structural.pptx";

    console.log("Reading original PowerPoint...");

    const buffer = await readFile(originalPath);
    const zip = await JSZip.loadAsync(buffer);

    const slide2File = zip.file("ppt/slides/slide2.xml");

    if (!slide2File) {
        throw new Error("Could not find slide 2.");
    }

    let slide2Xml = await slide2File.async("string");

    // Remove an existing text box.
    const textToRemove =
    '<p:sp><p:nvSpPr><p:cNvPr id="7" name="TextBox 6"/>';

    const shapeStart = slide2Xml.indexOf(textToRemove);

    if (shapeStart === -1) {
        throw new Error(
            "Could not find the expected text box to remove. " +
            "The sample XML structure may be different.",
        );
    }

    const shapeEnd = slide2Xml.indexOf("</p:sp>", shapeStart);

    if (shapeEnd === -1) {
        throw new Error("Could not find the end of the text box.");
    }

    slide2Xml =
    slide2Xml.slice(0, shapeStart) +
    slide2Xml.slice(shapeEnd + "</p:sp>".length);

    // Add a new text box by duplicating an existing shape.
    const referenceStart = slide2Xml.indexOf(
        '<p:sp><p:nvSpPr><p:cNvPr id="8" name="TextBox 7"/>',
    );

    if (referenceStart === -1) {
        throw new Error(
            "Could not find the reference text box.",
        );
    }

    const referenceEnd = slide2Xml.indexOf("</p:sp>", referenceStart);

    if (referenceEnd === -1) {
        throw new Error("Could not find the reference shape end.");
    }

    let newShape = slide2Xml.slice(
        referenceStart,
        referenceEnd + "</p:sp>".length,
    );

    // Give the new shape a different PowerPoint ID and name.
    newShape = newShape
    .replace('id="8"', 'id="20"')
    .replace('name="TextBox 7"', 'name="New TextBox"')
    .replace(
        "Test equipment delivery is expected next week.",
        "A new risk review has been added.",
    );

    // Insert the new text box before the shape tree closes.
    const shapeTreeEnd = slide2Xml.lastIndexOf("</p:spTree>");

    if (shapeTreeEnd === -1) {
        throw new Error("Could not find the end of the slide shape tree.");
    }

    slide2Xml =
    slide2Xml.slice(0, shapeTreeEnd) +
    newShape +
    slide2Xml.slice(shapeTreeEnd);

    zip.file("ppt/slides/slide2.xml", slide2Xml);

    const modifiedBuffer = await zip.generateAsync({
        type: "nodebuffer",
    });

    await writeFile(outputPath, modifiedBuffer);

    console.log("\nStructural test presentation created!");
    console.log(`Output: ${outputPath}`);
    console.log("\nChanges made:");
    console.log("- Removed: Operations text box.");
    console.log("- Added: New TextBox containing a risk review.");
}

main().catch((error: unknown) => {
    console.error("Could not create structural sample:");
    console.error(error);
    process.exitCode = 1;
});
