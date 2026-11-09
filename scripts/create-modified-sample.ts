
import { readFile, writeFile } from "node:fs/promises";
import JSZip from "jszip";

async function main(): Promise<void> {
    const originalPath = "test-files/sample.pptx";
    const modifiedPath = "test-files/sample-modified.pptx";

    console.log("Reading original PowerPoint...");

    const fileBuffer = await readFile(originalPath);
    const zip = await JSZip.loadAsync(fileBuffer);

    // Modify a sentence on slide 2.
    const slide2File = zip.file("ppt/slides/slide2.xml");

    if (!slide2File) {
        throw new Error("Could not find slide 2.");
    }

    let slide2Xml = await slide2File.async("string");

    const oldDescription =
    "visual review is scheduled for Friday.";

        const newDescription =
        "visual review is scheduled for Monday.";

    if (!slide2Xml.includes(oldDescription)) {
        throw new Error("Could not find the expected text on slide 2.");
    }

    slide2Xml = slide2Xml.replace(oldDescription, newDescription);

    zip.file("ppt/slides/slide2.xml", slide2Xml);

    // Modify a date in the table on slide 3.
    const slide3File = zip.file("ppt/slides/slide3.xml");

    if (!slide3File) {
        throw new Error("Could not find slide 3.");
    }

    let slide3Xml = await slide3File.async("string");

    if (!slide3Xml.includes("23 Oct 2026")) {
        throw new Error("Could not find the expected date on slide 3.");
    }

    slide3Xml = slide3Xml.replace(
        "23 Oct 2026",
        "24 Oct 2026",
    );

    zip.file("ppt/slides/slide3.xml", slide3Xml);

    // Generate a new PowerPoint without modifying the original.
    const modifiedBuffer = await zip.generateAsync({
        type: "nodebuffer",
    });

    await writeFile(modifiedPath, modifiedBuffer);

    console.log("\nModified PowerPoint created successfully!");
    console.log(`Output: ${modifiedPath}`);
    console.log("\nChanges made:");
    console.log("- Slide 2: Friday changed to Monday.");
    console.log("- Slide 3: 23 Oct 2026 changed to 24 Oct 2026.");
}

main().catch((error: unknown) => {
    console.error("Could not create modified PowerPoint:");
    console.error(error);
    process.exitCode = 1;
});
