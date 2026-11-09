
import { readFile, writeFile } from "node:fs/promises";
import JSZip from "jszip";

async function main(): Promise<void> {
    const originalPath = "test-files/sample.pptx";
    const outputPath = "test-files/sample-paragraph.pptx";

    const buffer = await readFile(originalPath);
    const zip = await JSZip.loadAsync(buffer);

    const slideFile = zip.file("ppt/slides/slide1.xml");

    if (!slideFile) {
        throw new Error("Could not find slide 1.");
    }

    let xml = await slideFile.async("string");

    const oldText = "Owner: Production team";
    const newText = "Owner: Engineering team";

    if (!xml.includes(oldText)) {
        throw new Error("Could not find the original owner paragraph.");
    }

    xml = xml.replace(oldText, newText);

    zip.file("ppt/slides/slide1.xml", xml);

    const output = await zip.generateAsync({
        type: "nodebuffer",
    });

    await writeFile(outputPath, output);

    console.log("Paragraph test presentation created!");
    console.log(`Output: ${outputPath}`);
    console.log("Changed only the second paragraph of TextBox 3.");
}

main().catch((error: unknown) => {
    console.error("Could not create paragraph sample:");
    console.error(error);
    process.exitCode = 1;
});
