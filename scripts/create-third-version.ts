
import { readFile, writeFile } from "node:fs/promises";
import JSZip from "jszip";

async function main(): Promise<void> {
    const originalPath = "test-files/sample-modified.pptx";
    const outputPath = "test-files/sample-third-version.pptx";

    const buffer = await readFile(originalPath);
    const zip = await JSZip.loadAsync(buffer);

    const slideFile = zip.file("ppt/slides/slide2.xml");

    if (!slideFile) {
        throw new Error("Could not find slide 2.");
    }

    let xml = await slideFile.async("string");

    const oldText = "Design";
    const newText = "Design and UX";

    if (!xml.includes(oldText)) {
        throw new Error(`Could not find "${oldText}" in slide 2.`);
    }

    xml = xml.replace(oldText, newText);

    zip.file("ppt/slides/slide2.xml", xml);

    const output = await zip.generateAsync({
        type: "nodebuffer",
    });

    await writeFile(outputPath, output);

    console.log("Third presentation version created!");
    console.log(`Output: ${outputPath}`);
    console.log(`Changed: ${oldText} -> ${newText}`);
}

main().catch((error: unknown) => {
    console.error("Could not create third version:");
    console.error(error);
    process.exitCode = 1;
});
