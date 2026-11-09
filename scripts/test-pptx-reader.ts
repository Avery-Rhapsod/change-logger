
import { readPresentation } from "../src/lib/pptx/readPresentation";

async function main(): Promise<void> {
    const filePath = process.argv[2];

    if (!filePath) {
        console.error(
            "Usage: npm run test:pptx-reader -- test-files/sample.pptx",
        );
        process.exitCode = 1;
        return;
    }

    const presentation = await readPresentation(filePath);

    console.log("\n==================================");
    console.log("     STRUCTURED PRESENTATION");
    console.log("==================================");

    console.log(`File: ${presentation.fileName}`);
    console.log(`Slides: ${presentation.slides.length}`);

    for (const slide of presentation.slides) {
        console.log(`\n--- Slide ${slide.index} ---`);

        console.log(`Text shapes: ${slide.shapes.length}`);

        for (const shape of slide.shapes) {
            console.log(`\n${shape.name} (ID: ${shape.id})`);

            for (const paragraph of shape.paragraphs) {
                console.log(`  ${paragraph}`);
            }
        }

        console.log(`\nTables: ${slide.tables.length}`);

        for (const table of slide.tables) {
            console.log(`\n${table.name} (ID: ${table.id})`);

            for (const row of table.rows) {
                console.log(`  ${row.join(" | ")}`);
            }
        }
    }

    console.log("\n==================================");
    console.log("          TEST COMPLETE");
    console.log("==================================\n");
}

main().catch((error: unknown) => {
    console.error("PowerPoint reader test failed:");
    console.error(error);
    process.exitCode = 1;
});
