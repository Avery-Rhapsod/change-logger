
import { readPresentation } from "../src/lib/pptx/readPresentation";
import { comparePresentations } from "../src/lib/pptx/comparePresentations";

async function main(): Promise<void> {
    const originalPath = process.argv[2];
    const latestPath = process.argv[3];

    if (!originalPath || !latestPath) {
        console.error(
            "Usage: npm run compare:pptx -- <original.pptx> <latest.pptx>",
        );
        process.exitCode = 1;
        return;
    }

    console.log("Reading original presentation...");
    const original = await readPresentation(originalPath);

    console.log("Reading modified presentation...");
    const latest = await readPresentation(latestPath);

    console.log("\nComparing presentations...");

    const changes = comparePresentations(original, latest);

    console.log("\n==================================");
    console.log("       PRESENTATION DIFF");
    console.log("==================================");

    console.log(`Original: ${original.fileName}`);
    console.log(`Latest:   ${latest.fileName}`);
    console.log(`Changes detected: ${changes.length}`);

    for (const change of changes) {
        console.log("\n----------------------------------");
        console.log(`Slide: ${change.slideIndex}`);
        console.log(`Type: ${change.type.toUpperCase()}`);
        console.log(`Element: ${change.elementName}`);
        console.log(`Element type: ${change.elementType}`);

        console.log(`Old value: ${change.originalText ?? "(none)"}`);
        console.log(`New value: ${change.latestText ?? "(none)"}`);
    }

    console.log("\n==================================");
    console.log("       COMPARISON COMPLETE");
    console.log("==================================\n");
}

main().catch((error: unknown) => {
    console.error("Comparison failed:");
    console.error(error);
    process.exitCode = 1;
});
