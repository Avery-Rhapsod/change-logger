
import { readPresentation } from "../src/lib/pptx/readPresentation";
import { compareVersion } from "../src/lib/pptx/compareVersion";

async function main(): Promise<void> {
    const originalPath = "test-files/sample.pptx";
    const latestPath = "test-files/sample-modified.pptx";

    console.log("Reading presentation versions...");

    const original = await readPresentation(originalPath);
    const latest = await readPresentation(latestPath);

    // Simulated metadata for testing.
    const version = {
        id: "version-2",
        author: "Alice Martin",
        timestamp: "2026-10-22T14:32:00.000Z",
        fileName: "sample-modified.pptx",
        source: "manual" as const,
    };

    console.log("Comparing versions...");

    const changes = compareVersion(original, latest, version);

    console.log("\n==================================");
    console.log("   VERSION ATTRIBUTION TEST");
    console.log("==================================");

    console.log(`Version ID: ${version.id}`);
    console.log(`Author: ${version.author}`);
    console.log(`Timestamp: ${version.timestamp}`);
    console.log(`Changes detected: ${changes.length}`);

    for (const change of changes) {
        console.log("\n----------------------------------");
        console.log(`Slide: ${change.slideIndex + 1}`);
        console.log(`Element: ${change.elementName}`);
        console.log(`Type: ${change.type}`);
        console.log(`Old value: ${change.originalText}`);
        console.log(`New value: ${change.latestText}`);
        console.log(`Version: ${change.versionId}`);
        console.log(`Author: ${change.author}`);
        console.log(`Timestamp: ${change.timestamp}`);
    }

    console.log("\n==================================");
    console.log("   TEST COMPLETE");
    console.log("==================================");
}

main().catch((error: unknown) => {
    console.error("Version comparison failed:");
    console.error(error);
    process.exitCode = 1;
});
