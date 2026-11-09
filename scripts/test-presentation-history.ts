
import { buildPresentationHistory } from "../src/lib/pptx/buildPresentationHistory";

import type { PresentationHistory } from "../src/lib/pptx/versionTypes";

async function main(): Promise<void> {
    const history: PresentationHistory = {
        id: "project-aurora",
        name: "Project Aurora",

        versions: [
            {
                id: "version-1",
                author: null,
                timestamp: null,
                fileName: "test-files/sample.pptx",
                source: "manual",
            },
            {
                id: "version-2",
                author: "Alice Martin",
                timestamp: "2026-10-22T14:32:00.000Z",
                fileName: "test-files/sample-modified.pptx",
                source: "manual",
            },
            {
                id: "version-3",
                author: "Thomas Bernard",
                timestamp: "2026-10-23T17:48:00.000Z",
                fileName: "test-files/sample-third-version.pptx",
                source: "manual",
            },
        ],
    };

    console.log("Building presentation history...");

    const results = await buildPresentationHistory(history);

    console.log("\n==================================");
    console.log("   PRESENTATION ACTIVITY HISTORY");
    console.log("==================================");

    for (const result of results) {
        console.log("\n==================================");
        console.log(`Version: ${result.version.id}`);
        console.log(`Author: ${result.version.author ?? "Unknown"}`);
        console.log(`Timestamp: ${result.version.timestamp ?? "Unknown"}`);
        console.log(`Changes: ${result.changes.length}`);

        for (const change of result.changes) {
            console.log("\n----------------------------------");
            console.log(`Slide: ${change.slideIndex + 1}`);
            console.log(`Element: ${change.elementName}`);
            console.log(`Type: ${change.type}`);
            console.log(`Old value: ${change.originalText}`);
            console.log(`New value: ${change.latestText}`);
            console.log(`Attributed to: ${change.author ?? "Unknown"}`);
        }
    }

    console.log("\n==================================");
    console.log("   HISTORY TEST COMPLETE");
    console.log("==================================");
}

main().catch((error: unknown) => {
    console.error("History test failed:");
    console.error(error);
    process.exitCode = 1;
});
