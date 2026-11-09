
import { comparePresentations } from "./comparePresentations";

import type { PresentationChange } from "./comparePresentations";
import type { PresentationContent } from "./types";
import type { PresentationVersion } from "./versionTypes";

/**
 * Compares two consecutive presentation versions.
 *
 * Changes are associated with the newer version,
 * because that is the version in which they appear.
 */
export function compareVersion(
    previous: PresentationContent,
    current: PresentationContent,
    version: PresentationVersion,
): PresentationChange[] {
    const changes = comparePresentations(previous, current);

    return changes.map((change) => ({
        ...change,

        // Make the change ID unique across versions.
        id: `${version.id}:${change.id}`,

        // Attach the metadata of the newer version.
        versionId: version.id,
        author: version.author,
        attributionType: version.attributionType,
        timestamp: version.timestamp,
    }));
}
