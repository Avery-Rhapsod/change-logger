
export type AttributionType =
| "verified-editor"
| "version-author"
| "unknown";

export interface PresentationVersion {
    id: string;

    // Person associated with this saved version.
    author: string | null;

    // Reliability and meaning of the author information.
    attributionType: AttributionType;

    // Date and time when this version was saved.
    timestamp: string | null;

    fileName: string;

    source: "sharepoint" | "manual";

    externalVersionId?: string;
}

export interface PresentationHistory {
    id: string;
    name: string;

    // Versions ordered from oldest to newest.
    versions: PresentationVersion[];
}
