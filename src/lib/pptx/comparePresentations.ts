import type {
    PresentationContent,
    SlideContent,
    TextShape,
    TableContent,
} from "./types";
import type { PresentationChangeRecord } from "./changes/changesTypes";

export type ChangeType = "added" | "removed" | "modified";

export type ElementType =
| "slide"
| "shape"
| "table"
| "table-cell";

export interface PresentationChange {
    id: string;
    type: ChangeType;
    elementType: ElementType;
    slideIndex: number;
    slideId?: string;
    elementName: string;

    originalText: string | null;
    latestText: string | null;

    // Identifiers used to highlight the rendered PowerPoint content.
    shapeId?: string;
    paragraphIndex?: number;
    tableId?: string;
    rowIndex?: number;
    columnIndex?: number;

    versionId?: string;
    author?: string | null;
    timestamp?: string | null;
}

interface ChangeLocation {
    shapeId?: string;
    paragraphIndex?: number;
    tableId?: string;
    rowIndex?: number;
    columnIndex?: number;
}

function createChange(
    type: ChangeType,
    elementType: ElementType,
    slideIndex: number,
    elementName: string,
    originalText: string | null,
    latestText: string | null,
    location: ChangeLocation = {},
): PresentationChange {
    return {
        id: `${slideIndex}-${elementType}-${elementName}-${type}`,
        type,
        elementType,
        slideIndex,
        elementName,
        originalText,
        latestText,
        ...location,
    };
}

function compareShapes(
    originalShapes: TextShape[],
    latestShapes: TextShape[],
    slideIndex: number,
): PresentationChange[] {
    const changes: PresentationChange[] = [];

    const originalById = new Map(
        originalShapes.map((shape) => [shape.id, shape]),
    );

    const latestById = new Map(
        latestShapes.map((shape) => [shape.id, shape]),
    );

    // Compare shapes that exist in both presentations.
    for (const original of originalShapes) {
        const latest = latestById.get(original.id);

        if (!latest) {
            changes.push(
                createChange(
                    "removed",
                    "shape",
                    slideIndex,
                    original.name,
                    original.text,
                    null,
                    { shapeId: original.id },
                ),
            );
            continue;
        }

        const paragraphCount = Math.max(
            original.paragraphs.length,
            latest.paragraphs.length,
        );

        for (let index = 0; index < paragraphCount; index++) {
            const originalText = original.paragraphs[index] ?? null;
            const latestText = latest.paragraphs[index] ?? null;

            if (originalText === latestText) {
                continue;
            }

            let type: ChangeType;

            if (originalText === null) {
                type = "added";
            } else if (latestText === null) {
                type = "removed";
            } else {
                type = "modified";
            }

            changes.push(
                createChange(
                    type,
                    "shape",
                    slideIndex,
                    `${latest.name} — paragraph ${index + 1}`,
                    originalText,
                    latestText,
                    {
                        shapeId: original.id,
                        paragraphIndex: index,
                    },
                ),
            );
        }
    }

    // Detect newly added shapes.
    for (const latest of latestShapes) {
        if (!originalById.has(latest.id)) {
            changes.push(
                createChange(
                    "added",
                    "shape",
                    slideIndex,
                    latest.name,
                    null,
                    latest.text,
                    { shapeId: latest.id },
                ),
            );
        }
    }

    return changes;
}

function getTableText(table: TableContent): string {
    return table.rows
    .map((row) => row.join(" | "))
    .join("\n");
}

function compareTables(
    originalTables: TableContent[],
    latestTables: TableContent[],
    slideIndex: number,
): PresentationChange[] {
    const changes: PresentationChange[] = [];

    const originalById = new Map(
        originalTables.map((table) => [table.id, table]),
    );

    const latestById = new Map(
        latestTables.map((table) => [table.id, table]),
    );

    for (const original of originalTables) {
        const latest = latestById.get(original.id);

        if (!latest) {
            changes.push(
                createChange(
                    "removed",
                    "table",
                    slideIndex,
                    original.name,
                    getTableText(original),
                             null,
                             { tableId: original.id },
                ),
            );
            continue;
        }

        const rowCount = Math.max(
            original.rows.length,
            latest.rows.length,
        );

        const columnCount = Math.max(
            ...original.rows.map((row) => row.length),
                                     ...latest.rows.map((row) => row.length),
                                     0,
        );

        for (let row = 0; row < rowCount; row++) {
            for (let column = 0; column < columnCount; column++) {
                const originalText =
                original.rows[row]?.[column] ?? null;

                const latestText =
                latest.rows[row]?.[column] ?? null;

                if (originalText === latestText) {
                    continue;
                }

                let type: ChangeType;

                if (originalText === null) {
                    type = "added";
                } else if (latestText === null) {
                    type = "removed";
                } else {
                    type = "modified";
                }

                changes.push(
                    createChange(
                        type,
                        "table-cell",
                        slideIndex,
                        `${latest.name} — row ${row + 1}, column ${column + 1}`,
                        originalText,
                        latestText,
                        {
                            tableId: latest.id,
                            rowIndex: row,
                            columnIndex: column,
                        },
                    ),
                );
            }
        }
    }

    // Detect newly added tables.
    for (const latest of latestTables) {
        if (!originalById.has(latest.id)) {
            changes.push(
                createChange(
                    "added",
                    "table",
                    slideIndex,
                    latest.name,
                    null,
                    getTableText(latest),
                             { tableId: latest.id },
                ),
            );
        }
    }

    return changes;
}

function compareSlides(
    original: SlideContent,
    latest: SlideContent,
): PresentationChange[] {
    return [
        ...compareShapes(
            original.shapes,
            latest.shapes,
            latest.index,
        ),
        ...compareTables(
            original.tables,
            latest.tables,
            latest.index,
        ),
    ];
}

export function comparePresentations(
    original: PresentationContent,
    latest: PresentationContent,
    changeRecords: PresentationChangeRecord[] = [],
): PresentationChange[] {
    const changes: PresentationChange[] = [];

    const originalSlides = new Map(
        original.slides.map((slide) => [slide.index, slide]),
    );

    const latestSlides = new Map(
        latest.slides.map((slide) => [slide.index, slide]),
    );

    for (const originalSlide of original.slides) {
        const latestSlide = latestSlides.get(originalSlide.index);

        if (!latestSlide) {
            changes.push(
                createChange(
                    "removed",
                    "slide",
                    originalSlide.index,
                    `Slide ${originalSlide.index}`,
                    `Slide ${originalSlide.index}`,
                    null,
                ),
            );
            continue;
        }

        changes.push(
            ...compareSlides(originalSlide, latestSlide),
        );
    }

    for (const latestSlide of latest.slides) {
        if (!originalSlides.has(latestSlide.index)) {
            changes.push(
                createChange(
                    "added",
                    "slide",
                    latestSlide.index,
                    `Slide ${latestSlide.index}`,
                    null,
                    `Slide ${latestSlide.index}`,
                ),
            );
        }
    }

    console.log("Change records:", changeRecords);
    for (const change of changes) {
        const slide =
        latestSlides.get(change.slideIndex) ??
        originalSlides.get(change.slideIndex);

        change.slideId = slide?.slideId;

        const record = changeRecords.find((item) => {
            if (change.elementType === "shape" && change.shapeId) {
                return (
                    item.level === "shape" &&
                    item.slideId === change.slideId &&
                    item.shapeId === change.shapeId
                );
            }

            if (
                (change.elementType === "table" ||
                change.elementType === "table-cell") &&
                change.tableId
            ) {
                return (
                    item.level === "graphicFrame" &&
                    item.slideId === change.slideId &&
                    item.graphicFrameId === change.tableId
                );
            }

            return false;
        });

        if (change.shapeId) {
            console.log("Attribution match:", {
                changeSlideId: change.slideId,
                changeShapeId: change.shapeId,
                matchedRecord: record,
            });
        }

        if (record) {
            change.author = record.metadata.author.name ?? null;
            change.timestamp = record.metadata.timestamp ?? null;
            change.versionId = record.metadata.version;
        }
    }

    return changes.sort(
        (a, b) => a.slideIndex - b.slideIndex,
    );
}
