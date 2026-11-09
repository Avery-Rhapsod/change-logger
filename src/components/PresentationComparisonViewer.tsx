
import { useEffect, useState } from "react"
import {
    parse,
    renderSlideToSvg,
    type Presentation,
} from "@web-ppt/core"
import type {
    PresentationChange,
    ChangeType,
} from "../lib/pptx/comparePresentations"

interface PresentationComparisonViewerProps {
    originalFile: File
    latestFile: File
    selectedSlideIndex: number
    onSlideChange: (index: number) => void
    changes: PresentationChange[]
}

interface LoadedPresentations {
    original: Presentation
    latest: Presentation
}

interface RenderedSlides {
    original: string | null
    latest: string | null
}

const highlightColors: Record<ChangeType, string> = {
    added: "#16a34a",
    removed: "#dc2626",
    modified: "#d97706",
}

function addSvgTooltip(element: Element, text: string) {
    const title = element.ownerDocument.createElementNS(
        "http://www.w3.org/2000/svg",
        "title",
    )

    title.textContent = text
    element.insertBefore(title, element.firstChild)
}

function highlightSlide(
    svg: string,
    changes: PresentationChange[],
    side: "original" | "latest",
): string {
    console.log(
        "AUTHORS RECEIVED BY VIEWER:",
        changes.map(change => ({
            type: change.elementType,
            author: change.author,
            slide: change.slideId,
        }))
    );

    if (changes.length === 0) return svg;

        const document = new DOMParser().parseFromString(
            svg,
            "image/svg+xml",
        )

        const root = document.documentElement

        if (document.querySelector("parsererror")) {
            return svg
        }

        for (const change of changes) {
            console.log("CHANGE DETECTED:", change)
            // Added content only exists in the latest version.
            if (change.type === "added" && side === "original") continue

                // Removed content only exists in the original version.
                if (change.type === "removed" && side === "latest") continue

                    const color = highlightColors[change.type]
                    const tooltip = [
                        change.author ? `Edited by: ${change.author}` : "Author unknown",
                        change.timestamp
                        ? `Date: ${new Date(change.timestamp).toLocaleString()}`
                        : null,
                    ].filter(Boolean).join(" | ")

                    console.log(
                        "SHAPE IDS IN SVG:",
                        Array.from(root.querySelectorAll("[data-el]")).map(
                            node => node.getAttribute("data-el")
                        )
                    );

                    console.log("LOOKING FOR SHAPE ID:", change.shapeId);
                    const element = Array.from(
                        root.querySelectorAll("[data-el]"),
                    ).find(
                        (node) => node.getAttribute("data-el") ===
                        (change.shapeId ?? change.tableId),
                    )

                    if (!element) continue
                        addSvgTooltip(element, tooltip)

                        // Highlight an individual table cell.
                        if (change.elementType === "table-cell") {
                            const cell = element.querySelector(
                                `[data-table-cell="${change.rowIndex}:${change.columnIndex}"]`
                            )

                            if (cell) {
                                addSvgTooltip(cell, tooltip);
                                const runs = cell.querySelectorAll("span[data-r]");

                                runs.forEach((run) => {
                                    run.setAttribute("data-tooltip", tooltip);
                                });

                                const rect = cell.querySelector("rect")

                                if (rect) {
                                    addSvgTooltip(rect, tooltip);

                                    rect.setAttribute(
                                        "fill",
                                        change.type === "added"
                                        ? "rgba(22, 163, 74, 0.30)"
                                        : change.type === "removed"
                                        ? "rgba(220, 38, 38, 0.30)"
                                        : "rgba(217, 119, 6, 0.35)"
                                    )

                                    rect.setAttribute(
                                        "stroke",
                                        highlightColors[change.type]
                                    )

                                    rect.setAttribute("stroke-width", "3")
                                }

                                cell.setAttribute(
                                    "filter",
                                    `drop-shadow(0 0 5px ${highlightColors[change.type]})`
                                )
                            }

                            continue
                        }

                        // Highlight an entire added or removed shape.
                        // Highlight an entire added or removed shape.
                        if (
                            change.elementType === "shape" &&
                            change.paragraphIndex === undefined
                        ) {
                            // Highlight the shape's text runs.
                            const runs = element.querySelectorAll("span[data-r]");

                            runs.forEach((run) => {
                                const htmlRun = run as HTMLElement;

                                htmlRun.style.display = "inline-block";
                                htmlRun.style.padding = "2px 4px";
                                htmlRun.style.backgroundColor = `${color}33`;
                                htmlRun.style.outline = `2px solid ${color}`;
                                htmlRun.style.borderRadius = "2px";
                                htmlRun.setAttribute("data-tooltip", tooltip);
                            });

                            // Also highlight the shape's background, if it has one.
                            const bounds = element.querySelector("rect");

                            if (bounds) {
                                bounds.setAttribute(
                                    "fill",
                                    change.type === "added"
                                    ? "rgba(22, 163, 74, 0.25)"
                                    : "rgba(220, 38, 38, 0.25)"
                                );

                                bounds.setAttribute("stroke", color);
                                bounds.setAttribute("stroke-width", "3");
                            } else {
                                element.setAttribute(
                                    "filter",
                                    `drop-shadow(0px 0px 4px ${color})`
                                );
                            }

                            continue;
                        }

                            // Highlight the changed paragraph's rendered text runs.
                            if (
                                change.elementType === "shape" &&
                                change.paragraphIndex !== undefined
                            ) {
                                const paragraph = Array.from(
                                    element.querySelectorAll("[data-p]"),
                                ).find(
                                    (node) =>
                                    node.getAttribute("data-p") ===
                                    String(change.paragraphIndex),
                                )

                                if (!paragraph) continue

                                    console.log("PARAGRAPH HTML:", paragraph.outerHTML);
                                    const runs = paragraph.querySelectorAll("span[data-r]")

                                    if (runs.length === 0) {
                                        paragraph.setAttribute(
                                            "style",
                                            `${paragraph.getAttribute("style") ?? ""};background-color:${color}33;outline:2px solid ${color}`,
                                        )
                                    }

                                    runs.forEach((run) => {
                                        const htmlRun = run as HTMLElement;
                                        htmlRun.style.display = "inline-block";
                                        htmlRun.style.padding = "2px 4px";
                                        htmlRun.style.backgroundColor = `${color}33`;
                                        htmlRun.style.outline = `2px solid ${color}`;
                                        htmlRun.style.borderRadius = "2px";
                                        htmlRun.setAttribute("data-tooltip", tooltip);
                                    });
                            }

                            // Highlight an entire table if the table itself was added or removed.
                            if (change.elementType === "table") {
                                element.setAttribute(
                                    "filter",
                                    `drop-shadow(0px 0px 4px ${color})`,
                                )
                            }
        }

        console.log(new XMLSerializer().serializeToString(root))
        return new XMLSerializer().serializeToString(root)
}

function PresentationComparisonViewer({
    originalFile,
    latestFile,
    selectedSlideIndex,
    onSlideChange,
    changes,
}: PresentationComparisonViewerProps) {
    const [presentations, setPresentations] =
    useState<LoadedPresentations | null>(null)

    const [renderedSlides, setRenderedSlides] =
    useState<RenderedSlides | null>(null)

    const [loading, setLoading] = useState(true)
    const [error, setError] = useState<string | null>(null)

    // Load both PowerPoint files.
    useEffect(() => {
        let cancelled = false
        let original: Presentation | null = null
        let latest: Presentation | null = null

        setPresentations(null)
        setRenderedSlides(null)
        setLoading(true)
        setError(null)

        async function loadPresentations() {
            try {
                original = await parse(originalFile)
                latest = await parse(latestFile)

                if (
                    original.slides.length === 0 ||
                    latest.slides.length === 0
                ) {
                    throw new Error(
                        "One of the presentations contains no slides.",
                    )
                }

                if (cancelled) {
                    original.dispose?.()
                    latest.dispose?.()
                    return
                }

                setPresentations({ original, latest })
                setLoading(false)
            } catch (err) {
                original?.dispose?.()
                latest?.dispose?.()

                if (!cancelled) {
                    setError(
                        err instanceof Error
                        ? err.message
                        : "Could not load the presentations.",
                    )
                    setLoading(false)
                }
            }
        }

        void loadPresentations()

        return () => {
            cancelled = true
            original?.dispose?.()
            latest?.dispose?.()
        }
    }, [originalFile, latestFile])

    // Render and highlight the selected slide.
    useEffect(() => {
        if (!presentations) return

            try {
                const originalSlide =
                presentations.original.slides[selectedSlideIndex]

                const latestSlide =
                presentations.latest.slides[selectedSlideIndex]

                const slideChanges = changes.filter(
                    (change) =>
                    change.slideIndex === selectedSlideIndex + 1,
                )

                const originalSvg = originalSlide
                ? renderSlideToSvg(
                    presentations.original,
                    originalSlide,
                    {
                        includeEditMarkers: true,
                        textMode: "html",
                    },
                )
                : null

                const latestSvg = latestSlide
                ? renderSlideToSvg(
                    presentations.latest,
                    latestSlide,
                    {
                        includeEditMarkers: true,
                        textMode: "html",
                    },
                )
                : null

                setRenderedSlides({
                    original: originalSvg
                    ? highlightSlide(
                        originalSvg,
                        slideChanges,
                        "original",
                    )
                    : null,

                    latest: latestSvg
                    ? highlightSlide(
                        latestSvg,
                        slideChanges,
                        "latest",
                    )
                    : null,
                })

                setError(null)
            } catch (err) {
                setError(
                    err instanceof Error
                    ? err.message
                    : "Could not render this slide.",
                )
            }
    }, [presentations, selectedSlideIndex, changes])

    const slideCount = presentations
    ? Math.max(
        presentations.original.slides.length,
        presentations.latest.slides.length,
    )
    : 0

    function previousSlide() {
        onSlideChange(Math.max(0, selectedSlideIndex - 1))
    }

    function nextSlide() {
        onSlideChange(
            Math.min(slideCount - 1, selectedSlideIndex + 1),
        )
    }

    return (
        <section className="presentation-comparison">
        <div className="comparison-viewer-heading">
        <div>
        <h3>Slide comparison</h3>
        <p>
        Both presentations stay synchronized as you navigate.
        </p>
        </div>

        {presentations && (
            <div className="slide-navigation">
            <button
            type="button"
            onClick={previousSlide}
            disabled={selectedSlideIndex === 0}
            aria-label="Previous slide"
            >
            ‹
            </button>

            <span>
            Slide {selectedSlideIndex + 1} of {slideCount}
            </span>

            <button
            type="button"
            onClick={nextSlide}
            disabled={selectedSlideIndex >= slideCount - 1}
            aria-label="Next slide"
            >
            ›
            </button>
            </div>
        )}
        </div>

        {loading && <p>Loading presentations...</p>}

        {error && (
            <p className="comparison-error">{error}</p>
        )}

        {presentations && (
            <div className="slide-comparison-grid">
            <div className="slide-preview">
            <h4>{originalFile.name}</h4>

            {renderedSlides?.original ? (
                <div
                className="slide-preview-canvas"
                dangerouslySetInnerHTML={{
                    __html: renderedSlides.original,
                }}
                />
            ) : (
                <p>
                This slide does not exist in the original.
                </p>
            )}
            </div>

            <div className="slide-preview">
            <h4>{latestFile.name}</h4>

            {renderedSlides?.latest ? (
                <div
                className="slide-preview-canvas"
                dangerouslySetInnerHTML={{
                    __html: renderedSlides.latest,
                }}
                />
            ) : (
                <p>
                This slide does not exist in the latest.
                </p>
            )}
            </div>
            </div>
        )}
        </section>
    )
}

export default PresentationComparisonViewer
