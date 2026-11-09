
import { useState } from "react"
import type { ComparisonResult } from "./ComparisonPanel"
import PresentationComparisonViewer from "./PresentationComparisonViewer"

interface ComparisonWorkspaceProps {
    originalFile: File
    latestFile: File
    result: ComparisonResult
    onBack: () => void
}

function ComparisonWorkspace({
    originalFile,
    latestFile,
    result,
    onBack,
}: ComparisonWorkspaceProps) {
    const [selectedSlideIndex, setSelectedSlideIndex] = useState(0)

    const slideChanges = result.changes.filter(
        (change) => change.slideIndex === selectedSlideIndex + 1
    )

    return (
        <main className="comparison-workspace">
        <header className="comparison-workspace-header">
        <div>
        <p className="eyebrow">PRESENTATION COMPARISON</p>
        <h1>Compare versions</h1>
        <p className="subtitle">
        Review the differences between your presentations.
        </p>
        </div>

        <button className="secondary-button" onClick={onBack}>
        ← Back to workspace
        </button>
        </header>

        <section className="comparison-document-headings">
        <div>
        <span>ORIGINAL VERSION</span>
        <strong>{originalFile.name}</strong>
        </div>

        <div>
        <span>LATEST VERSION</span>
        <strong>{latestFile.name}</strong>
        </div>
        </section>

        <PresentationComparisonViewer
        originalFile={originalFile}
        latestFile={latestFile}
        selectedSlideIndex={selectedSlideIndex}
        onSlideChange={setSelectedSlideIndex}
        changes={result.changes}
        />

        <section className="comparison-console">
        <div className="comparison-console-header">
        <div>
        <h2>Change console</h2>
        <p>
        Changes on slide {selectedSlideIndex + 1}
        </p>
        </div>

        <strong>
        {slideChanges.length} changes on this slide
        </strong>
        </div>

        <div className="comparison-console-summary">
        <span>{result.summary.added} added</span>
        <span>{result.summary.removed} removed</span>
        <span>{result.summary.modified} modified</span>
        </div>

        {slideChanges.length === 0 ? (
            <p>No changes were detected on this slide.</p>
        ) : (
            slideChanges.map((change) => (
                <article
                className="comparison-console-change"
                key={change.id}
                >
                <div>
                <strong>{change.elementName}</strong>
                <span className={`change-type ${change.type}`}>
                {change.type}
                </span>

                {change.author && (
                    <div className="change-author">
                    <span>Edited by {change.author}</span>

                    {change.timestamp && (
                        <time dateTime={change.timestamp}>
                        {new Date(change.timestamp).toLocaleString()}
                        </time>
                    )}
                    </div>
                )}
                </div>

                {change.type !== "added" && (
                    <p className="comparison-before">
                    − {change.originalText || "(empty)"}
                    </p>
                )}

                {change.type !== "removed" && (
                    <p className="comparison-after">
                    + {change.latestText || "(empty)"}
                    </p>
                )}
                </article>
            ))
        )}
        </section>
        </main>
    )
}

export default ComparisonWorkspace
