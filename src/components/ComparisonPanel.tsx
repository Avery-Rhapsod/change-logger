
import { useRef, useState } from "react"
import type { PresentationChange } from "../lib/pptx/comparePresentations"

export interface ComparisonResult {
    originalFileName: string
    latestFileName: string
    summary: {
        totalChanges: number
        added: number
        removed: number
        modified: number
    }
    changes: PresentationChange[]
}

interface ComparisonPanelProps {
    onComparisonComplete: (
        result: ComparisonResult,
        originalFile: File,
        latestFile: File
    ) => void
}

function ComparisonPanel({ onComparisonComplete }: ComparisonPanelProps) {
    const [originalFile, setOriginalFile] = useState<File | null>(null)
    const [latestFile, setLatestFile] = useState<File | null>(null)
    const [loading, setLoading] = useState(false)
    const [error, setError] = useState<string | null>(null)

    const originalInputRef = useRef<HTMLInputElement>(null)
    const latestInputRef = useRef<HTMLInputElement>(null)

    function handleClear() {
        setOriginalFile(null)
        setLatestFile(null)
        setError(null)

        if (originalInputRef.current) {
            originalInputRef.current.value = ""
        }

        if (latestInputRef.current) {
            latestInputRef.current.value = ""
        }
    }

    async function handleCompare() {
        if (!originalFile || !latestFile) {
            setError("Please select both PowerPoint files.")
            return
        }

        setLoading(true)
        setError(null)

        const formData = new FormData()
        formData.append("originalFile", originalFile)
            formData.append("latestFile", latestFile)

                try {
                    const response = await fetch("/api/compare", {
                        method: "POST",
                        body: formData,
                    })

                    const data = await response.json()

                    if (!response.ok) {
                        throw new Error(data.error || "Comparison failed.")
                    }

                    const comparisonResult = data as ComparisonResult

                    onComparisonComplete(
                        comparisonResult,
                        originalFile,
                        latestFile
                    )
                } catch (err) {
                    setError(
                        err instanceof Error
                        ? err.message
                        : "An unexpected error occurred."
                    )
                } finally {
                    setLoading(false)
                }
    }

    return (
        <section className="comparison-panel">
        <div className="section-heading">
        <div>
        <h2>Compare presentations</h2>
        <p>
        Find out what changed between two PowerPoint versions.
        </p>
        </div>
        </div>

        <div className="comparison-files">
        <label className="comparison-file">
        <strong>Original version</strong>
        <span>Choose the older PowerPoint file</span>
        <input
        ref={originalInputRef}
        type="file"
        accept=".pptx"
        disabled={loading}
        onChange={(event) =>
            setOriginalFile(event.target.files?.[0] ?? null)
        }
        />
        {originalFile && <small>{originalFile.name}</small>}
        </label>

        <label className="comparison-file">
        <strong>Latest version</strong>
        <span>Choose the newer PowerPoint file</span>
        <input
        ref={latestInputRef}
        type="file"
        accept=".pptx"
        disabled={loading}
        onChange={(event) =>
            setLatestFile(event.target.files?.[0] ?? null)
        }
        />
        {latestFile && <small>{latestFile.name}</small>}
        </label>
        </div>

        <div className="comparison-actions">
        <button
        className="primary-button"
        onClick={handleCompare}
        disabled={!originalFile || !latestFile || loading}
        >
        {loading ? "Comparing..." : "Compare versions"}
        </button>

        <button
        className="secondary-button"
        onClick={handleClear}
        disabled={(!originalFile && !latestFile) || loading}
        >
        Clear
        </button>
        </div>

        {error && (
            <p className="comparison-error">{error}</p>
        )}
        </section>
    )
}

export default ComparisonPanel
