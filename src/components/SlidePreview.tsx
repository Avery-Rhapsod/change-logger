
import { useEffect, useState } from "react"
import { parse, renderSlideToSvg } from "@web-ppt/core"

interface SlidePreviewProps {
    file: File
    label: string
}

function SlidePreview({ file, label }: SlidePreviewProps) {
    const [svg, setSvg] = useState<string | null>(null)
    const [loading, setLoading] = useState(true)
    const [error, setError] = useState<string | null>(null)

    useEffect(() => {
        let cancelled = false
        let presentation: Awaited<ReturnType<typeof parse>> | null = null

        setLoading(true)
        setSvg(null)
        setError(null)

        async function renderPreview() {
            try {
                presentation = await parse(file)

                if (presentation.slides.length === 0) {
                    throw new Error("This presentation contains no slides.")
                }

                const markup = renderSlideToSvg(
                    presentation,
                    presentation.slides[0]
                )

                if (cancelled) {
                    presentation.dispose?.()
                    return
                }

                setSvg(markup)
                setLoading(false)
            } catch (err) {
                presentation?.dispose?.()
                presentation = null

                if (!cancelled) {
                    setError(
                        err instanceof Error
                        ? err.message
                        : "Could not render this presentation."
                    )
                    setLoading(false)
                }
            }
        }

        void renderPreview()

        return () => {
            cancelled = true
            presentation?.dispose?.()
        }
    }, [file])

    return (
        <div className="slide-preview">
        <h3>{label}</h3>

        {loading && <p>Rendering slide...</p>}

        {error && (
            <p className="comparison-error">{error}</p>
        )}

        {svg && (
            <div
            className="slide-preview-canvas"
            dangerouslySetInnerHTML={{ __html: svg }}
            />
        )}
        </div>
    )
}

export default SlidePreview
