
import { useState } from "react"
import type { ComparisonResult } from "./components/ComparisonPanel"
import "./App.css"
import ComparisonPanel from "./components/ComparisonPanel"
import ComparisonWorkspace from "./components/ComparisonWorkspace"

function App() {
  const [activeComparison, setActiveComparison] = useState<{
    result: ComparisonResult
    originalFile: File
    latestFile: File
  } | null>(null)

  function handleComparisonComplete(
    result: ComparisonResult,
    originalFile: File,
    latestFile: File,
  ) {
    setActiveComparison({
      result,
      originalFile,
      latestFile,
    })
  }

  if (activeComparison) {
    return (
      <ComparisonWorkspace
      result={activeComparison.result}
      originalFile={activeComparison.originalFile}
      latestFile={activeComparison.latestFile}
      onBack={() => setActiveComparison(null)}
      />
    )
  }

  return (
    <div className="app">
    <main className="main">
    <div className="content">
    <div className="page-heading">
    <div>
    <p className="eyebrow">CHANGE LOGGER</p>
    <h1>Compare PowerPoint presentations</h1>
    <p className="subtitle">
    Select two versions to see what changed.
    </p>
    </div>
    </div>

    <ComparisonPanel
    onComparisonComplete={handleComparisonComplete}
    />
    </div>
    </main>
    </div>
  )
}

export default App
