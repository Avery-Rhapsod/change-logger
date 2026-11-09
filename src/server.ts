
import express from "express";
import multer from "multer";
import { readPresentation } from "./lib/pptx/readPresentation";
import { comparePresentations } from "./lib/pptx/comparePresentations";
import { readChangesInfo } from "./lib/pptx/changes/readChangesInfo";

const app = express();

// Keep uploaded files in memory rather than writing them to disk.
const upload = multer({
    storage: multer.memoryStorage(),
                      limits: {
                          fileSize: 50 * 1024 * 1024, // 50 MB per file
                          files: 2,
                      },
                      fileFilter: (_req, file, callback) => {
                          if (!file.originalname.toLowerCase().endsWith(".pptx")) {
                              callback(new Error("Only .pptx PowerPoint files are accepted."));
                              return;
                          }

                          callback(null, true);
                      },
});

// Simple endpoint to check whether the backend is running.
app.get("/api/health", (_req, res) => {
    res.json({ status: "ok" });
});

// Compare an original presentation with a latest presentation.
app.post(
    "/api/compare",
    upload.fields([
        { name: "originalFile", maxCount: 1 },
        { name: "latestFile", maxCount: 1 },
    ]),
    async (req, res) => {
        try {
            const files = req.files as {
                originalFile?: Express.Multer.File[];
                latestFile?: Express.Multer.File[];
            };

            const originalFile = files.originalFile?.[0];
            const latestFile = files.latestFile?.[0];

            if (!originalFile || !latestFile) {
                res.status(400).json({
                    error: "Please provide both the original and latest PowerPoint files.",
                });
                return;
            }

            const original = await readPresentation(
                originalFile.buffer,
                originalFile.originalname,
            );

            const latest = await readPresentation(
                latestFile.buffer,
                latestFile.originalname,
            );

            const changeRecords =
            await readChangesInfo(latestFile.buffer) ?? [];

            const changes = comparePresentations(
                original,
                latest,
                changeRecords,
            );

            res.json({
                originalFileName: original.fileName,
                latestFileName: latest.fileName,
                summary: {
                    totalChanges: changes.length,
                    added: changes.filter((change) => change.type === "added").length,
                     removed: changes.filter((change) => change.type === "removed").length,
                     modified: changes.filter((change) => change.type === "modified").length,
                },
                changes,
                changeRecords,
            });
        } catch {
            res.status(400).json({
                error: "The PowerPoint files could not be read or compared. Please check that both files are valid .pptx presentations.",
            });
        }
    },
);

// Handle upload errors, including invalid extensions and oversized files.
app.use(
    (
        error: unknown,
     _req: express.Request,
     res: express.Response,
     _next: express.NextFunction,
    ) => {
        const message =
        error instanceof Error ? error.message : "An unexpected error occurred.";

        res.status(400).json({ error: message });
    },
);

// Listen only on this computer, not on the local network.
const port = Number(process.env.PORT ?? 3001);

app.listen(port, "127.0.0.1", () => {
    console.log(`Change Logger backend running at http://127.0.0.1:${port}`);
});
