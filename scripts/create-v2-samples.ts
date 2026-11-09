
import pptxgen from "pptxgenjs";
import fs from "node:fs";
import path from "node:path";
import JSZip from "jszip";

const pptxgenInstance = pptxgen;

const OUT = path.resolve("test-files");

fs.mkdirSync(OUT, { recursive: true });

const C = {
    navy: "182B45",
    blue: "2769A0",
    teal: "239191",
    pale: "EBF2F8",
    background: "F7F9FB",
    dark: "2A3542",
    grey: "69798A",
    white: "FFFFFF",
    green: "30855C",
    amber: "C68523",
    red: "B44545",
};

type Mode = "baseline" | "modified";

function createDeck(mode: Mode) {
    const modified = mode === "modified";

    const pptx = new pptxgenInstance();
    pptx.layout = "LAYOUT_WIDE";
    pptx.author = "Change Logger Test Generator";
    pptx.subject = "Synthetic PowerPoint comparison test";
    pptx.title = modified
    ? "Project Aurora — Modified"
    : "Project Aurora — Baseline";
    pptx.company = "Change Logger";
    pptx.lang = "en-GB";

    pptx.theme = {
        headFontFace: "Aptos Display",
        bodyFontFace: "Aptos",
        lang: "en-GB",
    };

    const W = 13.333;
    const H = 7.5;

    function text(slide: any, value: string, x: number, y: number,
                  w: number, h: number, options: any = {}) {
        slide.addText(value, {
            x, y, w, h,
            fontFace: "Aptos",
            fontSize: 11,
            color: C.dark,
            margin: 0.04,
            breakLine: false,
            valign: "mid",
            ...options,
        });
                  }

                  function heading(slide: any, title: string, subtitle: string, page: number) {
                      slide.background = { color: C.background };

                      slide.addShape(pptx.ShapeType.rect, {
                          x: 0, y: 0, w: W, h: 0.15,
                          line: { color: C.teal, transparency: 100 },
                          fill: { color: C.teal },
                      });

                      text(slide, "PROJECT AURORA  /  DELIVERY REVIEW",
                           0.62, 0.30, 12, 0.25,
                           { fontSize: 8, bold: true, color: C.teal });

                      text(slide, title, 0.62, 0.67, 12, 0.55,
                           { fontFace: "Aptos Display", fontSize: 25, bold: true, color: C.navy });

                      text(slide, subtitle, 0.64, 1.27, 12, 0.32,
                           { fontSize: 10, color: C.grey });

                      slide.addShape(pptx.ShapeType.line, {
                          x: 0.62, y: 7.08, w: 12.05, h: 0,
                          line: { color: "DAE1E8", width: 0.8 },
                      });

                      text(slide, `${String(page).padStart(2, "0")}  /  ${modified ? "07" : "06"}`,
                           11.75, 7.15, 0.9, 0.2,
                           { fontSize: 8, color: C.grey, align: "right" });
                  }

                  function section(slide: any, label: string, x: number, y: number, w = 3.3) {
                      slide.addShape(pptx.ShapeType.roundRect, {
                          x, y, w, h: 0.34,
                          rectRadius: 0.05,
                          line: { color: C.pale, transparency: 100 },
                          fill: { color: C.pale },
                      });

                      text(slide, label.toUpperCase(), x + 0.12, y + 0.04, w - 0.2, 0.23,
                           { fontSize: 9, bold: true, color: C.blue });
                  }

                  function kpi(slide: any, x: number, value: string, label: string, color = C.blue) {
                      slide.addShape(pptx.ShapeType.roundRect, {
                          x, y: 2.0, w: 2.7, h: 1.05,
                          rectRadius: 0.06,
                          line: { color: "DCE4EB", width: 0.8 },
                          fill: { color: C.white },
                      });

                      text(slide, value, x + 0.15, 2.12, 2.4, 0.4,
                           { fontFace: "Aptos Display", fontSize: 22, bold: true, color });

                      text(slide, label, x + 0.16, 2.62, 2.4, 0.2,
                           { fontSize: 8, bold: true, color: C.grey });
                  }

                  function bullets(slide: any, items: string[], x: number, y: number,
                                   w: number, fontSize = 10) {
                      items.forEach((item, i) => {
                          text(slide, "•", x, y + i * 0.34, 0.2, 0.25,
                               { fontSize: fontSize + 1, color: C.teal, bold: true });

                          text(slide, item, x + 0.24, y + i * 0.34, w - 0.25, 0.29,
                               { fontSize });
                      });
                                   }

                                   function table(slide: any, headers: string[], rows: string[][],
                                                  x: number, y: number, w: number, h: number) {
                                       slide.addTable(
                                           [
                                               headers.map(value => ({
                                                   text: value,
                                                   options: {
                                                       bold: true,
                                                       color: C.white,
                                                       fill: C.navy,
                                                   },
                                               })),
                                               ...rows.map(row => row.map(value => ({
                                                   text: value,
                                                   options: {
                                                       color: C.dark,
                                                       fill: C.white,
                                                   },
                                               }))),
                                           ],
                                           {
                                               x, y, w, h,
                                               border: { type: "solid", color: "DCE4EB", pt: 0.7 },
                                               fontFace: "Aptos",
                                               fontSize: 9,
                                               color: C.dark,
                                               margin: 0.08,
                                               rowH: 0.43,
                                               colW: Array(headers.length).fill(w / headers.length),
                                                      autoFit: false,
                                           }
                                       );
                                                  }

                                                  // --------------------------------------------------
                                                  // SLIDE 1 — PROJECT OVERVIEW
                                                  // --------------------------------------------------

                                                  {
                                                      const s = pptx.addSlide();

                                                      heading(
                                                          s,
                                                          "Project Aurora",
                                                          modified
                                                          ? "Production readiness review  •  21 October 2026  •  Internal"
                                                          : "Production readiness review  •  14 October 2026  •  Internal",
                                                          1
                                                      );

                                                      kpi(s, 0.7, modified ? "81%" : "72%", "OVERALL READINESS", C.teal);
                                                      kpi(s, 3.65, modified ? "5 / 6" : "4 / 6", "WORKSTREAMS ON TRACK");
                                                      kpi(s, 6.6, modified ? "30 Oct" : "23 Oct", "TARGET DEMONSTRATION", C.amber);
                                                      kpi(s, 9.55, modified ? "2" : "3", "OPEN DELIVERY RISKS", C.red);

                                                      section(s, "Objective and scope", 0.75, 3.45, 3.2);

                                                      text(s,
                                                           modified
                                                           ? "Objective: prepare the first production demonstrator for controlled pilot deployment.\nAdditional scope: validate audit logging and operational handover."
                                                           : "Objective: prepare the first production demonstrator for internal review.\nScope: validate the core workflow, confirm operational readiness and collect stakeholder feedback.",
                                                           0.78, 3.92, 5.65, 1.1,
                                                           { fontSize: 11, valign: "top" }
                                                      );

                                                      section(s, "Ownership and status", 6.85, 3.45, 3.2);

                                                      text(s,
                                                           modified
                                                           ? "Project owner: Production team\nExecutive sponsor: Operations director\nCurrent status: AT RISK"
                                                           : "Project owner: Production team\nExecutive sponsor: Operations director\nCurrent status: IN PROGRESS",
                                                           6.9, 3.92, 5.3, 1.0,
                                                           { fontSize: 11, valign: "top" }
                                                      );

                                                      bullets(s, modified
                                                      ? [
                                                          "Next checkpoint: controlled pilot review on 30 October.",
                                                          "Decision required: approve revised pilot access controls.",
                                                          "Audit logging added to the acceptance criteria.",
                                                      ]
                                                      : [
                                                          "Next checkpoint: integrated prototype review on 23 October.",
                                                          "Decision required: confirm the pilot environment by 19 October.",
                                                          "Escalate unresolved supplier lead times at the weekly steering meeting.",
                                                      ],
                                                      0.8, 5.35, 11.5
                                                      );

                                                      if (!modified) {
                                                          section(s, "Supplier watch", 9.3, 6.35, 2.5);
                                                          text(s, "Equipment delivery not yet confirmed.",
                                                               9.35, 6.72, 3.0, 0.25, { fontSize: 9, color: C.red });
                                                      } else {
                                                          section(s, "New control", 9.3, 6.35, 2.5);
                                                          text(s, "Pilot access review added for 26 October.",
                                                               9.35, 6.72, 3.0, 0.25, { fontSize: 9, color: C.green });
                                                      }
                                                  }

                                                  // --------------------------------------------------
                                                  // SLIDE 2 — WORKSTREAMS
                                                  // --------------------------------------------------

                                                  {
                                                      const s = pptx.addSlide();

                                                      heading(s, "Workstream updates",
                                                              "Progress, next activities and delivery ownership", 2);

                                                      const columns = [
                                                          {
                                                              x: 0.75, title: "DESIGN",
                                                              description: modified
                                                              ? "Wireframes approved; visual review completed with minor revisions."
                                                              : "Wireframes approved; visual review is scheduled for Friday.",
                                                              tasks: ["Complete accessibility review", "Prepare revised component library"],
                                                          },
                                                          {
                                                              x: 4.8, title: "ENGINEERING",
                                                              description: modified
                                                              ? "Prototype is functional; integration and audit-log testing are underway."
                                                              : "Prototype is functional; integration testing has started.",
                                                              tasks: modified
                                                              ? ["Resolve authentication and audit-log issues", "Complete regression testing", "Publish release candidate"]
                                                              : ["Resolve authentication edge cases", "Publish test build for review"],
                                                          },
                                                          {
                                                              x: 8.85, title: "OPERATIONS",
                                                              description: modified
                                                              ? "Test equipment delivery confirmed for 24 October."
                                                              : "Test equipment delivery is expected next week.",
                                                              tasks: modified
                                                              ? ["Confirm installation date", "Prepare acceptance checklist"]
                                                              : ["Confirm supplier dispatch date", "Prepare installation checklist"],
                                                          },
                                                      ];

                                                      for (const col of columns) {
                                                          section(s, col.title, col.x, 2.0, 3.55);

                                                          text(s, col.description, col.x, 2.48, 3.55, 0.8,
                                                               { fontSize: 10, valign: "top" });

                                                          bullets(s, col.tasks, col.x, 3.5, 3.6);
                                                      }

                                                      section(s, "Cross-workstream dependencies", 0.75, 5.35, 4.0);

                                                      bullets(s,
                                                              modified
                                                              ? [
                                                                  "Engineering requires final design approval before release.",
                                                              "Operations requires a stable test build before equipment acceptance.",
                                                              "Quality assurance now validates audit-log completeness.",
                                                              ]
                                                              : [
                                                                  "Engineering needs approved design tokens before the final UI build.",
                                                              "Operations requires a stable test build before equipment acceptance.",
                                                              ],
                                                              0.8, 5.8, 11.5
                                                      );

                                                      if (modified) {
                                                          section(s, "Quality assurance", 8.85, 6.55, 3.5);
                                                          text(s, "New section: automated smoke testing and release blockers.",
                                                               8.9, 6.9, 3.5, 0.25, { fontSize: 8 });
                                                      }
                                                  }

                                                  // --------------------------------------------------
                                                  // SLIDE 3 — MILESTONES
                                                  // --------------------------------------------------

                                                  {
                                                      const s = pptx.addSlide();

                                                      heading(s, "Milestones and delivery schedule",
                                                              "Baseline plan agreed at the weekly production meeting", 3);

                                                      text(s,
                                                           "The schedule tracks the checkpoints required for the first demonstrator.",
                                                           0.78, 1.85, 11.6, 0.4);

                                                      table(s,
                                                            ["Milestone", "Owner", "Target date", "Status"],
                                                            modified
                                                            ? [
                                                                ["Requirements review", "Project lead", "16 Oct 2026", "Complete"],
                                                            ["Prototype demonstration", "Engineering", "30 Oct 2026", "At risk"],
                                                            ["Internal feedback", "All teams", "02 Nov 2026", "Planned"],
                                                            ["Operational handover", "Operations", "04 Nov 2026", "Planned"],
                                                            ]
                                                            : [
                                                                ["Requirements review", "Project lead", "16 Oct 2026", "Complete"],
                                                            ["Prototype demonstration", "Engineering", "23 Oct 2026", "In progress"],
                                                            ["Internal feedback", "All teams", "27 Oct 2026", "Planned"],
                                                            ["Pilot readiness sign-off", "Operations", "30 Oct 2026", "Planned"],
                                                            ],
                                                            0.78, 2.35, 11.75, 2.8
                                                      );

                                                      section(s, "Delivery notes", 0.8, 5.5, 2.8);

                                                      bullets(s,
                                                              modified
                                                              ? [
                                                                  "The demonstration has moved by one week.",
                                                              "Operational handover is now a separate milestone.",
                                                              "A revised delivery plan is required before sign-off.",
                                                              ]
                                                              : [
                                                                  "Dates assume test equipment arrives before the installation window.",
                                                              "The project lead will publish a revised plan after the prototype demonstration.",
                                                              ],
                                                              0.82, 5.95, 11.4
                                                      );
                                                  }

                                                  // --------------------------------------------------
                                                  // SLIDE 4 — BUDGET
                                                  // --------------------------------------------------

                                                  {
                                                      const s = pptx.addSlide();

                                                      heading(s, "Budget and resource allocation",
                                                              "Planning estimate for the demonstrator and pilot preparation", 4);

                                                      kpi(s, 0.75, modified ? "€92k" : "€84k", "APPROVED BUDGET");
                                                      kpi(s, 3.7, modified ? "€68k" : "€61k", "FORECAST SPEND", C.teal);
                                                      kpi(s, 6.65, modified ? "€24k" : "€23k", "REMAINING", C.green);
                                                      kpi(s, 9.6, modified ? "7.0 FTE" : "6.5 FTE", "ALLOCATED CAPACITY", C.amber);

                                                      table(s,
                                                            ["Work package", "Allocated", "Forecast", "Variance"],
                                                            modified
                                                            ? [
                                                                ["Design and research", "€18,000", "€16,500", "+€1,500"],
                                                            ["Engineering", "€42,000", "€29,500", "+€12,500"],
                                                            ["Equipment and setup", "€22,000", "€18,500", "+€3,500"],
                                                            ["Operations support", "€10,000", "€3,500", "+€6,500"],
                                                            ]
                                                            : [
                                                                ["Design and research", "€18,000", "€16,500", "+€1,500"],
                                                            ["Engineering", "€36,000", "€27,000", "+€9,000"],
                                                            ["Equipment and setup", "€20,000", "€14,500", "+€5,500"],
                                                            ["Operations support", "€10,000", "€3,000", "+€7,000"],
                                                            ],
                                                            0.8, 3.45, 11.7, 2.55
                                                      );

                                                      text(s,
                                                           modified
                                                           ? "Forecast now includes additional integration testing and equipment contingency."
                                                           : "Forecast excludes optional pilot expansion and post-launch support.",
                                                           0.82, 6.2, 11.5, 0.35,
                                                           { fontSize: 9, color: modified ? C.red : C.grey }
                                                      );
                                                  }

                                                  // --------------------------------------------------
                                                  // SLIDE 5 — RISKS
                                                  // --------------------------------------------------

                                                  {
                                                      const s = pptx.addSlide();

                                                      heading(s, "Risks and dependencies",
                                                              "Items requiring active mitigation before pilot readiness", 5);

                                                      table(s,
                                                            ["Risk", "Impact", "Mitigation", "Owner"],
                                                            modified
                                                            ? [
                                                                ["Equipment delivery delay", "High", "Twice-daily supplier escalation", "Operations"],
                                                            ["Authentication defect", "High", "Audit-log and regression tests", "Engineering"],
                                                            ["Late design approval", "Medium", "Time-box stakeholder review", "Design"],
                                                            ["Test capacity constraint (closed)", "Low", "Backup test slot released", "Project lead"],
                                                            ]
                                                            : [
                                                                ["Equipment delivery delay", "High", "Daily supplier check-in", "Operations"],
                                                            ["Authentication defect", "Medium", "Add integration test coverage", "Engineering"],
                                                            ["Late design approval", "Medium", "Time-box stakeholder review", "Design"],
                                                            ["Limited test capacity", "Low", "Reserve backup test slot", "Project lead"],
                                                            ],
                                                            0.78, 1.95, 11.75, 2.85
                                                      );

                                                      section(s, "Escalation criteria", 0.8, 5.15, 3.1);

                                                      bullets(s,
                                                              modified
                                                              ? [
                                                                  "Escalate any critical-path delay greater than one working day.",
                                                              "Review supplier recovery plans every morning.",
                                                              "Escalate unresolved high-impact defects immediately.",
                                                              ]
                                                              : [
                                                                  "Escalate any critical-path delay greater than two working days.",
                                                              "Escalate unresolved high-impact defects at the next steering meeting.",
                                                              "Record mitigation owner and due date for every new risk.",
                                                              ],
                                                              0.82, 5.6, 11.5
                                                      );
                                                  }

                                                  // --------------------------------------------------
                                                  // SLIDE 6 — DECISIONS AND ACTIONS
                                                  // --------------------------------------------------

                                                  {
                                                      const s = pptx.addSlide();

                                                      heading(s, "Decisions and next steps",
                                                              "Actions agreed for the next delivery checkpoint", 6);

                                                      table(s,
                                                            ["Action", "Owner", "Due date", "Status"],
                                                            modified
                                                            ? [
                                                                ["Confirm pilot environment", "Operations director", "26 Oct", "Complete"],
                                                            ["Close accessibility review", "Design lead", "20 Oct", "Complete"],
                                                            ["Publish integrated test build", "Engineering lead", "24 Oct", "In progress"],
                                                            ["Prepare stakeholder demo", "Project lead", "29 Oct", "Open"],
                                                            ["Approve revised pilot controls", "Project sponsor", "27 Oct", "Open"],
                                                            ]
                                                            : [
                                                                ["Confirm pilot environment", "Operations director", "19 Oct", "Open"],
                                                            ["Close accessibility review", "Design lead", "20 Oct", "In progress"],
                                                            ["Publish integrated test build", "Engineering lead", "21 Oct", "Open"],
                                                            ["Prepare stakeholder demo", "Project lead", "22 Oct", "Open"],
                                                            ],
                                                            0.78, 1.95, 11.75, modified ? 3.0 : 2.75
                                                      );

                                                      section(s, "Decisions required", 0.8, 5.15, 2.9);

                                                      bullets(s,
                                                              modified
                                                              ? [
                                                                  "Approve revised pilot access controls.",
                                                              "Confirm the additional audit-log acceptance criteria.",
                                                              "Agree the operational handover date.",
                                                              ]
                                                              : [
                                                                  "Approve the pilot environment and access model.",
                                                              "Confirm whether the optional reporting module is in scope.",
                                                              "Agree the stakeholder group for the internal demonstration.",
                                                              ],
                                                              0.82, 5.6, 11.5
                                                      );
                                                  }

                                                  // --------------------------------------------------
                                                  // SLIDE 7 — NEW SLIDE IN MODIFIED VERSION
                                                  // --------------------------------------------------

                                                  if (modified) {
                                                      const s = pptx.addSlide();

                                                      heading(s, "Project performance and lessons learned",
                                                              "New slide  •  Review of delivery outcomes and improvements", 7);

                                                      kpi(s, 0.75, "81%", "READINESS AFTER ACTIONS", C.teal);
                                                      kpi(s, 3.7, "2", "RISKS STILL OPEN", C.amber);
                                                      kpi(s, 6.65, "5 days", "SCHEDULE RECOVERY", C.green);
                                                      kpi(s, 9.6, "92%", "TEST PASS RATE");

                                                      section(s, "What worked well", 0.8, 3.45, 3.0);

                                                      bullets(s, [
                                                          "Early design reviews reduced rework.",
                                                          "Daily integration checks improved defect visibility.",
                                                          "Named owners accelerated risk resolution.",
                                                      ], 0.82, 3.95, 5.3);

                                                      section(s, "Improvements for next phase", 6.65, 3.45, 4.0);

                                                      bullets(s, [
                                                          "Automate environment validation.",
                                                          "Agree acceptance criteria before build completion.",
                                                          "Include operational support in sprint planning.",
                                                      ], 6.7, 3.95, 5.3);

                                                      text(s,
                                                           "Next review: 4 November 2026  •  Owner: Project Management Office",
                                                           0.82, 5.7, 11.3, 0.5,
                                                           { fontSize: 10, color: C.grey, bold: true }
                                                      );
                                                  }

                                                  return pptx;
}

async function main() {
    const baseline = createDeck("baseline");
    const modified = createDeck("modified");

    await baseline.writeFile({
        fileName: path.join(OUT, "sample-v2.0.pptx"),
    });

    await modified.writeFile({
        fileName: path.join(OUT, "sample-v2.0-modified.pptx"),
    });

    console.log("Created sample-v2.0.pptx");
    console.log("Created sample-v2.0-modified.pptx");
    console.log("Baseline: 6 slides");
    console.log("Modified: 7 slides");
}

main().catch(error => {
    console.error("Sample generation failed:", error);
    process.exitCode = 1;
});
