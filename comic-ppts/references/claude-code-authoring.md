# Claude Code portability

This is a self-contained adaptation of the user's Codex comic presentation skill. Assets are local relative to this directory; no Codex plugins, API key, bundled artifact runtime or original workspace path is required for reading the skill or using its assets. The assets/reference.pptx retains the tested sample, and assets/preview.png is its first slide. artifact-template.json is reference metadata, not a Claude plugin configuration.

Use the PowerPoint authoring tools available on the destination computer. PptxGenJS, python-pptx, direct OOXML editing or installed PowerPoint automation are suitable options when available; none is bundled or silently assumed installed. Check the local environment before selecting one. Follow any applicable installed presentation skill. Keep text editable and embed image bytes rather than depending on external file paths. Work on a copy of the reference. Inspect rendered slides and the saved PPTX package for missing images, overflow and broken relationships.

For native transitions and animation, the optional scripts/apply-powerpoint-motion.ps1 requires Windows, desktop Microsoft PowerPoint and PowerShell. It is not a cross-platform helper. Its JSON plan names native shapes after deck export. Run with a fresh output filename. The example plan matches the bundled reference and should be adapted for new decks. Never strip effects through an exporter after adding them.

Example on Windows, using paths appropriate to the current project:

    powershell -NoProfile -File "<skill-folder>/scripts/apply-powerpoint-motion.ps1" -InputPath "lesson-base.pptx" -OutputPath "lesson-animated.pptx" -PlanPath "motion-plan.json" -CheckPlayback

The helper was tested on the source workstation in real PowerPoint playback. The destination installation has not yet been tested. If Windows/PowerPoint is absent, use a validated native OOXML motion workflow when available or return a complete static deck plus an explicit reveal plan. Never claim static previews prove native playback. No image-generation capability is bundled. For a new pose/cutout, use an available image tool or the exact artwork as it is; do not assume the supplied full hero illustrations already have transparent backgrounds.
