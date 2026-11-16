# Authoring and portability

This skill is a self-contained folder. Assets are local and relative to the skill folder; no plugin, API key, bundled runtime or original workspace path is required for reading the skill or using its assets. The assets/reference.pptx retains the tested sample, and assets/preview.png is its first slide. artifact-template.json is reference metadata, not an assistant-specific configuration.

Use the PowerPoint authoring tools available in your environment. PptxGenJS, python-pptx, direct OOXML editing or installed PowerPoint automation are suitable options when available; none is bundled or silently assumed installed. Check the local environment before selecting one, and follow any presentation skill or workflow your assistant provides. Keep text editable and embed image bytes rather than depending on external file paths. Work on a copy of the reference. Inspect rendered slides and the saved PPTX package for missing images, overflow and broken relationships. If no authoring tool is available, say so and give a slide-by-slide plan instead.

For native transitions and animation, the optional scripts/apply-powerpoint-motion.ps1 requires Windows, desktop Microsoft PowerPoint and PowerShell. It is not a cross-platform helper. Its JSON plan names native shapes after deck export. Run with a fresh output filename. The example plan matches the bundled reference and should be adapted for new decks. Never strip effects through an exporter after adding them.

Example on Windows, using paths appropriate to the current project:

    powershell -NoProfile -File "<skill-folder>/scripts/apply-powerpoint-motion.ps1" -InputPath "lesson-base.pptx" -OutputPath "lesson-animated.pptx" -PlanPath "motion-plan.json" -CheckPlayback

The helper was tested in real PowerPoint playback on the original workstation. Other installations have not been tested. If Windows/PowerPoint is absent, use a validated native OOXML motion workflow when available or return a complete static deck plus an explicit reveal plan. Never claim static previews prove native playback. No image-generation capability is bundled. For a new pose or cutout, use an available image tool or the exact artwork as it is; do not assume the supplied full hero illustrations already have transparent backgrounds.

Image files: opaque backgrounds, hero art and cast sheets are JPEG, and transparent panes, tabs and components are PNG or SVG. See references/distribution-notes.md.
