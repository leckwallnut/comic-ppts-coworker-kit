# Comic PPTs: coworker kit

An AI skill for making editable Speaking and Listening PowerPoint lessons in a superhero comic style, starring **The Mouth** (speaking) and **The Ear** (listening). It includes 20 comic backgrounds, 20 dark-on-light comic pane layouts, header tabs and footer answer boxes, hero candidates and a reference deck. It works with Claude Code and other assistants that support skills (the open [Agent Skills](https://agentskills.io) format).

## Install it in your AI assistant

Use the desktop or command-line version of your assistant, since it needs to work with files on your computer. Paste this:

```text
Please install the Comic PPTs skill from https://github.com/leckwallnut/comic-ppts-coworker-kit

The skill is the folder comic-ppts in that repository. Copy that whole folder, keeping its name and everything inside it, into your personal skills folder so I can use it in every project. Do not change any files. Then tell me where you put it, and check that you can see a skill named comic-ppts.
```

Then start a new chat (or restart the assistant) and ask: "Which skills can you use?" The list should include `comic-ppts`.

In Claude Code the personal skills folder is `~/.claude/skills/` (on Windows, `C:\Users\<your-name>\.claude\skills\`). You can also copy the folder by hand; `SKILL.md` must sit directly inside `comic-ppts`.

## Try it

```text
Use the Comic PPTs skill (comic-ppts) to create a 12-slide Speaking and Listening lesson on asking follow-up questions. Put stage titles in header tabs and model answers in footer boxes. Include click-to-reveal answers where supported. Keep text editable and check every slide.
```

## What is inside

| Resource | Where |
| --- | --- |
| The skill | [comic-ppts](comic-ppts/SKILL.md) |
| Design rules, motion and workflow | [references](comic-ppts/references) |
| Backgrounds, panes and characters | [assets/library](comic-ppts/assets/library) |
| 3-slide reference deck and preview | [reference.pptx](comic-ppts/assets/reference.pptx), [preview.png](comic-ppts/assets/preview.png) |
| Windows PowerPoint motion helper | [apply-powerpoint-motion.ps1](comic-ppts/scripts/apply-powerpoint-motion.ps1) |

## Good to know

- Building the PowerPoint needs an authoring tool on your computer (a presentation tool, python-pptx or pptxgenjs). The motion helper needs Windows, desktop PowerPoint and PowerShell.
- The artwork is AI-generated illustration. The Ear's design and the final Mouth appearance are not yet approved; the library holds candidates.
- Motion cannot be proven by still images. Click through the finished file in Slide Show before class.
- The kit is about 107 MB, mostly images.

## License

[CC BY-NC 4.0](LICENSE): free to share and adapt with credit, not for commercial use.
