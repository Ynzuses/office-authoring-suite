# Source review and implementation boundary

Discovery performed with the locally installed **skills CLI 1.5.23**, using `skills find docx` and `skills find pptx` on 2026-09-10. Implementation and native validation continued on 2026-09-12. Popularity was a discovery aid, not a quality or licensing judgment. We did not run `skills add` for any third-party package.

| Reviewed source | Pinned source identity | License observed | Decision |
|---|---|---|---|
| [GitHub md-to-docx on skills.sh](https://skills.sh/github/awesome-copilot/md-to-docx) | [SKILL.md](https://github.com/github/awesome-copilot/blob/7568a482ce2df38f8965ab5336a3220db796a4ba/skills/md-to-docx/SKILL.md) | [MIT](https://github.com/github/awesome-copilot/blob/7568a482ce2df38f8965ab5336a3220db796a4ba/LICENSE) | Useful ideas: explicit dependency boundary, relative local assets, separation of input and Word formatting. Our Python generator is independently implemented; it rejects missing assets and states its limited input contract. |
| [MiniMax pptx-generator on skills.sh](https://skills.sh/minimax-ai/skills/pptx-generator) | [SKILL.md](https://github.com/MiniMax-AI/skills/blob/60aaae52bb2af8162732751a4332f62a5fef518b/skills/pptx-generator/SKILL.md) | [MIT](https://github.com/MiniMax-AI/skills/blob/60aaae52bb2af8162732751a4332f62a5fef518b/LICENSE) | Useful ideas: plan slide roles, centralize visual choices, use reusable layout functions and verify output. We independently implemented Python layouts, native charts and native Office checks. We did not import its code, fixed palette, rigid page badges or obligatory revision cycles. |
| [Anthropic docx](https://skills.sh/anthropics/skills/docx), [pptx](https://skills.sh/anthropics/skills/pptx) | Commit `41bbe19d1a1a7eaab5e7bb9050a417e5c6cffc8f` | [Skill-specific proprietary terms](https://github.com/anthropics/skills/blob/41bbe19d1a1a7eaab5e7bb9050a417e5c6cffc8f/skills/pptx/LICENSE.txt) | Excluded from source reuse and deployment; no files copied into this suite. A repository's public visibility does not establish a permissive skill license. |
| [coleam00 pptx-generator](https://skills.sh/coleam00/second-brain-skills/pptx-generator) | Commit `75e1e9cd60ef46bffbd294fdfe4d0320cee2f563`, `.claude/skills/pptx-generator/SKILL.md` | No root or skill-specific license located in the inspected tree | Excluded from source reuse. We did not adopt its mandatory brand setup, five-slide batch cap, destructive cleanup or shape-only deck concatenation. |

No third-party skill text or implementation is vendored. The first two sources informed general workflow choices; their code and palettes were not copied. Our reusable lesson about namespace prefixes stored in compatibility attribute values came from an independently diagnosed local failure and an original checker, transferred here as the single implementation authority. The test fixture generalizes the failure without retaining the private document.

Primary library references used for our implementation:

- [python-docx quickstart](https://python-docx.readthedocs.io/en/latest/user/quickstart.html) and its public API, including semantic headings, runs, tables and picture sizing.
- [python-pptx quickstart](https://python-pptx.readthedocs.io/en/latest/user/quickstart.html), its editable shapes/charts APIs and the installed package implementation.
- The installed Word and PowerPoint AppleScript dictionaries for ordinary opening, source identity, PDF export and close-without-save behavior.

The package checker is deliberately a selected-invariant checker. Neither these references nor a successful smoke test imply full OOXML conformance or universal platform compatibility.

The 0.2.0 iteration adds independently implemented image placement/provenance checks, explicit layout-risk diagnostics and an indexed PowerPoint native backend, informed by observed local authoring/export failures. The Excel authoring siblings now carry a 0.2.0 layout-readability contract covering effective type size, title/notes obstruction, freeze-pane dividers, merged cells, wrapping, detail-sheet separation, and post-edit re-rendering. Synthetic fixtures and tests contain no course content or clinical evidence. No additional third-party skill or artwork was imported. File-access localisation handling was derived from the installed application's UI; unexpected prompts remain unhandled.

## 2026-10-08 community review

The handoff workspace reviewed public directory entries and pinned raw files from `github/awesome-copilot/md-to-docx`, `MiniMax-AI/skills/pptx-generator`, `claude-office-skills/skills`, `anthropics/skills`, `athola/claude-night-market`, `obra/superpowers`, `jrabbott/skills-catalogue/preserve-user-voice` and `andrewmccall/skills/writing-voice`. Exact commit identities, URLs, retrieval status and SHA-256 values are retained in the internal batch review record; those raw third-party files are not bundled in this payload and are not needed by a receiving installation.

Useful general lessons were re-expressed independently: explicit new-vs-existing document routing; slide contracts and master-region inspection; formula smoke tests and engine-specific preservation gates; read-only Git preflight; and refusal to claim a personal voice without user-provided evidence. The proprietary Anthropic sources were comparison-only. No third-party source text or implementation was copied into the Office payload.
