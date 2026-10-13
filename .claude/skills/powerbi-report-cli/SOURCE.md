# Source

Copied unchanged from [microsoft/skills-for-fabric](https://github.com/microsoft/skills-for-fabric),
`skills/powerbi-report-cli/`, at commit `bc42aaf734771bf752c1d7082dd91e5a6ca72b9f`
(2026-10-05, package version 0.3.19). MIT licensed; the licence is in `LICENSE` beside this file.

To update, copy the folder again from a newer commit and change the commit above. Do not edit the
skill files here, so the copy stays comparable with upstream.

## Limits in this repo

* Preview and screenshots need Power BI Desktop on Windows or a published report in the Fabric
  service. Cloud agent sessions have neither, so the Desktop check stays a manual step.
* Agents in this repo never call a live tenant or Fabric, so the skill's `management` mode
  (publish, download, rebind) is not used here.
* The Linear issue is the approved spec. The skill's `planning` approval gate does not apply to
  issue-driven work.
