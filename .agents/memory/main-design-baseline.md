---
name: Main and design baseline
description: Workflow rule for preserving the stable project baseline while design work is implemented and handed off locally.
---

Keep the legacy/main project available as the stable testing baseline while design work is developed separately. Approved design changes should be carried through the tracked implementation before local handoff, so the local PC receives a complete pull rather than requiring the user to recreate design work.

**Why:** The user needs ongoing testing against the main project and wants to avoid repeated design-to-local rework.

**How to apply:** Do not replace or destabilize the main baseline during design work. Track design and implementation changes together, verify the merged branch, and provide the local handoff from that complete branch.