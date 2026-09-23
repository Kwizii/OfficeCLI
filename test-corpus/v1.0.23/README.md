# OfficeCLI v1.0.23 test corpus

This directory is a verbatim archive of the `tests/` tree from the upstream OfficeCLI tag `v1.0.23` (`470f230e29d5a97491b960748330ba4ee022bcc8`).

The greenfield implementation baseline of this fork is the current OfficeCLI `v1.0.152`/`main` commit `ffa8a0afbe2e9686abd636368e3da38c50f22131`.

The archived tests are retained as raw regression inputs only. They are not copied implementations, they do not define the new ONLYOFFICE architecture, and they are not a compatibility target. The current OfficeCLI parser and current semantic behavior remain authoritative.

Archive facts at import time:

- source tree: `tests/`
- files: 186
- C# source files: 183
- import method: `git archive v1.0.23 tests`
- contents: unchanged from the tagged source tree

`MANIFEST.sha256` records the imported file hashes so the corpus can be checked for accidental edits.
