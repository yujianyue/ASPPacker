# ASP Single-File EXE Packer (Windows)

> **v1.1 Changes (important)**:
> 1. **Fixed "crash right after packing"**: the old build allocated the zip buffer without counting the central-directory copy of file names; zips with many files overflowed the heap and the process died the moment packing finished (right after the output file was written). Fixed and verified with AddressSanitizer (old formula always overflows, new one is clean); output zips also pass standard `zipfile` validation.
> 2. **Fixed "every page returns 404" in the packed exe**: the relative runtime folder (wwwroot) embedded via app.ini was never converted to an absolute path, so a freshly generated exe (no on-disk ini yet) served 404 for everything. Fixed.
> 3. **Crash self-diagnostics**: both the packer and the packed exe now write `<ProgramName>.crash.log` (exception code + address) next to the exe instead of dying silently.
> 4. **Multi-level data folders**: seed files are extracted with automatic sub-directory creation (e.g. `wwwroot\data\wage.mdb`).
> 5. **Customizable extraction-to-disk rules**: two new fields — "Data dirs" (default `/data/|/upload/`) and "Cache extensions" (default `.json|.ini`). Everything goes into the exe, but **only matched files are extracted to disk** for read/write; everything else stays memory-only.
> 6. Companion document: **ASP-Packaging-Guide.md** (how to prepare your ASP zip, supported ASP features, pre-flight checklist, FAQ).

## 1. What is this
Packs an **ASP program zip** together with the built-in **aspfox mini ASP server** into one standalone Windows exe. Double-click to serve over local / intranet / public IP+port. No IIS, no installation, code invisible (embedded in the exe).

## 2. How to use
1. Put `Packer.exe` on a Windows machine and **double-click it** (single file, nothing to install).
   - On start: the **Output folder** defaults to the folder containing `Packer.exe`; if that folder contains **exactly one** `.zip` / `.ico`, it is auto-filled into "ASP program zip" / "Icon". File dialogs also start there — paths are always built as "program folder + relative path" for easy distribution.
2. Fill in the fields (see section 3). The **ASP program zip** is required:
   - Zip your ASP site with any tool (7-Zip / `zip` / Windows built-in), containing `index.asp` etc.
   - If the whole site is zipped inside one top-level folder (e.g. `MySite/index.asp`), the packer **strips that single common top-level prefix** automatically so `index.asp` lands at the root; if the zip has root-level files alongside folders, nothing is stripped.
   - **Icon ico** is optional; leave empty for the built-in default icon.
3. Click **Pack**; progress appears in the log area, ending with the full output path.
4. Click **About** for product name, version, description, disclaimer, notes and feedback contact.

> The UI is fully Unicode (UTF-16 resources, Segoe UI); no garbled text on any Windows locale.

## 3. Fields reference
| Field | Meaning | Stored in | Default |
|---|---|---|---|
| ASP program zip | Site source; auto-extracted, top-prefix stripped, repacked in store mode, **appended to the exe tail** | self-attached segment (app.zip) | — |
| Icon ico (optional) | Icon of the output exe | resource 101 (icon group) | built-in icon |
| Title | Window / tray name | app.ini `Title` | My ASP App |
| Deny folders | Requests under these paths get 403, `\|` separated | app.ini `DenyDirs` | /data |
| Deny file extensions | Direct downloads of these extensions get 403, `\|` separated | app.ini `DenyExts` | .mdb |
| Local-only paths | Only reachable from 127.x, other IPs get 403, `\|` separated | app.ini `LocalOnly` | install.asp\|repass.asp |
| Data dirs | Files whose path matches are **extracted to disk for read/write**, `\|` separated, `-` = none | app.ini `DataDirs` | /data/\|/upload/ |
| Cache extensions | Files with these extensions are **extracted to disk for read/write**, `\|` separated, `-` = none | app.ini `CacheExts` | .json\|.ini |
| Runtime folder | Matched data is extracted here, next to the exe | app.ini `Root` | wwwroot |
| Output folder | Where the exe is written; empty = zip's folder | — | zip's folder |

> Priority layers: **compiled defaults < embedded config (app.ini) < on-disk aspfox.ini** (runtime override).

## 4. Behavior of the output exe
- **Program code and config live in a self-attached segment at the end of the exe** (last 20 bytes are the marker). No Windows resource-injection API involved — reliable even when security software blocks resource modification.
- **Size self-check**: the log prints `Output size: N bytes (shell X + program data Y + config Z + marker 20)`. Output ≈ shell + your program size; far smaller means the wrong (tiny) zip was selected.
- **Fool-proof**: selecting the built-in placeholder sample package is rejected with a clear error instead of producing a placeholder exe.
- HTTP listens on `0.0.0.0` (all NICs); ports tried in order `80 → 8080 → 8000 → 8888`.
- **All files in the zip go into the exe**; at runtime they split into two classes:
  - **Memory-only (never on disk)**: `.asp` / `.asa` / `assets/*` / `manual.html`, plus **every file not matched by the disk rules** (images, css, js, html...) — invisible, copy-protected, served straight from memory;
  - **Memory + disk (read/write)**: files matching "Data dirs" (e.g. `data/`, `upload/`) or "Cache extensions" (e.g. `.json`, `.ini`) are extracted to `<Runtime folder>\` on first run (multi-level sub-dirs auto-created; existing files are never overwritten). Afterwards **the on-disk copy wins over the embedded one** — editing the file on disk takes effect immediately.
- ⚠️ **The folder holding your Access database (.mdb/.accdb) must be listed in "Data dirs"**, otherwise the database never lands on disk and Jet cannot open it (the default `/data/` covers the common `data/wage.mdb` layout; add e.g. `/db/` if your DB lives elsewhere).
- Tray icon right-click menu: About / Settings (port, root, deny rules) / Restart / Exit / **Open LAN access (firewall)**.
- Public access needs port forwarding on your router; 360 / Huorong / PC Manager-style security suites have their own firewall and need the program allowed separately.
- `.mdb` and files starting with `data` are not directly downloadable (403) by default, but remain visible and read/writable for the program on disk.

## 5. Known limitations
- Both the packer and the packed exe run on **Windows** only; the build environment here has no Windows, so **no real-machine run test was performed** — please verify double-click and LAN access on the target machine.
- Prefer ASCII / UTF-8 file names inside the zip; file dialogs are Unicode, very old ANSI paths may misbehave.
- The embedded app.zip is stored uncompressed (invisible but not encrypted).
- If a binary resource inside an ASP page happens to contain the `PK\x03\x04` signature, server scanning may misjudge it (extremely rare; standard ASP sites are unaffected).
- `.asp/.asa/assets/*` are never extracted to disk (fixed rule); whether any other file lands on disk is decided by "Data dirs" and "Cache extensions".
- Note: the tray menu / settings dialog of the **packed exe** (server shell) is currently still Chinese-only; the English build covers the packer UI and documents. A fully-English shell can be provided on request.
