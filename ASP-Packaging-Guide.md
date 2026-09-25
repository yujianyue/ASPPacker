# ASP Program Zip — Packaging Guide (aspfox Packer Edition)

> Audience: preparing an ASP program (e.g. a wage-inquiry system built on classic ASP + Access) for packing with the "ASP Single-File EXE Packer".
> Goal: turn an existing ASP site into one EXE — double-click to serve over local / intranet / public IP+port.

---

## 1. Zip structure requirements (most important)

### 1.1 Entry file
- The site root must contain **`index.asp`**. Without it visitors get 404.
- Login pages, install pages etc. keep their original relative paths.

### 1.2 Folder layout (both supported)
```
Layout A (recommended): files at zip root
  index.asp
  login.asp
  data/wage.mdb
  inc/conn.asp

Layout B: everything inside one top-level folder (prefix stripped automatically)
  MyWageSystem/
    index.asp
    login.asp
    data/wage.mdb
```
> Layout B is equivalent to A — the single common top-level prefix is stripped automatically. Stripping only happens when ALL entries share one top folder; if root-level files exist alongside folders, nothing is stripped.

### 1.3 Compression
- **zip format** (store or deflate). The packer re-packs everything into store mode inside the exe.
- **Everything goes into the exe**; at runtime files split into two classes:
  - **Memory-only (never on disk)**: `.asp`, `.asa`, `assets/*` (source protection), plus **every file not matched by the disk rules** (images, css, js, html...).
  - **Memory + disk (read/write)**: files matching **Data dirs** (default `/data/|/upload/`, path-segment match) or **Cache extensions** (default `.json|.ini`, extension match) are extracted to `exe-adjacent wwwroot\` on first run (existing files never overwritten, multi-level sub-dirs auto-created). Afterwards **the on-disk copy wins over the embedded one**.
- **⚠️ The folder holding your Access database must be listed in the packer's "Data dirs"** (add `/db/` if your DB lives in `db/`), otherwise the DB never lands on disk and Jet cannot open it. The default `/data/` covers layouts like `data/wage.mdb`.

### 1.4 File names
- **Chinese file names** and folders are supported (e.g. `help/说明.htm`).
- Avoid `% " * : < > ? / \ |` in names.

---

## 2. Feature support matrix

Script engines: **VBScript** (default) + **JScript** (per-page `<%@ Language="JScript" %>`); `global.asa` always runs as VBScript. 32-bit process, Windows 7+ built-in components.

### Response (fully supported)
| Member | Notes |
|---|---|
| Write / WriteLine | Output |
| BinaryWrite | Binary output (captchas, downloads) |
| Redirect / End / Clear / Flush | Redirect / stop / clear / flush |
| ContentType / Charset / Status | Page type, charset, status code |
| Buffer / Expires / ExpiresAbsolute / CacheControl / Pics / AppendToLog | Caching & misc |
| AddHeader | Custom response headers |
| Cookies (Expires / Domain / Path / Secure / HttpOnly / HasKeys) | Writing cookies |
| CodePage / LCID | Code page & locale |
| IsClientConnected | Client connection state |

### Request (fully supported)
| Member | Notes |
|---|---|
| QueryString / Form / Cookies | GET / POST / Cookie (HasKeys & multi-value supported) |
| ServerVariables | 23 common keys: REMOTE_ADDR, SCRIPT_NAME, QUERY_STRING, CONTENT_TYPE, CONTENT_LENGTH, HTTP_VERSION, SERVER_NAME, SERVER_PORT, SERVER_PROTOCOL, SERVER_SOFTWARE, REQUEST_METHOD, PATH_INFO, PATH_TRANSLATED, URL, LOCAL_ADDR, HTTPS, INSTANCE_ID, GATEWAY_INTERFACE, APPL_MD_PATH, APPL_PHYSICAL_PATH, SERVER_PORT_SECURE, REMOTE_HOST, REQUEST_URI |
| TotalBytes / BinaryRead | Raw request body (the basis for binary upload parsing) |
| ClientCertificate / Browser | Certificate collection / browser object |

### Server
| Member | Notes |
|---|---|
| CreateObject(progID) | **Any system-registered COM**: ADODB.Connection/Recordset/Stream, Scripting.FileSystemObject, MSXML2.DOMDocument, WScript.Shell work out of the box; third-party components must be registered on the target machine first |
| MapPath | Relative or `/`-rooted paths → physical paths |
| HTMLEncode / URLEncode | Encoding |
| Execute / Transfer | In-page execution / hand-off (IIS 5.0 semantics) |
| GetLastError | ASPError object (Number/Description/Source/File/Line/Column/ASPCode/ASPDescription/Category) |
| ScriptTimeout | Script timeout |

### Session / Application
- `Session`: SessionID, Timeout, Abandon, Contents, StaticObjects, CodePage, LCID; `For Each` enumeration, Count, Remove, RemoveAll.
- `Application`: Lock, UnLock, Contents, StaticObjects; same enumeration support.
- `global.asa`: `Application_OnStart`, `Session_OnStart` supported; `global.asa` and any `.asa` file itself always returns 403 (same as IIS).
- **Includes**: `<!--#include file="inc/conn.asp"-->` and `<!--#include virtual="/inc/conn.asp"-->` are supported.

### Database (Access)
- `Server.CreateObject("ADODB.Connection")` over **Jet OLEDB (32-bit)** — `.mdb` works out of the box, no Office needed.
- Recommended connection string (MapPath, never hard-code drive letters):
  ```asp
  Set conn = Server.CreateObject("ADODB.Connection")
  conn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath("data/wage.mdb")
  ```
- `data/*.mdb` is blocked from direct HTTP download by default (see security fields).

---

## 3. Encoding (CodePage) essentials

1. Preferred: declare per page — `<%@ CodePage=936 %>` (legacy GB2312/GBK) or `<%@ CodePage=65001 %>` (UTF-8).
2. Or set `DefaultCodePage` (936 / 65001) in the runtime ini, leaving pages undeclared.
3. Without any declaration the **system ANSI code page** applies (936 on Simplified Chinese systems).
4. `Response.Charset` recognizes: utf-8, gb2312/gbk/gb18030, big5, shift_jis, euc-kr, iso-8859-1, windows-1252.
5. **Keep one encoding per site**: file encoding, CodePage and `meta charset` must agree, otherwise expect mojibake.

---

## 4. Packer fields (app.ini mapping)

All fields are embedded into the exe; a run-folder `aspfox.ini` may override them later:

| Field | Default | Meaning |
|---|---|---|
| Title | My ASP App | Tray tip / browser title / output file name |
| Deny folders | /data | `\|` separated, e.g. `/admin\|/data` |
| Deny extensions | .mdb | `\|` separated |
| Deny filename prefix | data | `data` → blocks `data*` |
| Data dirs | /data/\|/upload/ | Matched paths extracted to disk for read/write, `-` = none |
| Cache extensions | .json\|.ini | Matched extensions extracted to disk for read/write, `-` = none |
| Local-only paths | install.asp\|repass.asp | Install / password pages only from this machine |
| Runtime folder | wwwroot | Data extracted here (next to the exe); code never lands on disk |

> Priority: **compiled < embedded app.ini < on-disk aspfox.ini**. Restart via the tray menu after editing.

---

## 5. Not supported / limited

| Item | Status | Suggestion |
|---|---|---|
| Unregistered 3rd-party COM | CreateObject fails | Ship and `regsvr32` it on the target machine |
| .NET components (COM interop) | Limited | Prefer plain COM or rewrite in native ASP |
| IIS-only components (MSWC tools etc.) | Partially missing | Avoid; the core trio + ADO are complete |
| Encrypted / split zips | Not supported | Pack as plain, unencrypted zip |
| Non-zip formats (rar/7z) | Not supported | Convert to zip first |
| Access `.accdb` | Needs 32-bit Access Database Engine on the target | Prefer `.mdb` (Jet 4.0, most stable) |
| Very large uploads | Memory-bound | Keep big attachments on disk via classic forms |
| Multi-site / virtual dirs | Single site | Pack multiple EXEs instead |

---

## 6. Pre-flight checklist

- [ ] **index.asp** at zip root (or inside the single top-level folder)
- [ ] Plain **unencrypted zip**
- [ ] **DB folder listed in "Data dirs"** (default /data/; add e.g. /db/ if needed)
- [ ] Runtime-written files (uploads, caches) match "Data dirs" or "Cache extensions"
- [ ] DB opened via `Server.MapPath` relative path (no `C:\...` hard-coded letters)
- [ ] Page CodePage / file encoding / meta charset agree
- [ ] Included files (`inc/conn.asp` etc.) referenced with correct paths
- [ ] No IIS-only components; required 3rd-party COM registrable on the target
- [ ] Sensitive folders (e.g. `/data`) configured in "Deny folders"
- [ ] Output folder empty = output written next to the source zip
- [ ] After packing, the log shows `Output size ≈ shell (180KB) + your program size`; far smaller means the wrong zip was selected

---

## 7. FAQ

**Q: The packed program shows a placeholder page or 404 everywhere?**
Placeholder = the built-in sample package was selected by mistake (the packer now rejects it outright); 404 = legacy bug in embedded Root handling (fixed in v1.1). Re-pack with the current build.

**Q: The packer crashes right after generating?**
v1.0 had a zip central-directory sizing bug (heap overflow with many files, crashing at the end of packing). Fixed in v1.1 and verified with AddressSanitizer (old formula always overflows, new one is clean).

**Q: The program crashed — how to investigate?**
Since v1.1 a crash writes `<ProgramName>.crash.log` (exception code + address) next to the exe. Send that file's content to the developer for diagnosis.

**Q: The packed program serves 404 everywhere?**
Make sure `index.asp` is at the zip root; a single top-level folder is stripped automatically. Still 404 → you are on an old build (embedded Root bug); re-pack with v1.1+.

**Q: Pages show mojibake?**
See "Encoding essentials": CodePage must match the real file encoding; legacy systems use 936, new programs 65001.

**Q: Where is the Access database at runtime?**
Files matched by "Data dirs" are extracted under `wwwroot\` next to the exe (e.g. `wwwroot\data\wage.mdb`) on first run — open and maintain them with Access directly; changes apply immediately, no repack needed.

**Q: Runtime error "cannot open database / file not found"?**
The file never landed on disk — by default only `data/`, `upload/` folders and `.json/.ini` extensions are extracted. Add the DB folder to "Data dirs" in the packer and re-pack.

---

*aspfox / ASP Single-File EXE Packer v1.1 (English build) · Feedback: yujianyue 15058593138@qq.com*
