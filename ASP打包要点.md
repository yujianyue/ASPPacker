# ASP 程序压缩包制作要点（aspfox 通用打包器适配指南）

> 适用对象：准备用「ASP 单文件 EXE 通用打包器」打包的 ASP 程序（如工资查询系统等 Access + ASP 老系统）。
> 目标：把现有 ASP 站点打成一个 EXE，双击即用（本机 / 内网 / 公网 IP+端口 访问）。

---

## 一、压缩包结构要求（最重要）

### 1. 入口文件
- 站点根目录必须有 **`index.asp`**（首页入口）。没有会显示 404。
- 登录页、安装页等按原有相对路径存放即可。

### 2. 目录形态（两种都支持）
```
形态 A（推荐）：文件直接在 zip 根
  index.asp
  login.asp
  data/wage.mdb
  inc/conn.asp

形态 B：整体包在一个顶层文件夹里（打包器自动剥离该前缀）
  工资查询系统/
    index.asp
    login.asp
    data/wage.mdb
```
> 形态 B 会自动把 `工资查询系统/` 这层剥掉，等价于形态 A。**只有一层共同顶层目录时才剥离**；zip 里既有根级文件又有文件夹时不动。

### 3. 压缩格式
- **zip 格式**（store 或 deflate 均可）。打包器会自动解压并转为 store（无压缩）模式内嵌进 EXE，无需手工处理。
- **zip 里所有文件都会打包进 exe**，但分两类运行：
  - **纯内存（不落盘）**：`.asp`、`.asa`、`assets/*` 源码类固定纯内存；此外**未命中落盘规则的文件**（图片、css、js、htm 等）也只存于 exe，请求时从内存直接输出。
  - **内存 + 落盘（可读写）**：命中「动态数据目录」（默认 `/data/|/upload/`）或「缓存后缀」（默认 `.json|.ini`）的文件，首次运行自动释放到 EXE 同级 `wwwroot\` 目录（已存在不覆盖，多级子目录自动创建），之后**磁盘版优先于内存版**——改磁盘文件即生效。
- **⚠️ Access 数据库所在目录必须列入打包器的「动态数据目录」**（如库在 `db/` 就把 `/db/` 加进去），否则数据库不落盘、Jet 打不开会报错。默认 `/data/` 已覆盖 `data/wage.mdb` 这类常见布局。

### 4. 文件名
- 支持**中文文件名**与中文目录（如 `help/说明.htm`）。
- 建议避免 `%` `" * : < > ? / \ |` 等字符。

---

## 二、支持情况总表

脚本引擎：**VBScript**（默认）+ **JScript**（`<%@ Language="JScript" %>` 按页支持）；`global.asa` 固定按 VBScript 执行。32 位进程，Windows 7 及以上自带组件。

### Response（全部支持）
| 成员 | 说明 |
|---|---|
| Write / WriteLine | 输出 |
| BinaryWrite | 二进制输出（验证码、下载等） |
| Redirect / End / Clear / Flush | 跳转 / 终止 / 清空 / 冲刷 |
| ContentType / Charset / Status | 页面类型、字符集、状态码 |
| Buffer / Expires / ExpiresAbsolute / CacheControl / Pics / AppendToLog | 缓存与杂项 |
| AddHeader | 自定义响应头 |
| Cookies（含 Expires / Domain / Path / Secure / HttpOnly / HasKeys） | 写 Cookie |
| CodePage / LCID | 代码页与区域设置 |
| IsClientConnected | 客户端连接状态 |

### Request（全部支持）
| 成员 | 说明 |
|---|---|
| QueryString / Form / Cookies | GET / POST / Cookie（Cookies 支持 HasKeys 与多值） |
| ServerVariables | 23 个常用键：REMOTE_ADDR、SCRIPT_NAME、QUERY_STRING、CONTENT_TYPE、CONTENT_LENGTH、HTTP_VERSION、SERVER_NAME、SERVER_PORT、SERVER_PROTOCOL、SERVER_SOFTWARE、REQUEST_METHOD、PATH_INFO、PATH_TRANSLATED、URL、LOCAL_ADDR、HTTPS、INSTANCE_ID、GATEWAY_INTERFACE、APPL_MD_PATH、APPL_PHYSICAL_PATH、SERVER_PORT_SECURE、REMOTE_HOST、REQUEST_URI |
| TotalBytes / BinaryRead | 原始请求体（文件上传按二进制流解析的基础） |
| ClientCertificate / Browser | 证书集合 / 浏览器对象 |

### Server
| 成员 | 说明 |
|---|---|
| CreateObject(progID) | **任意系统已注册 COM**：ADODB.Connection/Recordset/Stream、Scripting.FileSystemObject、MSXML2.DOMDocument、WScript.Shell 等开箱即用；第三方组件需先在目标机器注册 |
| MapPath | 相对路径 / `/` 开头 → 站点根映射物理路径 |
| HTMLEncode / URLEncode | 编码 |
| Execute / Transfer | 页面内嵌执行 / 转交（IIS 5.0 语义） |
| GetLastError | 返回 ASPError 对象（Number/Description/Source/File/Line/Column/ASPCode/ASPDescription/Category） |
| ScriptTimeout | 脚本超时 |

### Session / Application
- `Session`：SessionID、Timeout、Abandon、Contents、StaticObjects、CodePage、LCID；支持 `For Each` 枚举、Count、Remove、RemoveAll。
- `Application`：Lock、UnLock、Contents、StaticObjects；同样支持枚举与 Count。
- `global.asa`：支持 `Application_OnStart`、`Session_OnStart` 事件；`global.asa` / 任何 `.asa` 文件本身始终 403 禁止直接访问（与 IIS 一致）。
- **模板与包含**：支持 `<!--#include file="inc/conn.asp"-->` 与 `<!--#include virtual="/inc/conn.asp"-->`。

### 数据库（Access）
- 通过 `Server.CreateObject("ADODB.Connection")` 走 **Jet OLEDB（32 位）**，`.mdb` 开箱即用，无需安装 Office。
- 连接串写法（推荐 MapPath 拼路径，不要写死盘符）：
  ```asp
  Set conn = Server.CreateObject("ADODB.Connection")
  conn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath("data/wage.mdb")
  ```
- `data/*.mdb` 默认禁止 HTTP 直接下载（见下节安全参数）。

---

## 三、编码（CodePage）要点

1. 首选：每个页面顶部写 `<%@ CodePage=936 %>`（GB2312/GBK 老程序）或 `<%@ CodePage=65001 %>`（UTF-8 程序）。
2. 或在打包器里留空、由运行时 ini 统一指定 `DefaultCodePage`（936 / 65001）。
3. 未指定时默认**系统 ANSI 代码页**（简体中文系统 = 936）。
4. `Response.Charset` 支持：utf-8、gb2312/gbk/gb18030、big5、shift_jis、euc-kr、iso-8859-1、windows-1252。
5. **同一站点请保持编码统一**：页面文件编码、CodePage、`meta charset` 三者一致，否则乱码。

---

## 四、打包器参数（app.ini 映射）

打包界面各参数写入 EXE 内嵌配置，运行时磁盘 `aspfox.ini` 可再覆盖：

| 参数 | 默认值 | 说明 |
|---|---|---|
| 软件标题 | 我的ASP应用 | 托盘提示 / 浏览器标题 / 输出文件名 |
| 不可访问文件夹 | /data | 多个用 `\|` 分隔，如 `/admin\|/data` |
| 禁止下载后缀 | .mdb | 多个用 `\|` 分隔 |
| 禁止文件名前缀 | data | `data` → 拦截 `data*` |
| 动态数据目录 | /data/\|/upload/ | 命中路径段的文件落盘供读写，多个用 `\|` 分隔，填 `-` 表示无 |
| 缓存后缀 | .json\|.ini | 命中扩展名的文件落盘供读写，多个用 `\|` 分隔，填 `-` 表示无 |
| 仅本机可访问路径 | install.asp\|repass.asp | 安装/改密页只允许本机打开 |
| 网站根目录 | wwwroot | 数据释放目录（EXE 同级），代码不落盘 |

> 三层优先级：**编译内置 < EXE 内嵌配置（打包器写入）< 运行目录 aspfox.ini**。改配置后用托盘菜单「重启」生效。

---

## 五、不支持 / 受限清单

| 项目 | 情况 | 处理建议 |
|---|---|---|
| 第三方未注册 COM | CreateObject 失败 | 随程序分发并 `regsvr32` 注册后再用 |
| .NET 组件（COM 互操作） | 受限 | 改用纯 COM 或改写为 ASP 原生逻辑 |
| IIS 专属组件（MSWC 工具类等） | 部分缺失 | 少用；核心三对象+ADO 全齐 |
| zip 加密 / 分卷 | 不支持 | 打包时选「普通/无加密」 |
| 除 zip 外的压缩格式（rar/7z） | 不支持 | 先转 zip |
| Access `.accdb` | 需目标机装有 Access Database Engine（32 位） | 建议转存为 `.mdb`（Jet 4.0 最稳） |
| 超大文件上传 | 受内存限制 | 大附件建议走磁盘目录 + 传统表单 |
| 多站点/虚拟目录 | 单站点 | 打多个 EXE 即可 |

---

## 六、打包前自检清单

- [ ] zip 根（或唯一顶层目录内）有 **index.asp**
- [ ] zip 为**普通 zip 格式**、未加密
- [ ] **数据库所在目录已列入「动态数据目录」**（默认 /data/，库在别的目录如 db/ 需自行添加）
- [ ] 运行时需要读写的文件（上传目录、缓存文件）已命中「动态数据目录」或「缓存后缀」
- [ ] 数据库用 `Server.MapPath` 相对路径连接（无 `C:\...` 死盘符）
- [ ] 页面 CodePage / 文件编码 / meta charset 三者一致
- [ ] `inc/conn.asp` 等被包含文件路径大小写与实际一致（服务器不区分大小写，但建议规范）
- [ ] 不依赖 IIS 专属组件；用到的第三方 COM 目标机可注册
- [ ] 敏感目录（如 `/data`）已在打包器「不可访问文件夹」里配置
- [ ] 打包器「输出目录」留空 = 输出到源码 zip 所在目录
- [ ] 打包完成后看日志「输出大小 ≈ 壳 178KB + 你的程序大小」，明显偏小说明选错 zip

---

## 七、常见问题

**问：之前生成的程序打开后一直显示"占位提示"或 404？**
占位提示 = 选错了打包器自带的示例包（现已内置防呆检测直接报错）；404 = 旧版内嵌 Root 相对路径解析缺陷（v1.1 已修复）。请使用新版打包器重新生成。

**问：打包器点击生成后闪退？**
v1.0 存在 zip 中央目录容量计算缺陷（文件较多时堆溢出，生成完成后崩溃）。v1.1 已修复，并用 AddressSanitizer 对照验证（旧公式必溢出、新公式干净）。

**问：程序闪退了怎么排查？**
v1.1 起崩溃不再静默：EXE 同级会生成 `<程序名>.crash.log`（含异常代码与地址），把该文件内容反馈给开发者即可定位。

**问：生成的程序运行后是 404？**
确认 zip 根有 `index.asp`；若把整个文件夹压进 zip 也没关系（自动剥顶层）；仍 404 请确认使用 v1.1 及以上打包器（旧版 Root 缺陷）。

**问：页面乱码？**
见「编码要点」：CodePage 与文件实际编码保持一致；老系统统一 936，新程序统一 65001。

**问：Access 数据库在哪改？**
命中「动态数据目录」的文件首次运行后在 EXE 同级 `wwwroot\` 下（如 `wwwroot\data\wage.mdb`），可直接用 Access 打开维护；改完即生效，无需重新打包。

**问：运行时报"数据库打不开 / 找不到文件"？**
该文件没落盘——默认只有 `data/`、`upload/` 目录和 `.json/.ini` 后缀落盘。把数据库所在目录加入打包器「动态数据目录」后重新打包即可。

---

*aspfox / 通用打包器 v1.1 · 问题反馈：yujianyue 15058593138@qq.com*
