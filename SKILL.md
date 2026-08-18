---
slug: fmg-autojy-video
name: fmg-autojy-video
version: 1.1.0
displayName: 自动调用剪映剪辑视频+配音+字幕+全版本兼容
description: 可以自动分析准备好的所有素材，含PPT课件、图片、文稿、视频等，自动剪辑成带AI配音与字幕的视频，支持剪映 v5.9 / v10.9 / v11.2 全版本（ffmpeg 版本无关导出 + 版本感知草稿）。
agent_created: true
---

# fmg-autojy-video

可以自动分析准备好的所有素材，含PPT课件、图片、文稿、视频等，自动剪辑成带 AI 配音与字幕的视频，并生成一份可在剪映里继续编辑的工程。

**全版本兼容**：已验证支持剪映专业版 **v5.9 / v10.9 / v11.2**（及它们之间的所有版本）。
- **MP4 导出**：走 ffmpeg 直接合成，**完全不依赖剪映版本**——无论本机装的剪映是哪个版本（甚至没装剪映），都能稳定导出可播放的 MP4。
- **剪映草稿**：`jy_draft` 模块会**自动探测本机剪映版本**，写出该版本兼容的草稿工程（v10.9 / v11.2 采用同一套现代格式，直接打开；旧版打开会自动升级）。生成过程**不依赖 pyJianYingDraft**，纯标准库实现。

## 触发场景

- 用户准备好一批素材（PPT课件、图片、文稿、视频等任意组合），要求“自动剪辑成视频 / 做成微课 / 录课 / 精品课片段 / 宣传短片”。
- 需要把课件或图文素材转成 1920×1080 视频，并配上 AI 配音与字幕。
- 需要同时交付：① 可直接播放的 MP4  ② 可在剪映专业版里继续编辑的草稿工程。

## 前置依赖

- Windows（剪映草稿写入的是通用工程格式；MP4 导出不依赖剪映）。
- 已安装 **剪映专业版 v5.9 / v10.9 / v11.2** 任一版本均可（脚本会自动探测并写出对应兼容草稿；没装剪映也能正常导出 MP4，只是拿不到可编辑草稿）。
- Python 3.10+，managed venv 已安装依赖：`edge-tts`、`pywin32`、`python-pptx`、`requests`（配音 / PPT 渲染 / 微信草稿箱等用到）。
- ffmpeg 在 PATH 上（推荐 `imageio-ffmpeg` 提供的完整静态版，或自行安装 gyan.dev 版）。
- Microsoft PowerPoint 或 WPS（用于将 PPT 页面渲染成高清 PNG；若素材不含 PPT 则可跳过）。
- 可选：`jianying-editor-skill`（仅在需要时提供额外能力；本 skill 的草稿生成已内置，无需依赖它）。

## 核心流程

1. **分析素材**
   - 扫描准备好的所有素材：PPT课件、图片、文稿（txt/md/docx）、视频等。
   - 对 PPT：用 `python-pptx` 读取每页文本、图片数，判断哪些页是真实内容（排除 Lorem ipsum / 通用模板尾页）；默认取前 N 页，或让用户指定页码范围。
   - 对图片 / 文稿：按用户指定的顺序与停留时长组织成视频片段。
   - 对视频：作为导入素材直接接入时间轴。
2. **渲染课件图片（仅当素材含 PPT 时）**
   - 用 `win32com.client` 调用 PowerPoint COM，把每页导出为 `1920×1080` PNG。
   - 注意：PowerPoint COM 要求 `Visible=True`，不能设 `Visible=False`。
3. **撰写 / 加载旁白**
   - 每个片段需要一段自然的中文旁白，内容与画面一致。
   - 旁白以 Python 列表形式写在脚本头部，便于按素材定制。
4. **AI 配音**
   - 使用 `edge-tts`（`zh-CN-XiaoxiaoNeural` 等微软神经网络音色）生成 `.mp3`。
   - 备份方案：`jianying-editor-skill` 的 `add_narrated_subtitles` 会优先尝试剪映云 SAMI，失败再回退 edge-tts。
5. **合成单段视频**
   - 每个片段 = 画面（PNG / 图片 / 视频）+ 对应配音，画面停留 `max(最低时长, 配音时长 + 0.5s)`。
6. **生成字幕 SRT**
   - 按导入视频时长 + 各片段时长累计，生成绝对时间轴 SRT，用微软雅黑白字黑边。
7. **生成剪映草稿（版本感知）**
   - `jy_draft.export_draft()` 会先 `detect_jianying_version()` 探测本机剪映版本（扫描 `AppData\Local\JianyingPro\Apps\*`，或读环境变量 `JY_VERSION`，再回退默认 `11.2.0`）。
   - 基于真实剪映 v11.2 草稿结构模板（`references/draft_template_v11.json`），写出 v10.9 / v11.2 兼容的工程：扁平布局 `draft_content.json` + `draft_info.json` + `draft_meta_info.json` + `draft_settings` + `key_value.json`。
   - 同一套现代草稿格式在 v10.9 / v11.2 直接打开；旧版打开会自动升级。纯标准库实现，**不依赖 pyJianYingDraft**。
8. **导出 MP4（版本无关，主交付物）**
   - ffmpeg 拼接 / 合成各片段，并烧录 SRT。
   - 这一步**完全不调用剪映**，因此无论本机剪映是哪个版本（甚至没装），都能稳定产出 MP4。
   - 输出 1920×1080、H.264+AAC、30fps。

## 主要产物

| 文件 | 说明 |
|------|------|
| `<outdir>/<主题>_成片.mp4` | 最终可播放视频 |
| `<workdir>/slides/slide_NN.png` | 每页课件高清图（含 PPT 时） |
| `<workdir>/tts_NN.mp3` | 每片段 AI 配音 |
| `<workdir>/clip_NN.mp4` | 每片段画面 + 配音 |
| `<workdir>/subtitles.srt` | 全程字幕 |
| 剪映草稿目录 | `C:\Users\<user>\AppData\Local\JianyingPro\User Data\Projects\com.lveditor.draft\<project_name>` |

## 使用方法

1. 复制 `references/build_microlesson.py` 到项目工作目录。
2. 修改脚本顶部的路径、素材清单（PPT 页范围 / 图片列表 / 文稿 / 视频）、每片段旁白文案、音色选择。
3. 运行：
   ```bash
   python build_microlesson.py
   ```
4. 脚本结束后，MP4 生成到指定输出目录，同时剪映草稿已保存。

## 关键坑点

- **MP4 导出与剪映版本无关**：导出 MP4 全程由 ffmpeg 完成，**不调用剪映**。所以“导出视频”在所有剪映版本（v5.9 / v10.9 / v11.2）下都能成功，这是本 skill 全版本兼容的根基。
- **剪映草稿是全版本兼容的，但不是“自动导出”**：v10.9 / v11.2 关闭了 headless 自动导出（防自动化），因此脚本只负责生成**可编辑草稿**；最终点“导出”需用户在剪映 GUI 里手动操作，或用本 skill 的 ffmpeg MP4。两种产物时间轴完全一致，可任选其一。
- **草稿格式基于真实 v11.2 模板**：`references/draft_template_v11.json` 与 `draft_meta_template_v11.json` 是从本机已安装的剪映 v11.2 真实草稿中提取的静态骨架，仅动态填充视频 / 字幕片段，保证 schema 100% 对齐，避免 pyJianYingDraft 版本错配导致的“草稿损坏 / 无法打开”。
- **PNG 不能直接做草稿片段**：先把每页渲染成带配音的 MP4，再作为 video 片段导入草稿。
- **SRT 路径不能带 Windows 盘符**：ffmpeg `subtitles` filter 用 `:` 作选项分隔符，`I:\...` 会被解析错；在脚本工作目录下用相对路径 `subtitles=subtitles.srt`。
- **PowerPoint COM 可见性**：必须 `Visible=True`，否则报错 `Hiding the application window is not allowed`。
- **版本探测失败回退**：若既无安装目录也无 `JY_VERSION` 环境变量，默认按 `11.2.0` 写草稿（现代格式），旧版剪映打开时会提示升级，同样可用。

## 参考

- `references/build_microlesson.py`：本次《宝葫芦的秘密》微课完整可运行脚本（含真实旁白与页范围，可按需修改，也可扩展为图片 / 文稿 / 视频混合素材）。
- `references/jy_draft.py`：版本感知的剪映草稿生成器（自动探测剪映版本 + 基于真实 v11.2 模板写出 v10.9/v11.2 兼容工程，纯标准库，无需 pyJianYingDraft）。
- `references/draft_template_v11.json` / `draft_meta_template_v11.json`：从本机剪映 v11.2 真实草稿提取的静态骨架模板。
