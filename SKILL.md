---
slug: fmg-autojy-video
version: 1.0.0
displayName: PPT课件转剪映微课
agent_created: true
---

# fmg-ppt-jianying-microlesson

把一份 PPT 课件 + 一段导入视频，自动生成一节带 AI 配音、字幕的剪映微课草稿，并导出 MP4。

## 触发场景

- 用户上传或引用 `.pptx` 课件 +（可选）导入视频，要求“做成微课视频/录课/精品课片段”。
- 需要把课件转成 1920×1080 视频，并配上教师旁白字幕。
- 需要同时交付：① 可直接播放的 MP4  ② 可在剪映专业版里继续编辑的草稿工程。

## 前置依赖

- Windows + 已安装 **剪映专业版**（v5.9+，实测 v10.9 写草稿可用）。
- Python 3.10+，managed venv 已安装 `jianying-editor-skill` 的依赖（含 `edge-tts`、`pywin32`、`python-pptx`）。
- ffmpeg 在 PATH 上（推荐 `imageio-ffmpeg` 提供的完整静态版，或自行安装 gyan.dev 版）。
- Microsoft PowerPoint 或 WPS（用于把 PPT 页面渲染成高清 PNG）。
- 已安装 `jianying-editor-skill` 并能在 `~/.workbuddy/skills/jianying-editor` 或 `H:/WorkBuddy/skill/jianying-editor-skill` 找到。

## 核心流程

1. **解析 PPT**
   - 用 `python-pptx` 读取每页文本、图片数，判断哪些页是真实内容（排除 Lorem ipsum / 通用模板尾页）。
   - 默认取前 N 页，或让用户指定页码范围。
2. **渲染课件图片**
   - 用 `win32com.client` 调用 PowerPoint COM，把每页导出为 `1920×1080` PNG。
   - 注意：PowerPoint COM 要求 `Visible=True`，不能设 `Visible=False`。
3. **撰写/加载旁白**
   - 每页需要一段自然的中文教师旁白，内容与页面一致。
   - 旁白以 Python 列表形式写在脚本头部，便于按课件定制。
4. **AI 配音**
   - 使用 `edge-tts`（`zh-CN-XiaoxiaoNeural` 等微软神经网络音色）生成 `.mp3`。
   - 备份方案：`jianying-editor-skill` 的 `add_narrated_subtitles` 会优先尝试剪映云 SAMI，失败再回退 edge-tts。
5. **合成单页视频**
   - 每页 = 静态 PNG 画面 + 对应配音，画面停留 `max(最低时长, 配音时长 + 0.5s)`。
6. **生成字幕 SRT**
   - 按导入视频时长 + 单页时长累计，生成绝对时间轴 SRT，用微软雅黑白字黑边。
7. **生成剪映草稿**
   - 用 `jy_wrapper.JyProject(width=1920, height=1080)` 新建工程。
   - 依次导入导入视频、15 页单页视频，并添加字幕文本段。
   - **关键**：每两段之间留 0.1s 间隙，避免微秒边界触发 `SegmentOverlap`。
8. **导出 MP4**
   - ffmpeg concat 拼接导入视频 + 单页视频，并烧录 SRT。
   - 输出 1920×1080、H.264+AAC、30fps。

## 主要产物

| 文件 | 说明 |
|------|------|
| `<outdir>/宝葫芦的秘密_微课.mp4` | 最终可播放视频 |
| `<workdir>/slides/slide_NN.png` | 每页课件高清图 |
| `<workdir>/tts_NN.mp3` | 每页 AI 配音 |
| `<workdir>/slide_NN.mp4` | 每页画面+配音片段 |
| `<workdir>/subtitles.srt` | 全程字幕 |
| 剪映草稿目录 | `C:\Users\<user>\AppData\Local\JianyingPro\User Data\Projects\com.lveditor.draft\<project_name>` |

## 使用方法

1. 复制 `references/build_microlesson.py` 到项目工作目录。
2. 修改脚本顶部的路径、PPT 页范围、每页旁白文案、音色选择。
3. 运行：
   ```bash
   python build_microlesson.py
   ```
4. 脚本结束后，MP4 生成到指定输出目录，同时剪映草稿已保存。

## 关键坑点

- **PNG 不能直接导入剪映草稿**：`add_media_safe` 走 `_add_video_safe`，对静态图解析容易失败。应先把每页合成带配音的 MP4 再导入。
- **剪映 headless 导出不可用**：v10.9 + WorkBuddy 沙箱无 GUI，`auto_exporter.py` 的 uiautomation 导出走不通；必须 ffmpeg 直接合成 MP4。
- **SRT 路径不能带 Windows 盘符**：ffmpeg `subtitles` filter 用 `:` 作选项分隔符，`I:\...` 会被解析错；在脚本工作目录下用相对路径 `subtitles=subtitles.srt`。
- **PowerPoint COM 可见性**：必须 `Visible=True`，否则报错 `Hiding the application window is not allowed`。
- ** WorkBuddy 沙箱删除限制**：第二次 `JyProject(..., overwrite=True)` 可能因沙箱拦截 `shutil.rmtree` 失败；用带时间戳的唯一工程名避免覆盖。

## 参考

- `references/build_microlesson.py`：本次《宝葫芦的秘密》微课完整可运行脚本（含真实旁白与页范围，可按需修改）。
