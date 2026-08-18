# fmg-autoJY-video（PPT 课件转剪映微课 · 全版本兼容）

把一份 PPT 课件 + 一段导入视频，自动生成一节带 **AI 配音、字幕** 的剪映微课草稿，并导出可播放的 **MP4**。

> 已验证支持 **剪映专业版 v5.9 / v10.9 / v11.2**（及之间所有版本）。

## 功能

- PowerPoint COM 把课件每页渲染为 1920×1080 PNG
- `edge-tts` 神经网络音色（默认 `zh-CN-XiaoxiaoNeural`）生成中文教师旁白
- ffmpeg 把「每页 PNG + 对应配音」合成为单页视频，并生成全程 SRT 字幕（按标点切短句、单行、与配音严格同步）
- **版本感知的剪映草稿**：`jy_draft` 自动探测本机剪映版本，写出该版本兼容的可编辑工程（v10.9 / v11.2 采用同一套现代格式，直接打开；旧版打开自动升级）。**纯标准库实现，不依赖 pyJianYingDraft**
- 直接 ffmpeg 导出可播放 MP4（**完全不调用剪映**，故导出与剪映版本无关）

## 全版本兼容说明

| 产物 | 是否依赖剪映版本 | 说明 |
|------|------------------|------|
| **MP4（主交付物）** | 否 | 全程 ffmpeg 合成，v5.9 / v10.9 / v11.2 及未安装剪映都能导出 |
| **剪映草稿** | 自动适配 | `jy_draft` 探测版本后写对应兼容草稿；v10.9/v11.2 直接打开，旧版自动升级 |

> 剪映 v10.9+ 关闭了 headless 自动导出（防自动化），因此脚本只生成**可编辑草稿**；
> 最终点「导出」需在剪映 GUI 手动操作，或直接采用本 skill 的 ffmpeg MP4。两种产物时间轴完全一致。

## 前置依赖

- Windows（剪映草稿为通用工程格式；MP4 导出不依赖剪映）
- 已安装剪映专业版 v5.9 / v10.9 / v11.2 任一版本均可（不装也能导出 MP4，只是没有可编辑草稿）
- Python 3.10+，依赖：`edge-tts` / `pywin32` / `python-pptx` / `requests`
- ffmpeg 在 PATH 上（推荐 `imageio-ffmpeg` 提供的完整静态版）
- Microsoft PowerPoint 或 WPS（用于把 PPT 页面渲染成高清 PNG）

## 使用方法

1. 复制 `references/build_microlesson.py` 到你的项目工作目录
2. 修改脚本顶部的：输出目录、PPT 路径、页范围、每页旁白文案、AI 音色
3. 运行：

   ```bash
   python build_microlesson.py
   ```

4. 脚本结束后，MP4 生成到指定输出目录，同时剪映草稿已保存到本地（可用对应版本剪映打开）

## 文件结构

| 文件 | 说明 |
|------|------|
| `SKILL.md` | 技能说明（触发场景 / 流程 / 坑点） |
| `references/build_microlesson.py` | 完整可运行脚本（含《宝葫芦的秘密》真实旁白示例） |
| `references/jy_draft.py` | 版本感知的剪映草稿生成器（探测版本 + 基于真实 v11.2 模板写出兼容工程） |
| `references/draft_template_v11.json` | 从本机剪映 v11.2 真实草稿提取的静态骨架模板 |
| `references/draft_meta_template_v11.json` | 配套的 `draft_meta_info.json` 模板 |

## 关键坑点

- **MP4 导出与剪映版本无关**：全程 ffmpeg，不调用剪映，全版本稳定导出
- **草稿基于真实 v11.2 模板**：避免 pyJianYingDraft 版本错配导致草稿损坏；`jy_draft` 仅动态填充视频 / 字幕片段
- PowerPoint COM 必须 `Visible=True`，否则 `Hidden application window not allowed`
- ffmpeg `subtitles` 滤镜的 SRT 路径不能含 Windows 盘符（`:` 冲突），用相对路径
- 版本探测失败会回退默认 `11.2.0` 写草稿（现代格式），旧版剪映打开时会提示升级，同样可用

## License

MIT — 详见 [LICENSE](LICENSE)。
