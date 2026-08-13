# fmg-autoJY-video（调用剪映自动剪辑视频）

分析所有素材（含PPT课件），自动调用剪映生成带 **AI 配音、字幕** 的剪映微课草稿，并导出可播放的 **MP4**。

> 原名 `fmg-ppt-jianying-microlesson`，在 SkillHub / GitHub 统一命名为 `fmg-autoJY-video`。

## 功能

- PowerPoint COM 把课件每页渲染为 1920×1080 PNG
- `edge-tts` 神经网络音色（默认 `zh-CN-XiaoxiaoNeural`）生成中文教师旁白
- ffmpeg 把「每页 PNG + 对应配音」合成为单页视频，并生成全程 SRT 字幕
- 生成**剪映专业版可编辑草稿**（可在剪映内继续剪辑）
- 直接 ffmpeg 导出可播放 MP4（绕过剪映 headless 导出限制，保证有成品文件）

## 前置依赖

- Windows + 已安装**剪映专业版**（v5.9+，实测 v10.9 写草稿可用）
- Python 3.10+，依赖：`edge-tts` / `pywin32` / `python-pptx`
- 已安装 `jianying-editor-skill`（提供 `JyProject` 草稿封装）
- ffmpeg 在 PATH 上（推荐 `imageio-ffmpeg` 提供的完整静态版）
- Microsoft PowerPoint 或 WPS（用于把 PPT 页面渲染成高清 PNG）

## 使用方法

1. 复制 `references/build_microlesson.py` 到你的项目工作目录
2. 修改脚本顶部的：输出目录、PPT 路径、页范围、每页旁白文案、AI 音色
3. 运行：

   ```bash
   python build_microlesson.py
   ```

4. 脚本结束后，MP4 生成到指定输出目录，同时剪映草稿已保存到本地

## 文件结构

| 文件 | 说明 |
|------|------|
| `SKILL.md` | 技能说明（触发场景 / 流程 / 坑点） |
| `references/build_microlesson.py` | 完整可运行脚本（含《宝葫芦的秘密》真实旁白示例） |

## 关键坑点

- 剪映 v10.9 + 无 GUI 环境下 headless 导出不可用 → 用「剪映草稿 + ffmpeg 直出 MP4」混合方案
- PowerPoint COM 必须 `Visible=True`，否则 `Hidden application window not allowed`
- ffmpeg `subtitles` 滤镜的 SRT 路径不能含 Windows 盘符（`:` 冲突），用相对路径
- 多段视频连续加入时间轴时，浮点秒→微秒会产生边界重叠，需每段间留 0.1s 间隙

## License

MIT — 详见 [LICENSE](LICENSE)。
