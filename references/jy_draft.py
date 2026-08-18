"""
jy_draft.py — 版本感知的剪映草稿生成器（自包含，无需 pyJianYingDraft）

核心能力：
1. detect_jianying_version()  —— 自动探测本机已安装的剪映专业版版本
                                 （扫描 AppData\\Local\\JianyingPro\\Apps\\*，
                                  返回最新版本号，如 "11.2.0.14339"；
                                  找不到则回退到环境变量 JY_VERSION，再回退 "11.2.0"）。
2. build_v11_draft()          —— 基于真实剪映 v11.2 草稿结构（已提取为模板），
                                 直接写出 v10.9 / v11.2 兼容的草稿工程。
                                 纯标准库实现，不依赖任何第三方草稿库。
3. export_draft()             —— 统一出口：写出版本对应的剪映草稿目录。

设计说明：
- 剪映 v10.9 / v11.2 采用同一套“现代草稿格式”（version=360000 / new_version=110.0.0）。
  本模块生成的草稿即用此格式，可在 v10.9 / v11.2 中直接打开。
- 旧版（v5.x）打开时会自动提示“升级草稿”，同样可用。
- 注意：剪映 v10.9+ 关闭了 headless 自动导出（防自动化），因此“导出 MP4”一律由
  调用方用 ffmpeg 完成（版本无关、无需剪映），本模块只负责生成“可在剪映里继续编辑/手动导出”的工程。
"""

import os
import sys
import json
import time
import uuid
import subprocess
from pathlib import Path

# ----------------------------------------------------------------------------
# 路径常量
# ----------------------------------------------------------------------------
HERE = Path(__file__).resolve().parent
TEMPLATE_DRAFT = HERE / "draft_template_v11.json"
TEMPLATE_META = HERE / "draft_meta_template_v11.json"

# 剪映草稿默认根目录（v11.2 实测路径）
DEFAULT_JY_ROOT = Path(os.path.expandvars(
    r"%LOCALAPPDATA%\JianyingPro\User Data\Projects\com.lveditor.draft"
))


# ----------------------------------------------------------------------------
# 1. 版本探测
# ----------------------------------------------------------------------------
def _scan_installed_jianying_versions():
    """扫描 JianyingPro\\Apps 下所有版本目录，返回版本号字符串列表。"""
    roots = []
    local = os.path.expandvars(r"%LOCALAPPDATA%")
    candidates = [
        Path(local) / "JianyingPro" / "Apps",
        Path(r"C:\Users") ,  # 兜底，不遍历
    ]
    found = []
    apps = Path(local) / "JianyingPro" / "Apps"
    if apps.is_dir():
        for sub in apps.iterdir():
            if sub.is_dir():
                # 目录名形如 11.2.0.14339 或 2026818181301363_1
                name = sub.name
                if name[0].isdigit() and "." in name:
                    found.append(name)
    return found


def detect_jianying_version():
    """
    探测本机剪映专业版版本。
    优先：扫描安装目录；其次：环境变量 JY_VERSION；最后：默认 "11.2.0"。
    返回 (version_str, source) 。
    """
    env = os.environ.get("JY_VERSION", "").strip()
    if env:
        return env, "env"
    vers = _scan_installed_jianying_versions()
    if vers:
        # 取“最高”版本（按数字段比较，过滤非纯数字段）
        def keyfn(v):
            parts = v.split(".")
            return [int(p) if p.isdigit() else 0 for p in parts]
        try:
            best = max(vers, key=keyfn)
            return best, "installed"
        except Exception:
            return vers[0], "installed"
    return "11.2.0", "default"


# ----------------------------------------------------------------------------
# 2. 工具
# ----------------------------------------------------------------------------
def _uid():
    return uuid.uuid4().hex


def _media_wh(path, ffmpeg):
    """用 ffmpeg -i 解析视频宽高（失败回退 1920x1080）。"""
    try:
        r = subprocess.run([str(ffmpeg), "-hide_banner", "-i", str(path), "-f", "null", "-"],
                           capture_output=True, text=True)
        out = r.stderr
        import re
        m = re.search(r"Stream.*Video:.*?(\d{2,5})x(\d{2,5})", out)
        if m:
            return int(m.group(1)), int(m.group(2))
    except Exception:
        pass
    return 1920, 1080


# ----------------------------------------------------------------------------
# 3. 生成 v10.9 / v11.2 兼容草稿
# ----------------------------------------------------------------------------
def build_v11_draft(project_dir, clips, subtitles,
                    fps=30, width=1920, height=1080,
                    project_name="", ffmpeg=None,
                    subtitle_size=3.4, subtitle_y=-0.82):
    """
    写出一份剪映 v10.9 / v11.2 兼容草稿（扁平布局）。

    参数：
      project_dir : 草稿工程目录（将写入 draft_content.json 等）
      clips       : list[dict] 视频片段，每项为
                    {"path": str, "start": float(秒), "duration": float(秒),
                     "name": str(可选), "width": int(可选), "height": int(可选)}
      subtitles   : list[dict] 字幕，每项为
                    {"text": str, "start": float(秒), "duration": float(秒)}
      fps,width,height : 画布参数（默认 1920x1080 / 30fps）
      project_name: 工程名

    返回：project_dir (Path)
    """
    if not TEMPLATE_DRAFT.exists():
        raise FileNotFoundError(f"找不到模板：{TEMPLATE_DRAFT}")
    tpl = json.load(open(TEMPLATE_DRAFT, encoding="utf-8"))

    project_dir = Path(project_dir)
    project_dir.mkdir(parents=True, exist_ok=True)
    project_name = project_name or project_dir.name

    video_materials = []
    speed_materials = []
    video_segments = []
    render_idx = 0

    for c in clips:
        path = str(c["path"])
        start_us = int(round(c["start"] * 1_000_000))
        dur_us = int(round(c["duration"] * 1_000_000))
        w = c.get("width") or (ffmpeg and _media_wh(path, ffmpeg)[0]) or width
        h = c.get("height") or (ffmpeg and _media_wh(path, ffmpeg)[1]) or height
        name = c.get("name") or os.path.basename(path)

        vid = _uid()
        sid = _uid()

        video_materials.append({
            "audio_fade": None,
            "category_id": "",
            "category_name": "local",
            "check_flag": 63487,
            "crop": {"upper_left_x": 0.0, "upper_left_y": 0.0,
                     "upper_right_x": 1.0, "upper_right_y": 0.0,
                     "lower_left_x": 0.0, "lower_left_y": 1.0,
                     "lower_right_x": 1.0, "lower_right_y": 1.0},
            "crop_ratio": "free",
            "crop_scale": 1.0,
            "duration": dur_us,
            "height": h,
            "id": vid,
            "local_material_id": "",
            "material_id": vid,
            "material_name": name,
            "media_path": "",
            "path": os.path.abspath(path),
            "type": "video",
            "video_algorithm": {"algorithms": [], "complement_frame_config": None,
                                 "deflicker": None, "gameplay_configs": [],
                                 "motion_blur_config": None, "noise_reduction": None,
                                 "path": "", "quality_enhance": None, "time_range": None},
            "source_platform": 0,
            "team_id": "",
            "width": w,
        })
        speed_materials.append({
            "curve_speed": None,
            "id": sid,
            "mode": 0,
            "speed": 1.0,
            "type": "speed",
        })
        video_segments.append({
            "enable_adjust": True,
            "enable_color_correct_adjust": False,
            "enable_color_curves": True,
            "enable_color_match_adjust": False,
            "enable_color_wheels": True,
            "enable_lut": True,
            "enable_smart_color_adjust": False,
            "last_nonzero_volume": 1.0,
            "reverse": False,
            "track_attribute": 0,
            "track_render_index": 0,
            "visible": True,
            "id": _uid(),
            "material_id": vid,
            "target_timerange": {"start": start_us, "duration": dur_us},
            "common_keyframes": [],
            "keyframe_refs": [],
            "source_timerange": {"start": 0, "duration": dur_us},
            "speed": 1.0,
            "volume": 1.0,
            "extra_material_refs": [sid],
            "is_tone_modify": False,
            "clip": {"alpha": 1.0,
                     "flip": {"horizontal": False, "vertical": False},
                     "rotation": 0.0,
                     "scale": {"x": 1.0, "y": 1.0},
                     "transform": {"x": 0.0, "y": 0.0}},
            "uniform_scale": {"on": True, "value": 1.0},
            "hdr_settings": {"intensity": 1.0, "mode": 1, "nits": 1000},
            "render_index": render_idx,
        })
        render_idx += 1

    text_materials = []
    text_segments = []
    for j, s in enumerate(subtitles):
        txt = s["text"]
        start_us = int(round(s["start"] * 1_000_000))
        dur_us = int(round(s["duration"] * 1_000_000))
        tid = _uid()
        content = {
            "styles": [{
                "fill": {"alpha": 1.0,
                         "content": {"render_type": "solid",
                                      "solid": {"alpha": 1.0, "color": [1.0, 1.0, 1.0]}}},
                "range": [0, len(txt)],
                "size": subtitle_size,
                "bold": False,
                "italic": False,
                "underline": False,
                "strokes": [{"content": {"solid": {"alpha": 0.6, "color": [0.0, 0.0, 0.0]}},
                             "width": 0.036}],
            }],
            "text": txt,
        }
        text_materials.append({
            "id": tid,
            "content": json.dumps(content, ensure_ascii=False),
            "typesetting": 0,
            "alignment": 0,
            "letter_spacing": 0.0,
            "line_spacing": 0.02,
            "line_feed": 1,
            "line_max_width": 0.82,
            "force_apply_line_max_width": False,
            "check_flag": 15,
            "type": "text",
            "global_alpha": 1.0,
        })
        text_segments.append({
            "enable_adjust": True,
            "enable_color_correct_adjust": False,
            "enable_color_curves": True,
            "enable_color_match_adjust": False,
            "enable_color_wheels": True,
            "enable_lut": True,
            "enable_smart_color_adjust": False,
            "last_nonzero_volume": 1.0,
            "reverse": False,
            "track_attribute": 0,
            "track_render_index": 0,
            "visible": True,
            "id": _uid(),
            "material_id": tid,
            "target_timerange": {"start": start_us, "duration": dur_us},
            "common_keyframes": [],
            "keyframe_refs": [],
            "source_timerange": None,
            "speed": 1.0,
            "volume": 1.0,
            "extra_material_refs": [],
            "is_tone_modify": False,
            "clip": {"alpha": 1.0,
                     "flip": {"horizontal": False, "vertical": False},
                     "rotation": 0.0,
                     "scale": {"x": 1.0, "y": 1.0},
                     "transform": {"x": 0.0, "y": subtitle_y}},
            "uniform_scale": {"on": True, "value": 1.0},
            "render_index": 15000 + j,
        })

    # 组装 tracks
    video_track = {"attribute": 0, "flag": 0, "id": _uid(),
                   "is_default_name": False, "name": "Video",
                   "type": "video", "segments": video_segments}
    text_track = {"attribute": 0, "flag": 0, "id": _uid(),
                  "is_default_name": False, "name": "Subtitles",
                  "type": "text", "segments": text_segments}

    total_us = 0
    for c in clips:
        total_us = max(total_us, int(round((c["start"] + c["duration"]) * 1_000_000)))

    # 填充模板
    tpl["materials"]["videos"] = video_materials
    tpl["materials"]["speeds"] = speed_materials
    tpl["materials"]["texts"] = text_materials
    tpl["tracks"] = [video_track, text_track]
    tpl["duration"] = total_us
    tpl["id"] = _uid()
    tpl["name"] = project_name
    tpl["canvas_config"] = {"width": width, "height": height, "ratio": "16:9"}
    now_ms = int(time.time() * 1000)
    tpl["create_time"] = now_ms
    tpl["update_time"] = now_ms

    draft_content = json.dumps(tpl, ensure_ascii=False, indent=1)
    (project_dir / "draft_content.json").write_text(draft_content, encoding="utf-8")
    # draft_info.json 与 draft_content.json 内容一致（剪映 v11.2 实测）
    (project_dir / "draft_info.json").write_text(draft_content, encoding="utf-8")

    # draft_meta_info.json
    if TEMPLATE_META.exists():
        meta = json.load(open(TEMPLATE_META, encoding="utf-8"))
        meta["draft_id"] = str(uuid.uuid4()).upper()
        meta["draft_name"] = project_name
        meta["draft_fold_path"] = str(project_dir).replace("/", "\\")
        meta["draft_root_path"] = str(project_dir.parent).replace("/", "\\")
        meta["tm_draft_create"] = now_ms * 1000  # 实测值为 16 位时间戳
        meta["tm_draft_modified"] = now_ms * 1000
        (project_dir / "draft_meta_info.json").write_text(
            json.dumps(meta, ensure_ascii=False, indent=1), encoding="utf-8")

    # draft_settings
    (project_dir / "draft_settings").write_text(
        "[General]\ndraft_create_time=0\ndraft_last_edit_time=0\nreal_edit_keys=1\nreal_edit_seconds=0\n",
        encoding="utf-8")
    # key_value.json
    (project_dir / "key_value.json").write_text("{}\n", encoding="utf-8")

    return project_dir


# ----------------------------------------------------------------------------
# 4. 统一出口
# ----------------------------------------------------------------------------
def export_draft(clips, subtitles, project_name="",
                 draft_root=None, fps=30, width=1920, height=1080,
                 ffmpeg=None, jy_version=None, subtitle_size=3.4, subtitle_y=-0.82):
    """
    生成剪映草稿（版本感知）。

    参数：
      clips / subtitles : 见 build_v11_draft
      project_name      : 工程名
      draft_root        : 草稿根目录；默认用剪映工程目录（若存在）否则用工作目录下的 jianying_drafts/
      jy_version        : 指定版本（默认自动探测）

    返回：(project_dir, version_str)
    """
    version_str, src = detect_jianying_version() if jy_version is None else (jy_version, "explicit")
    print(f"[jy_draft] 探测到剪映版本：{version_str}（来源：{src}）")

    if draft_root is None:
        if DEFAULT_JY_ROOT.exists():
            draft_root = DEFAULT_JY_ROOT
        else:
            draft_root = Path.cwd() / "jianying_drafts"
    draft_root = Path(draft_root)
    draft_root.mkdir(parents=True, exist_ok=True)

    # 工程目录（避免覆盖，带时间戳）
    safe_name = (project_name or "micro_lesson").replace(" ", "_")
    project_dir = draft_root / f"{safe_name}_{int(time.time())}"
    n = 1
    while project_dir.exists():
        project_dir = draft_root / f"{safe_name}_{int(time.time())}_{n}"
        n += 1

    # v10.9 / v11.2 采用同一现代格式；旧版打开会自动升级
    out = build_v11_draft(
        project_dir, clips, subtitles,
        fps=fps, width=width, height=height,
        project_name=safe_name, ffmpeg=ffmpeg,
        subtitle_size=subtitle_size, subtitle_y=subtitle_y,
    )
    print(f"[jy_draft] 已生成剪映草稿（v{version_str} 兼容）：{out}")
    return out, version_str


if __name__ == "__main__":
    # 简单自测：生成一份示例草稿
    pd = export_draft(
        clips=[{"path": r"I:\宝葫芦\素材\宝葫芦的秘密导入.mp4", "start": 0.0, "duration": 5.0,
                "name": "导入视频"},
               {"path": r"I:\宝葫芦\work\slide_01.mp4", "start": 5.0, "duration": 3.5,
                "name": "幻灯片1"}],
        subtitles=[{"text": "今天，让我们一起走进一个神奇的童话世界。", "start": 5.0, "duration": 3.5}],
        project_name="jy_draft_selftest",
    )
    print("SELFTEST OK ->", pd)
