import os
import sys
import json
import re
import time
import subprocess
import asyncio
from pathlib import Path

# ============================================================
# 微课生成脚本（v3）：PPT 课件 + 导入视频 → 剪映草稿 + 导出 MP4
# 改进：字幕按标点断开为短句（每行不超过 16 字、绝对单行）、字号更小(18)、
#        白字 + 半透明黑底框（无硬黑边）、每句配独立 AI 配音，字幕与配音严格同步。
# ============================================================

WORK_DIR = Path(r"I:\宝葫芦\work")
ASSET_DIR = Path(r"I:\宝葫芦\素材")
SLIDE_DIR = WORK_DIR / "slides"
OUT_MP4 = ASSET_DIR / "宝葫芦的秘密_微课.mp4"

FFMPEG = Path(r"H:\WorkBuddy\tools\ffmpeg-bin\ffmpeg.exe")
PYTHON = Path(r"C:\Users\Administrator\.workbuddy\binaries\python\envs\default\Scripts\python.exe")
SKILL_ROOT = Path(r"H:\WorkBuddy\skill\jianying-editor-skill")

IMPORT_VIDEO = ASSET_DIR / "宝葫芦的秘密导入.mp4"
PPTX = ASSET_DIR / "宝葫芦的秘密课件.pptx"

# 女教师音色，清晰亲切
EDGE_VOICE = "zh-CN-XiaoxiaoNeural"

# 字幕参数
SUB_FONT_SIZE = 18          # 字号（更小）
SUB_MAX_CHARS = 16          # 每行最多字数（绝对单行）
SUB_MIN_CHARS = 3           # 软断点最小长度

# 课件核心内容：第 1-15 页。第 16-22 页是通用商业模板（Lorem ipsum/未来规划），与课文无关，已排除。
SLIDES = [
    {"narration": "今天，让我们一起走进一个神奇的童话世界。", "min_dur": 3.5},
    {"narration": "推开这扇窗，去认识一个特别的葫芦。", "min_dur": 3.5},
    {"narration": "同学们好！今天我们要学习四年级下册语文课文《宝葫芦的秘密》节选。", "min_dur": 5.0},
    {"narration": "《宝葫芦的秘密》是著名作家张天翼爷爷的童话代表作。故事的主人公叫王葆。", "min_dur": 6.0},
    {"narration": "首先，请大家自由朗读课文。读准字音，读顺句子，遇到生字词多读几遍；借助工具书或联系上下文理解词语；最后想一想，课文主要讲了一件什么事？", "min_dur": 13.0},
    {"narration": "读完课文，请快速浏览，找一找：文中的王葆是怎么知道宝葫芦的故事的？试着用自己的话归纳出来。", "min_dur": 10.0},
    {"narration": "原来，奶奶每逢要求王葆干什么，就得给他讲一个故事。这在王葆家里已经成了——规矩。从“规矩”二字，你能感受到祖孙之间怎样的情感呢？", "min_dur": 11.0},
    {"narration": "现在，请大家默读课文的第十五到第十七自然段，边读边想：文中都列举了哪些关于宝葫芦的故事？通过阅读，你又发现了什么？", "min_dur": 12.0},
    {"narration": "看，张三劈面撞见了一位神仙，得到了一个宝葫芦。这些故事都有相同的结构：先讲宝葫芦的来历，再讲它的作用。比如，张三想吃水蜜桃，立刻就有一盘水蜜桃。是不是很神奇？", "min_dur": 16.0},
    {"narration": "接下来，请同学们按照“来历—作用”的框架，发挥想象，针对课文中已有的人物进行故事创编。张三、李四、王五、赵六，他们会用宝葫芦做什么呢？", "min_dur": 15.0},
    {"narration": "联系生活想一想：你有没有需要宝葫芦的时候？如果你也有了一个宝葫芦，你会让它帮你做什么呢？", "min_dur": 11.0},
    {"narration": "可是，宝葫芦真的能解决所有问题吗？不会做的数学题，有了宝葫芦，仍不会做；不会种的向日葵，仍不会种；不会和同学相处，还是不会。看来，不劳而获，换不来真正的本领。", "min_dur": 16.0},
    {"narration": "正如课文告诉我们的：没有付出的收获，就是不劳而获；不劳而获的幸福，不是真的幸福。真正的幸福，要靠自己的努力和付出。", "min_dur": 12.0},
    {"narration": "课后拓展：当王葆真的得到一个宝葫芦时，他逐渐认识到，靠宝葫芦不劳而获，带给他的不是幸福，而是烦恼。这是怎么回事呢？感兴趣的同学可以读一读《宝葫芦的秘密》整本书。", "min_dur": 14.0},
    {"narration": "今天的学习就到这里，谢谢大家！", "min_dur": 4.0},
]


def run(cmd, **kw):
    print("[RUN]", " ".join(str(c) for c in cmd))
    r = subprocess.run(cmd, capture_output=True, text=True, **kw)
    if r.returncode != 0:
        print("STDOUT:", r.stdout[-800:] if r.stdout else "")
        print("STDERR:", r.stderr[-1200:] if r.stderr else "")
        raise RuntimeError(f"Command failed: {cmd} -> {r.returncode}")
    return r


def get_media_duration(path):
    """使用 ffmpeg -i 解析时长（秒）。"""
    r = subprocess.run([str(FFMPEG), "-hide_banner", "-i", str(path), "-f", "null", "-"],
                       capture_output=True, text=True)
    m = re.search(r"Duration:\s+(\d+):(\d+):(\d+\.\d+)", r.stderr)
    if not m:
        return None
    h, m_, s = m.groups()
    return int(h) * 3600 + int(m_) * 60 + float(s)


def format_srt_time(seconds):
    milliseconds = int(round(seconds * 1000))
    hours = milliseconds // 3600000
    minutes = (milliseconds % 3600000) // 60000
    secs = (milliseconds % 60000) // 1000
    ms = milliseconds % 1000
    return f"{hours:02d}:{minutes:02d}:{secs:02d},{ms:03d}"


def split_chunks(text, max_chars=SUB_MAX_CHARS, min_chars=SUB_MIN_CHARS):
    """按标点断开成短句，每行不超过 max_chars（+拖尾标点，最多 1 字），绝对单行。"""
    PUNCT = set("。！？；，、：—…")
    segs = []
    buf = ""
    for ch in text:
        buf += ch
        if ch in "。！？；":            # 句末标点：强制断
            segs.append(buf)
            buf = ""
        elif ch in "，、：—…":        # 句中标点：达到最小长度才断
            if len(buf) >= min_chars:
                segs.append(buf)
                buf = ""
    if buf:
        segs.append(buf)
    # 二次切分：超长且无标点可断的中段，硬切断；剩余以标点开头时把标点并入当前块，避免孤立标点
    out = []
    for seg in segs:
        seg = seg.strip()
        if not seg:
            continue
        while len(seg) > max_chars:
            cut = seg[:max_chars]
            rest = seg[max_chars:]
            if rest and rest[0] in PUNCT:   # 把拖尾标点并入，防止出现孤立标点短句
                cut = cut + rest[0]
                rest = rest[1:]
            out.append(cut)
            seg = rest
        if seg:
            out.append(seg)
    return out


async def generate_tts(text, out_path, retries=4):
    import edge_tts
    last_err = None
    for attempt in range(1, retries + 1):
        try:
            communicate = edge_tts.Communicate(text, EDGE_VOICE)
            await communicate.save(str(out_path))
            if out_path.exists() and out_path.stat().st_size > 0:
                return
            last_err = RuntimeError("empty file")
        except Exception as e:
            last_err = e
        print(f"    [retry {attempt}/{retries}] TTS 失败: {last_err}")
        await asyncio.sleep(2 * attempt)
    raise RuntimeError(f"TTS 多次失败: {text!r} -> {last_err}")


async def generate_all_tts():
    """逐短句生成配音（每句独立音频，保证字幕与配音严格同步）。"""
    print("\n=== 生成 AI 配音（edge-TTS，逐短句） ===")
    chunk_timings = {}
    total_chunks = 0
    for i, s in enumerate(SLIDES, 1):
        chunks = split_chunks(s["narration"])
        dur_list = []
        part_mp3s = []
        for k, chunk in enumerate(chunks, 1):
            mp3 = WORK_DIR / f"tts_{i:02d}_{k:02d}.mp3"
            if not (mp3.exists() and mp3.stat().st_size > 0):
                await generate_tts(chunk, mp3)
            else:
                print(f"    [skip] slide {i:02d} chunk {k:02d} 已存在")
            d = get_media_duration(mp3) or 1.0
            dur_list.append((chunk, d))
            part_mp3s.append(mp3)
        # 拼接为整页音频
        slide_mp3 = WORK_DIR / f"tts_{i:02d}.mp3"
        list_path = WORK_DIR / f"tts_{i:02d}_list.txt"
        with open(list_path, "w", encoding="utf-8") as f:
            for p in part_mp3s:
                f.write(f"file '{p.name}'\n")
        run([str(FFMPEG), "-hide_banner", "-y", "-f", "concat", "-safe", "0",
             "-i", str(list_path), "-c", "copy", str(slide_mp3)], cwd=str(WORK_DIR))
        total = sum(d for _, d in dur_list)
        chunk_timings[i] = dur_list
        total_chunks += len(chunks)
        print(f"  slide {i:02d}: {len(chunks)} 短句, 总时长 {total:.2f}s")
    print(f"  共 {total_chunks} 条字幕（全部单行）")
    return chunk_timings


def normalize_import_video():
    out = WORK_DIR / "import_norm.mp4"
    if out.exists():
        return out
    print("\n=== 统一导入视频分辨率到 1920x1080 ===")
    run([
        str(FFMPEG), "-hide_banner", "-y", "-i", str(IMPORT_VIDEO),
        "-vf", "scale=1920:1080:force_original_aspect_ratio=decrease,pad=1920:1080:-1:-1:black,setsar=1,format=yuv420p",
        "-af", "aformat=sample_fmts=fltp:sample_rates=44100:channel_layouts=stereo",
        "-c:v", "libx264", "-preset", "fast", "-crf", "23",
        "-c:a", "aac", "-b:a", "192k", "-ar", "44100", "-ac", "2",
        "-r", "30", "-pix_fmt", "yuv420p",
        str(out),
    ])
    return out


def build_slide_videos():
    print("\n=== 合成每页幻灯片视频（画面+AI配音） ===")
    durations = []
    for i, s in enumerate(SLIDES, 1):
        png = SLIDE_DIR / f"slide_{i:02d}.png"
        mp3 = WORK_DIR / f"tts_{i:02d}.mp3"
        out = WORK_DIR / f"slide_{i:02d}.mp4"

        audio_dur = get_media_duration(mp3) or 5.0
        dur = max(s["min_dur"], audio_dur + 0.4)
        durations.append(dur)
        print(f"  slide {i:02d}: audio {audio_dur:.2f}s -> video {dur:.2f}s")

        run([
            str(FFMPEG), "-hide_banner", "-y",
            "-loop", "1", "-i", str(png),
            "-i", str(mp3),
            "-c:v", "libx264", "-preset", "fast", "-crf", "23",
            "-pix_fmt", "yuv420p",
            "-t", f"{dur:.3f}",
            "-vf", "scale=1920:1080:force_original_aspect_ratio=decrease,pad=1920:1080:-1:-1:black,format=yuv420p",
            "-c:a", "aac", "-b:a", "192k", "-ar", "44100", "-ac", "2",
            "-r", "30",
            str(out),
        ])
    return durations


def build_srt(import_dur, slide_durs, chunk_timings):
    """逐短句生成字幕，单行、与配音同步出现。"""
    srt_path = WORK_DIR / "subtitles.srt"
    print("\n=== 生成字幕 SRT（逐句同步、单行） ===")
    entries = []
    idx = 1
    cursor = import_dur
    maxlen = 0
    for i, (s, dur) in enumerate(zip(SLIDES, slide_durs), 1):
        slide_start = cursor
        offset = 0.0
        for (chunk, cd) in chunk_timings.get(i, [(s["narration"], dur)]):
            start = slide_start + offset
            end = slide_start + offset + cd
            entries.append((idx, start, end, chunk))
            maxlen = max(maxlen, len(chunk))
            idx += 1
            offset += cd
        cursor = slide_start + dur

    with open(srt_path, "w", encoding="utf-8-sig") as f:
        for n, start, end, text in entries:
            f.write(f"{n}\n")
            f.write(f"{format_srt_time(start)} --> {format_srt_time(end)}\n")
            f.write(f"{text}\n\n")
    print(f"  wrote {len(entries)} cues -> {srt_path}（最长 {maxlen} 字，确保单行）")
    return srt_path


def concat_final_video(srt_path):
    print("\n=== 拼接导入视频与课件视频，并烧录字幕 ===")
    import_norm = WORK_DIR / "import_norm.mp4"

    list_path = WORK_DIR / "concat_list.txt"
    with open(list_path, "w", encoding="utf-8") as f:
        f.write(f"file '{import_norm.name}'\n")
        for i in range(1, len(SLIDES) + 1):
            f.write(f"file 'slide_{i:02d}.mp4'\n")

    # 字体样式：微软雅黑、字号更小(18)、白字 + 半透明黑底框（无硬黑边）、底部居中、单行
    style = (
        "FontName=Microsoft YaHei,"
        f"FontSize={SUB_FONT_SIZE},"
        "PrimaryColour=&H00FFFFFF,"
        "OutlineColour=&H00000000,"
        "Outline=0,"
        "Shadow=0,"
        "BorderStyle=4,"
        "BackColour=&H99000000,"
        "Alignment=2,"
        "MarginV=36"
    )
    vf = f"subtitles={srt_path.name}:force_style='{style}'"

    run([
        str(FFMPEG), "-hide_banner", "-y",
        "-f", "concat", "-safe", "0", "-i", str(list_path),
        "-vf", vf,
        "-c:v", "libx264", "-preset", "medium", "-crf", "23",
        "-c:a", "aac", "-b:a", "192k", "-ar", "44100", "-ac", "2",
        "-r", "30", "-pix_fmt", "yuv420p",
        "-movflags", "+faststart",
        str(OUT_MP4),
    ], cwd=str(WORK_DIR))
    print(f"  已导出：{OUT_MP4}")


def build_jy_project(import_dur, slide_durs, chunk_timings):
    """同时生成一个剪映专业版可编辑的草稿工程（字幕逐短句、与配音同步）。"""
    print("\n=== 生成剪映草稿工程 ===")
    sys.path.insert(0, str(SKILL_ROOT / "scripts"))
    from jy_wrapper import JyProject
    import pyJianYingDraft as draft

    import time
    project_name = f"宝葫芦的秘密_微课_{int(time.time())}"
    project = JyProject(project_name, width=1920, height=1080, overwrite=False)

    project.add_media_safe(str(IMPORT_VIDEO), "0s", duration=f"{import_dur:.3f}s", track_name="Video")

    cursor = import_dur + 0.1
    for i, dur in enumerate(slide_durs, 1):
        slide_mp4 = WORK_DIR / f"slide_{i:02d}.mp4"
        project.add_media_safe(str(slide_mp4), f"{cursor:.3f}s", duration=f"{dur:.3f}s", track_name="Video")
        cursor += dur + 0.1

    # 逐短句字幕，与配音同步
    cursor = import_dur + 0.1
    for i, (s, dur) in enumerate(zip(SLIDES, slide_durs), 1):
        offset = 0.0
        for (chunk, cd) in chunk_timings.get(i, [(s["narration"], dur)]):
            project.add_text_simple(
                chunk,
                start_time=f"{cursor + offset:.3f}s",
                duration=f"{cd:.3f}s",
                track_name="Subtitles",
                clip_settings=draft.ClipSettings(transform_y=-0.82),
                style=draft.TextStyle(size=2.4),
                border=draft.TextBorder(color=(0.0, 0.0, 0.0), alpha=0.55, width=14.0),
            )
            offset += cd
        cursor += dur + 0.1

    project.save()
    print(f"  剪映草稿已保存：{project.name}")


def verify_output():
    dur = get_media_duration(OUT_MP4)
    size = OUT_MP4.stat().st_size if OUT_MP4.exists() else 0
    print("\n=== 最终产物校验 ===")
    print(f"  MP4: {OUT_MP4}")
    print(f"  大小: {size / 1024 / 1024:.1f} MB")
    print(f"  时长: {dur:.2f}s ({int(dur // 60)}m {dur % 60:.1f}s)")


async def main():
    WORK_DIR.mkdir(parents=True, exist_ok=True)
    ASSET_DIR.mkdir(parents=True, exist_ok=True)
    SLIDE_DIR.mkdir(parents=True, exist_ok=True)

    if not IMPORT_VIDEO.exists():
        raise FileNotFoundError(f"找不到导入视频：{IMPORT_VIDEO}")
    if not PPTX.exists():
        raise FileNotFoundError(f"找不到课件：{PPTX}")
    if not FFMPEG.exists():
        raise FileNotFoundError(f"找不到 ffmpeg：{FFMPEG}")

    import_dur = get_media_duration(IMPORT_VIDEO)
    print(f"导入视频时长: {import_dur:.2f}s")

    chunk_timings = await generate_all_tts()
    import_norm = normalize_import_video()
    slide_durs = build_slide_videos()
    srt_path = build_srt(import_dur, slide_durs, chunk_timings)
    try:
        build_jy_project(import_dur, slide_durs, chunk_timings)
    except Exception as e:
        print(f"[WARN] 剪映草稿生成失败（不影响 MP4）：{e}")
    concat_final_video(srt_path)
    verify_output()


if __name__ == "__main__":
    asyncio.run(main())
