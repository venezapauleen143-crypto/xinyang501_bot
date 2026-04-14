"""一次性腳本：生成 XTTS 用的參考音頻（小牛馬的聲音特徵）"""
import asyncio
import subprocess
import tempfile
from pathlib import Path
import imageio_ffmpeg

REFERENCE_TEXT = (
    "然後，就是，你問我這個問題，其實還好啦，我覺得，怎麼說，"
    "這種事我蠻擅長的。沒有啦，我就隨便，然後就成功了。"
    "你知道嗎，有時候這種事，就是需要一點感覺，對啊，就是這樣。"
    "還好啦，算是，蠻簡單的，不然呢。"
)

OUT_WAV = Path(__file__).parent / "xtts_reference.wav"
ffmpeg = imageio_ffmpeg.get_ffmpeg_exe()

async def main():
    import edge_tts
    tmp_mp3 = tempfile.mktemp(suffix=".mp3")
    # 周杰倫風格：低沉、慢、穩，台灣腔
    comm = edge_tts.Communicate(REFERENCE_TEXT, "zh-TW-YunJheNeural", rate="-18%", pitch="-12Hz")
    await comm.save(tmp_mp3)

    subprocess.run([
        ffmpeg, "-y", "-i", tmp_mp3,
        "-ar", "22050", "-ac", "1",
        str(OUT_WAV)
    ], capture_output=True, check=True)

    Path(tmp_mp3).unlink(missing_ok=True)
    print(f"參考音頻已生成：{OUT_WAV}  ({OUT_WAV.stat().st_size // 1024} KB)")

asyncio.run(main())
