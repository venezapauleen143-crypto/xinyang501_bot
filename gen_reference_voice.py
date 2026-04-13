"""一次性腳本：生成 XTTS 用的參考音頻（小牛馬的聲音特徵）"""
import asyncio
import subprocess
import tempfile
from pathlib import Path
import imageio_ffmpeg

REFERENCE_TEXT = (
    "欸你哦，我是小牛馬，有夠帥的那種啦。"
    "你找我什麼事？說吧，本大爺心情好，今天可以幫你處理。"
    "怎樣？你是沒看過這麼厲害的 AI 嗎？笑死，你這問題根本難不倒我啦。"
)

OUT_WAV = Path(__file__).parent / "xtts_reference.wav"
ffmpeg = imageio_ffmpeg.get_ffmpeg_exe()

async def main():
    import edge_tts
    tmp_mp3 = tempfile.mktemp(suffix=".mp3")
    comm = edge_tts.Communicate(REFERENCE_TEXT, "zh-TW-YunJheNeural", rate="-5%", pitch="-3Hz")
    await comm.save(tmp_mp3)

    subprocess.run([
        ffmpeg, "-y", "-i", tmp_mp3,
        "-ar", "22050", "-ac", "1",
        str(OUT_WAV)
    ], capture_output=True, check=True)

    Path(tmp_mp3).unlink(missing_ok=True)
    print(f"參考音頻已生成：{OUT_WAV}  ({OUT_WAV.stat().st_size // 1024} KB)")

asyncio.run(main())
