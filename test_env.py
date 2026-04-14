import sys
print(sys.version)
from dotenv import load_dotenv
from pathlib import Path
load_dotenv(Path(__file__).parent / ".env")
import os
tok = os.getenv("TELEGRAM_BOT_TOKEN")
print("Token:", "OK" if tok else "MISSING")
print("Done")
