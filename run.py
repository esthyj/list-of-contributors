from pathlib import Path
import sys
from nbclient import NotebookClient
from nbformat import read, write, NO_CONVERT
import subprocess

BASE = Path(__file__).parent
RTF_NOTEBOOK = BASE / "rtf_txt.ipynb"
WORD_SCRIPT = BASE / "word.py"


def run_notebook(ipynb_path: Path, out_path: Path, timeout=1200):
    if not ipynb_path.exists():
        raise FileNotFoundError(f"노트북 없음: {ipynb_path}")
    nb = read(ipynb_path.open("r", encoding="utf-8"), as_version=NO_CONVERT)
    client = NotebookClient(
        nb, timeout=timeout, kernel_name="python3", allow_errors=False
    )
    client.execute()
    out_path.parent.mkdir(parents=True, exist_ok=True)
    write(nb, out_path.open("w", encoding="utf-8"))


def main():
    print("▶ rtf_txt.ipynb 실행 중…")
    run_notebook(RTF_NOTEBOOK, BASE / "rtf_txt_executed.ipynb")
    print("✓ rtf_txt.ipynb 완료")

    print("▶ word.py 실행 중…")
    subprocess.run([sys.executable, str(WORD_SCRIPT)], check=True)
    print("✓ word.py 완료")

    print("모든 실행 완료!")


if __name__ == "__main__":
    main()
