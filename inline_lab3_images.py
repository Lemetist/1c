import base64
import re
from pathlib import Path


HTML_PATH = Path("/home/lemetist/1c/Лаба 3.html")


def replace_src(match: re.Match[str]) -> str:
    relative_path = match.group(1)
    image_path = HTML_PATH.parent / relative_path
    encoded = base64.b64encode(image_path.read_bytes()).decode("ascii")
    return f'src="data:image/png;base64,{encoded}"'


def main() -> None:
    html = HTML_PATH.read_text(encoding="utf-8")
    html = re.sub(r'src="(lab3_code_images/code_\d+\.png)"', replace_src, html)
    HTML_PATH.write_text(html, encoding="utf-8")


if __name__ == "__main__":
    main()
