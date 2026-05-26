from __future__ import annotations

import argparse
import re
import sys
from email import policy
from email.message import EmailMessage, Message
from email.parser import BytesParser
from pathlib import Path


BASE_DIR = Path(__file__).resolve().parent
SOURCE_CANDIDATES = [
    BASE_DIR / "Reference Data" / "Email Templates" / "Steven" / "GFTEncroachemtn&Video.eml",
    BASE_DIR / "Reference Data" / "Email Templates" / "Steven" / "GFTEncroachment&Video.eml",
]


def _find_source_path() -> Path:
    for path in SOURCE_CANDIDATES:
        if path.exists():
            return path
    expected = "\n".join(f"  - {path}" for path in SOURCE_CANDIDATES)
    raise FileNotFoundError(f"Could not find the GFT source template. Checked:\n{expected}")


def _parse_message(path: Path) -> Message:
    with path.open("rb") as handle:
        return BytesParser(policy=policy.default).parse(handle)


def _first_part(message: Message, content_type: str) -> Message | None:
    if message.get_content_type() == content_type:
        return message
    for part in message.walk():
        if part.get_content_type() == content_type:
            return part
    return None


def _line_split_after_thanks(text: str) -> tuple[str, str]:
    lines = text.splitlines(keepends=True)
    for idx, line in enumerate(lines):
        if "thank you," in line.lower():
            return "".join(lines[: idx + 1]), "".join(lines[idx + 1 :])
    raise ValueError('Could not find "Thank you," in plain text content.')


def _html_split_after_thanks(html: str) -> tuple[str, str]:
    match = re.search(r"thank you,", html, flags=re.IGNORECASE)
    if not match:
        raise ValueError('Could not find "Thank you," in HTML content.')

    paragraph_end = re.search(r"</p\s*>", html[match.end() :], flags=re.IGNORECASE)
    if paragraph_end:
        split_at = match.end() + paragraph_end.end()
    else:
        split_at = match.end()

    return html[:split_at], html[split_at:]


def _set_text_part_content(part: Message, content: str, subtype: str) -> None:
    charset = part.get_content_charset() or "utf-8"
    cte = (part.get("Content-Transfer-Encoding") or "").lower()
    valid_ctes = {"7bit", "8bit", "quoted-printable", "base64"}
    kwargs = {"subtype": subtype, "charset": charset}
    if cte in valid_ctes:
        kwargs["cte"] = cte
    part.set_content(content, **kwargs)


def format_eml_in_place(target_path: Path, source_path: Path) -> None:
    if not target_path.exists():
        raise FileNotFoundError(f"Target file does not exist: {target_path}")
    if target_path.suffix.lower() != ".eml":
        raise ValueError(f"Target is not an .eml file: {target_path}")

    source_message = _parse_message(source_path)
    target_message = _parse_message(target_path)

    source_html_part = _first_part(source_message, "text/html")
    target_html_part = _first_part(target_message, "text/html")
    if source_html_part is None or target_html_part is None:
        raise ValueError("Both source and target messages must contain a text/html part.")

    source_html = source_html_part.get_content()
    target_html = target_html_part.get_content()
    _, source_html_signature = _html_split_after_thanks(source_html)
    target_html_intro, _ = _html_split_after_thanks(target_html)
    _set_text_part_content(target_html_part, target_html_intro + source_html_signature, "html")

    source_text_part = _first_part(source_message, "text/plain")
    target_text_part = _first_part(target_message, "text/plain")
    if source_text_part is not None and target_text_part is not None:
        source_text = source_text_part.get_content()
        target_text = target_text_part.get_content()
        _, source_text_signature = _line_split_after_thanks(source_text)
        target_text_intro, _ = _line_split_after_thanks(target_text)
        _set_text_part_content(target_text_part, target_text_intro + source_text_signature, "plain")

    target_path.write_bytes(target_message.as_bytes(policy=policy.SMTP))


def main(argv: list[str]) -> int:
    parser = argparse.ArgumentParser(
        description=(
            "Format dropped .eml files in place by replacing the signature after "
            '"Thank you," with the Steven GFT signature block.'
        )
    )
    parser.add_argument("eml_files", nargs="+", help="One or more .eml files to update in place.")
    args = parser.parse_args(argv)

    source_path = _find_source_path()
    for raw_path in args.eml_files:
        target_path = Path(raw_path).expanduser().resolve()
        format_eml_in_place(target_path, source_path)
        print(f"Updated in place: {target_path}")
    return 0


if __name__ == "__main__":
    try:
        raise SystemExit(main(sys.argv[1:]))
    except Exception as exc:
        print(f"Error: {exc}", file=sys.stderr)
        raise SystemExit(1)
