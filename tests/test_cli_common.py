import argparse
import base64

from backend.cli.common import (
    _bare_subject,
    _ensure_utf8_charset,
    _inject_before_body,
    _normalize_cc,
    _resolve_body,
    _wrap_body_font,
)


def _body_args(**overrides):
    values = {
        "body": None,
        "body_file": None,
        "body_stdin": False,
        "body_stdin_base64": False,
        "body_base64": None,
    }
    values.update(overrides)
    return argparse.Namespace(**values)


def test_wrap_body_font_escapes_plain_text_and_preserves_lines():
    result = _wrap_body_font("A & B\nsecond line")

    assert "A &amp; B<br>second line" in result
    assert "font-family:Calibri" in result


def test_wrap_body_font_preserves_html_input():
    result = _wrap_body_font("<p>A & B</p>")

    assert "<p>A & B</p>" in result
    assert "&amp;" not in result


def test_ensure_utf8_charset_inserts_meta_into_head():
    result = _ensure_utf8_charset("<html><head><title>x</title></head><body>y</body></html>")

    assert 'charset=utf-8' in result.lower()
    assert result.lower().index('charset=utf-8') < result.lower().index("</head>")


def test_ensure_utf8_charset_replaces_existing_charset():
    result = _ensure_utf8_charset(
        '<html><head><meta charset="windows-1252"></head><body>x</body></html>'
    )

    assert "windows-1252" not in result.lower()
    assert result.lower().count("utf-8") == 1


def test_inject_before_body_places_content_inside_body():
    result = _inject_before_body("<p>new</p>", "<html><body>old</body></html>")

    assert "<body><p>new</p>old</body>" in result


def test_normalize_cc_accepts_none_string_and_list():
    assert _normalize_cc(None) == ""
    assert _normalize_cc("a@example.com") == "a@example.com"
    assert _normalize_cc(["a@example.com", "b@example.com,c@example.com"]) == (
        "a@example.com,b@example.com,c@example.com"
    )


def test_bare_subject_removes_repeated_reply_and_forward_prefixes():
    assert _bare_subject("RE: FW: Re: Project Update") == "Project Update"


def test_resolve_body_reads_utf8_file(tmp_path):
    body_file = tmp_path / "body.html"
    body_file.write_text("<p>Cost is $60.</p>", encoding="utf-8")
    args = _body_args(body_file=str(body_file))

    assert _resolve_body(args) is True
    assert args.body == "<p>Cost is $60.</p>"


def test_resolve_body_decodes_base64():
    body = "<p>??</p>"
    payload = base64.b64encode(body.encode("utf-8")).decode("ascii")
    args = _body_args(body_base64=payload)

    assert _resolve_body(args) is True
    assert args.body == body


def test_resolve_body_rejects_invalid_base64(capsys):
    args = _body_args(body_base64="not-base64")

    assert _resolve_body(args) is False
    assert "invalid --body-base64 payload" in capsys.readouterr().err
