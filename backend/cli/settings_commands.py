"""Out-of-office and Outlook signature commands."""

from __future__ import annotations

import os
import sys


PR_OOF_STATE = 'http://schemas.microsoft.com/mapi/proptag/0x661D000B'


def cmd_get_ooo(args):
    """Get current Out of Office status."""
    try:
        import pythoncom
        pythoncom.CoInitialize()
        import win32com.client
        outlook = win32com.client.Dispatch('Outlook.Application')
        namespace = outlook.GetNamespace('MAPI')
        store = namespace.DefaultStore
        pa = store.PropertyAccessor

        oof_state = pa.GetProperty(PR_OOF_STATE)
        print(f"{'=' * 50}")
        print(f" Out of Office — {store.DisplayName}")
        print(f"{'=' * 50}")
        print(f"\n  Status: {'🟢 ENABLED' if oof_state else '⚪ DISABLED'}")
        if oof_state:
            print("\n  To disable: py -3 scripts/outlook_skill.py disable-ooo")
        else:
            print("\n  To enable: py -3 scripts/outlook_skill.py set-ooo")
        print(f"\n  💡 To view/edit auto-reply message: Outlook > File > Automatic Replies")

        pythoncom.CoUninitialize()
        return 0
    except Exception as e:
        print(f"Error: {str(e)}", file=sys.stderr)
        return 1

def cmd_set_ooo(args):
    """Enable Out of Office auto-reply."""
    try:
        import pythoncom
        pythoncom.CoInitialize()
        import win32com.client
        outlook = win32com.client.Dispatch('Outlook.Application')
        namespace = outlook.GetNamespace('MAPI')
        store = namespace.DefaultStore
        pa = store.PropertyAccessor

        current = pa.GetProperty(PR_OOF_STATE)
        if current:
            print(f"⚠️ Out of Office is already enabled for {store.DisplayName}.")
            pythoncom.CoUninitialize()
            return 0

        pa.SetProperty(PR_OOF_STATE, True)
        verified = pa.GetProperty(PR_OOF_STATE)
        if verified:
            print(f"✅ Out of Office ENABLED for {store.DisplayName}.")
            print(f"   💡 Make sure your auto-reply message is set in Outlook > File > Automatic Replies")
        else:
            print(f"⚠️ Failed to enable OOF — state did not change.")

        pythoncom.CoUninitialize()
        return 0
    except Exception as e:
        print(f"Error: {str(e)}", file=sys.stderr)
        return 1

def cmd_disable_ooo(args):
    """Disable Out of Office auto-reply."""
    try:
        import pythoncom
        pythoncom.CoInitialize()
        import win32com.client
        outlook = win32com.client.Dispatch('Outlook.Application')
        namespace = outlook.GetNamespace('MAPI')
        store = namespace.DefaultStore
        pa = store.PropertyAccessor

        current = pa.GetProperty(PR_OOF_STATE)
        if not current:
            print(f"⚪ Out of Office is already disabled for {store.DisplayName}.")
            pythoncom.CoUninitialize()
            return 0

        pa.SetProperty(PR_OOF_STATE, False)
        verified = pa.GetProperty(PR_OOF_STATE)
        if not verified:
            print(f"✅ Out of Office DISABLED for {store.DisplayName}.")
        else:
            print(f"⚠️ Failed to disable OOF — state did not change.")

        pythoncom.CoUninitialize()
        return 0
    except Exception as e:
        print(f"Error: {str(e)}", file=sys.stderr)
        return 1

def _get_signatures_path():
    """Return the Outlook signatures folder path."""
    return os.path.join(os.environ['APPDATA'], 'Microsoft', 'Signatures')

def _read_sig_file(filepath):
    """Read a signature file, handling UTF-16, UTF-8, and legacy encodings."""
    with open(filepath, 'rb') as f:
        raw = f.read()
    if raw[:2] in (b'\xff\xfe', b'\xfe\xff'):
        return raw.decode('utf-16')
    try:
        return raw.decode('utf-8')
    except UnicodeDecodeError:
        return raw.decode('latin-1')

def cmd_get_signature(args):
    """List all signatures or view a specific one."""
    try:
        sig_path = _get_signatures_path()
        if not os.path.exists(sig_path):
            print("No signatures folder found at:", sig_path)
            return 1

        htm_files = [f for f in os.listdir(sig_path) if f.endswith('.htm')]
        if not htm_files:
            print("No signatures found.")
            return 1

        name = getattr(args, 'name', None)
        fmt = getattr(args, 'format', 'text')

        if name:
            htm_path = os.path.join(sig_path, f"{name}.htm")
            txt_path = os.path.join(sig_path, f"{name}.txt")
            if not os.path.exists(htm_path):
                print(f"Signature '{name}' not found.")
                print(f"Available: {', '.join(f[:-4] for f in htm_files)}")
                return 1
            if fmt == 'html':
                print(_read_sig_file(htm_path))
            else:
                if os.path.exists(txt_path):
                    print(_read_sig_file(txt_path))
                else:
                    print(_read_sig_file(htm_path))
        else:
            print(f"{'=' * 60}")
            print(f" Signatures ({len(htm_files)} found)")
            print(f"{'=' * 60}")
            for htm_file in htm_files:
                sig_name = htm_file[:-4]
                txt_path = os.path.join(sig_path, f"{sig_name}.txt")
                print(f"\n--- {sig_name} ---")
                if os.path.exists(txt_path):
                    print(_read_sig_file(txt_path).strip())
                else:
                    print("(no .txt version available)")

        return 0
    except Exception as e:
        print(f"Error: {str(e)}", file=sys.stderr)
        return 1

def _update_rtf_signature(rtf_path, find_text, replace_text):
    """Update text in an RTF signature file. Returns True if updated."""
    if not os.path.exists(rtf_path):
        return False
    with open(rtf_path, 'rb') as f:
        content = f.read().decode('latin-1')
    escaped_find = _rtf_escape(find_text)
    escaped_replace = _rtf_escape(replace_text)
    if escaped_find in content:
        content = content.replace(escaped_find, escaped_replace)
        with open(rtf_path, 'w', encoding='latin-1', newline='') as f:
            f.write(content)
        return True
    if find_text in content:
        content = content.replace(find_text, escaped_replace)
        with open(rtf_path, 'w', encoding='latin-1', newline='') as f:
            f.write(content)
        return True
    return False

def _rtf_escape(text):
    """Escape non-ASCII characters for RTF using \\uN notation."""
    result = []
    for ch in text:
        if ord(ch) > 127:
            result.append(f'\\u{ord(ch)}?')
        else:
            result.append(ch)
    return ''.join(result)

def _rebuild_rtf_content(rtf_path, lines):
    """Replace all content paragraphs in RTF with new lines. Generates new RTF if needed."""
    import re as _re

    content = ''
    if os.path.exists(rtf_path):
        with open(rtf_path, 'rb') as f:
            content = f.read().decode('latin-1')

    # Try to find and replace existing content block
    if content:
        pattern = _re.compile(
            r'(\\ltrch\\fcs0 \\insrsid\d+ \\hich\\af31506\\dbch\\af31505\\loch\\f31506 )'
            r'(.+?)'
            r'(\r\n\\par \}\\pard)',
            _re.DOTALL
        )
        match = pattern.search(content)
        if match:
            rtf_prefix = r'\hich\af31506\dbch\af31505\loch\f31506 '
            rtf_lines = []
            for line in lines:
                if line.strip():
                    rtf_lines.append(rtf_prefix + _rtf_escape(line))
                else:
                    rtf_lines.append('')
            new_content = '\r\n\\par '.join(rtf_lines)
            content = content[:match.start(2)] + new_content + content[match.end(2):]
            with open(rtf_path, 'w', encoding='latin-1', newline='') as f:
                f.write(content)
            return True

    # Fallback: generate minimal valid RTF from scratch
    escaped_lines = []
    for line in lines:
        escaped_lines.append(_rtf_escape(line) if line.strip() else '')
    par_content = '\\par\r\n'.join(escaped_lines)

    rtf = (
        r'{\rtf1\ansi\ansicpg1252\deff0\nouicompat'
        r'{\fonttbl{\f0\fswiss\fcharset0 Aptos;}}'
        '\r\n'
        r'{\*\generator Microsoft Word 15}'
        r'\viewkind4\uc1'
        '\r\n'
        r'\pard\f0\fs22 '
        + par_content +
        '\r\n\\par }'
    )
    with open(rtf_path, 'w', encoding='latin-1', newline='') as f:
        f.write(rtf)
    return True

def cmd_update_signature(args):
    """Update a signature's content."""
    try:
        sig_path = _get_signatures_path()
        if not os.path.exists(sig_path):
            print("No signatures folder found at:", sig_path)
            return 1

        name = args.name
        htm_path = os.path.join(sig_path, f"{name}.htm")
        txt_path = os.path.join(sig_path, f"{name}.txt")
        rtf_path = os.path.join(sig_path, f"{name}.rtf")

        if not os.path.exists(htm_path):
            htm_files = [f for f in os.listdir(sig_path) if f.endswith('.htm')]
            print(f"Signature '{name}' not found.")
            if htm_files:
                print(f"Available: {', '.join(f[:-4] for f in htm_files)}")
            return 1

        import shutil
        shutil.copy2(htm_path, htm_path + '.bak')
        if os.path.exists(txt_path):
            shutil.copy2(txt_path, txt_path + '.bak')
        if os.path.exists(rtf_path):
            shutil.copy2(rtf_path, rtf_path + '.bak')

        if args.text:
            lines = args.text.split('\\n')
            line_end = '\r\n'
            paragraphs = []
            for line in lines:
                if line.strip():
                    paragraphs.append(f'<p class=MsoNormal>{line}<o:p></o:p></p>')
                else:
                    paragraphs.append(f'<p class=MsoNormal><o:p>&nbsp;</o:p></p>')
            body_html = line_end.join(paragraphs)

            # Preserve original HTML template (head/styles) if backup exists
            bak_path = htm_path + '.bak'
            if os.path.exists(bak_path):
                template = _read_sig_file(bak_path)
            else:
                template = _read_sig_file(htm_path)

            # Replace everything between <div> and </div> with new paragraphs
            import re as _re
            div_pattern = _re.compile(r'(<div[^>]*>).*?(</div>)', _re.DOTALL)
            match = div_pattern.search(template)
            if match:
                htm_content = template[:match.start(1)] + match.group(1) + line_end + body_html + line_end + match.group(2) + template[match.end(2):]
            else:
                htm_content = (
                    '<html xmlns:o="urn:schemas-microsoft-com:office:office"\r\n'
                    'xmlns:w="urn:schemas-microsoft-com:office:word"\r\n'
                    'xmlns="http://www.w3.org/TR/REC-html40">\r\n'
                    '<head><meta http-equiv=Content-Type content="text/html; charset=utf-8">\r\n'
                    '<meta name=Generator content="Microsoft Word 15"></head>\r\n'
                    f'<body>\r\n<div>\r\n\r\n{body_html}\r\n\r\n</div>\r\n</body>\r\n</html>\r\n'
                )

            with open(htm_path, 'w', encoding='utf-8') as f:
                f.write(htm_content)
            txt_content = '\r\n'.join(line if line.strip() else '' for line in lines)
            with open(txt_path, 'w', encoding='utf-16') as f:
                f.write(txt_content)

            # Update RTF
            rtf_ok = _rebuild_rtf_content(rtf_path, lines)

            print(f"✅ Signature '{name}' replaced from text.")
            print(f"   HTM: ✓ | TXT: ✓ | RTF: {'✓' if rtf_ok else '✗ (manual check needed)'}")
            print(f"   Backup: {htm_path}.bak")
            print(f"   ⚠️ Restart Outlook for changes to take effect.")
            for line in lines:
                print(f"   {line}")

        elif args.body:
            with open(htm_path, 'w', encoding='utf-8') as f:
                f.write(args.body)
            from html.parser import HTMLParser
            class _TextExtractor(HTMLParser):
                def __init__(self):
                    super().__init__()
                    self.result = []
                def handle_data(self, data):
                    self.result.append(data)
            extractor = _TextExtractor()
            extractor.feed(args.body)
            with open(txt_path, 'w', encoding='utf-16') as f:
                f.write(''.join(extractor.result))
            print(f"✅ Signature '{name}' replaced (full HTML body).")
            print(f"   ⚠️ RTF not updated — use --text mode for full update.")
            print(f"   Backup: {htm_path}.bak")
            print(f"   ⚠️ Restart Outlook for changes to take effect.")

        elif args.after and args.insert:
            htm_content = _read_sig_file(htm_path)

            if args.after not in htm_content:
                print(f"⚠️ Text '{args.after}' not found in signature '{name}'.")
                return 1

            anchor_idx = htm_content.find(args.after)
            end_p = htm_content.find('</p>', anchor_idx)
            if end_p == -1:
                insert_point = anchor_idx + len(args.after)
            else:
                insert_point = end_p + len('</p>')

            line_end = '\r\n' if '\r\n' in htm_content else '\n'
            new_paragraphs = ''
            for line in args.insert.split('\\n'):
                if line.strip():
                    new_paragraphs += f'{line_end}<p class=MsoNormal>{line}<o:p></o:p></p>'
                else:
                    new_paragraphs += f'{line_end}<p class=MsoNormal><o:p>&nbsp;</o:p></p>'

            htm_content = htm_content[:insert_point] + new_paragraphs + htm_content[insert_point:]
            with open(htm_path, 'w', encoding='utf-8') as f:
                f.write(htm_content)

            # Update .txt
            if os.path.exists(txt_path):
                txt_content = _read_sig_file(txt_path)
                txt_line_end = '\r\n' if '\r\n' in txt_content else '\n'
                txt_insert = txt_line_end + args.insert.replace('\\n', txt_line_end)
                anchor_idx_txt = txt_content.find(args.after)
                if anchor_idx_txt != -1:
                    eol = txt_content.find(txt_line_end, anchor_idx_txt)
                    if eol == -1:
                        eol = len(txt_content)
                    txt_content = txt_content[:eol] + txt_insert + txt_content[eol:]
                    with open(txt_path, 'w', encoding='utf-16') as f:
                        f.write(txt_content)

            # Update RTF — insert after anchor text
            if os.path.exists(rtf_path):
                rtf_content = open(rtf_path, 'rb').read().decode('latin-1')
                rtf_prefix = r'\hich\af31506\dbch\af31505\loch\f31506 '
                if args.after in rtf_content:
                    rtf_anchor_idx = rtf_content.find(args.after)
                    rtf_eol = rtf_content.find('\r\n', rtf_anchor_idx)
                    if rtf_eol == -1:
                        rtf_eol = len(rtf_content)
                    rtf_new_lines = ''
                    for line in args.insert.split('\\n'):
                        if line.strip():
                            rtf_new_lines += f'\r\n\\par {rtf_prefix}{_rtf_escape(line)}'
                        else:
                            rtf_new_lines += '\r\n\\par '
                    rtf_content = rtf_content[:rtf_eol] + rtf_new_lines + rtf_content[rtf_eol:]
                    with open(rtf_path, 'w', encoding='latin-1', newline='') as f:
                        f.write(rtf_content)

            insert_lines = args.insert.replace('\\n', '\n')
            print(f"✅ Signature '{name}' updated (inserted after '{args.after}').")
            print(f"   Added: {insert_lines}")
            print(f"   Backup: {htm_path}.bak")
            print(f"   ⚠️ Restart Outlook for changes to take effect.")

        elif args.find and args.replace is not None:
            htm_content = _read_sig_file(htm_path)

            if args.find not in htm_content:
                if os.path.exists(txt_path):
                    txt_content = _read_sig_file(txt_path)
                    if args.find not in txt_content:
                        print(f"⚠️ Text '{args.find}' not found in signature '{name}'.")
                        return 1

            count = htm_content.count(args.find)
            htm_content = htm_content.replace(args.find, args.replace)
            with open(htm_path, 'w', encoding='utf-8') as f:
                f.write(htm_content)

            if os.path.exists(txt_path):
                txt_content = _read_sig_file(txt_path)
                txt_content = txt_content.replace(args.find, args.replace)
                with open(txt_path, 'w', encoding='utf-16') as f:
                    f.write(txt_content)

            # Update RTF
            _update_rtf_signature(rtf_path, args.find, args.replace)

            print(f"✅ Signature '{name}' updated ({count} replacement(s)).")
            print(f"   '{args.find}' → '{args.replace}'")
            print(f"   Backup: {htm_path}.bak")
            print(f"   ⚠️ Restart Outlook for changes to take effect.")
        else:
            print("Error: provide --text, --body, --find/--replace, or --after/--insert.")
            return 1

        return 0
    except Exception as e:
        print(f"Error: {str(e)}", file=sys.stderr)
        return 1
