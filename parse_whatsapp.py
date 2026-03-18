import re
import os
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill
from openpyxl.utils import get_column_letter

# Regex for format 1: [2026-01-05, 09:35:42] Sender: Message
RE_FMT1 = re.compile(r'^\u200e?\[(\d{4}-\d{2}-\d{2}, \d{2}:\d{2}:\d{2})\] ([^:]+): (.*)')
RE_FMT1_SYS = re.compile(r'^\u200e?\[(\d{4}-\d{2}-\d{2}, \d{2}:\d{2}:\d{2})\] ([^:]+)$')

# Regex for format 2: 2026/02/20, 20:37 - Sender: Message
RE_FMT2 = re.compile(r'^(\d{4}/\d{2}/\d{2}, \d{2}:\d{2}) - ([^:]+): (.*)')
RE_FMT2_SYS = re.compile(r'^(\d{4}/\d{2}/\d{2}, \d{2}:\d{2}) - (.+)$')

# Attachment patterns
RE_ATTACH_FMT1 = re.compile(r'<attached: ([^>]+)>')
RE_ATTACH_FMT2 = re.compile(r'^(.+) \(file attached\)$')


def detect_format(lines):
    for line in lines[:20]:
        if RE_FMT1.match(line) or RE_FMT1_SYS.match(line):
            return 1
        if RE_FMT2.match(line) or RE_FMT2_SYS.match(line):
            return 2
    return 2


def is_new_message(line, fmt):
    if fmt == 1:
        return bool(RE_FMT1.match(line) or RE_FMT1_SYS.match(line))
    else:
        return bool(RE_FMT2.match(line) or RE_FMT2_SYS.match(line))


def parse_chat(filepath):
    with open(filepath, encoding='utf-8') as f:
        lines = f.readlines()

    fmt = detect_format(lines)
    messages = []
    current = None

    for raw_line in lines:
        line = raw_line.rstrip('\n')
        # Strip zero-width non-breaking space and left-to-right mark
        line = line.lstrip('\ufeff\u200e\u202a\u202c')

        if is_new_message(line, fmt):
            if current:
                messages.append(current)

            if fmt == 1:
                m = RE_FMT1.match(line) or RE_FMT1.match(line.lstrip('\u200e'))
                if not m:
                    # system message
                    ms = RE_FMT1_SYS.match(line)
                    if ms:
                        current = {'datetime': ms.group(1), 'sender': '', 'text': ms.group(2), 'attachments': []}
                    else:
                        current = None
                    continue
                current = {
                    'datetime': m.group(1),
                    'sender': m.group(2).strip(),
                    'text': m.group(3),
                    'attachments': []
                }
            else:
                m = RE_FMT2.match(line)
                if not m:
                    ms = RE_FMT2_SYS.match(line)
                    if ms:
                        current = {'datetime': ms.group(1), 'sender': '', 'text': ms.group(2), 'attachments': []}
                    else:
                        current = None
                    continue
                current = {
                    'datetime': m.group(1),
                    'sender': m.group(2).strip(),
                    'text': m.group(3),
                    'attachments': []
                }
        else:
            # continuation line
            if current is not None:
                current['text'] += '\n' + line

    if current:
        messages.append(current)

    # Post-process: extract attachments
    for msg in messages:
        text = msg['text']
        attachments = []

        if fmt == 1:
            # Extract <attached: filename> references
            found = RE_ATTACH_FMT1.findall(text)
            attachments.extend(found)
            # Clean attachment tags from text
            text = RE_ATTACH_FMT1.sub('', text).strip()
        else:
            # Check if entire message is "filename (file attached)"
            m = RE_ATTACH_FMT2.match(text.strip())
            if m:
                attachments.append(m.group(1))
                text = ''
            else:
                # Could have mixed text + attachment on same line? Unlikely but handle
                pass

        msg['text'] = text
        msg['attachments'] = attachments

    return messages


def create_xlsx(messages, output_path, sheet_name='Chat'):
    wb = Workbook()
    ws = wb.active
    ws.title = sheet_name[:31]

    # Header style
    header_font = Font(name='Arial', bold=True, color='FFFFFF')
    header_fill = PatternFill('solid', start_color='2F5496')
    header_align = Alignment(horizontal='center', vertical='center', wrap_text=True)

    headers = ['Date & Time', 'Sender', 'Message', 'Attachments']
    col_widths = [22, 28, 80, 50]

    for col, (h, w) in enumerate(zip(headers, col_widths), 1):
        cell = ws.cell(row=1, column=col, value=h)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = header_align
        ws.column_dimensions[get_column_letter(col)].width = w

    ws.row_dimensions[1].height = 20

    # Data style
    data_font = Font(name='Arial', size=10)
    data_align_wrap = Alignment(vertical='top', wrap_text=True)
    data_align_nowrap = Alignment(vertical='top', wrap_text=False)

    for row_idx, msg in enumerate(messages, 2):
        dt = msg['datetime']
        sender = msg['sender']
        text = msg['text']
        attachments = ', '.join(msg['attachments']) if msg['attachments'] else ''

        ws.cell(row=row_idx, column=1, value=dt).font = data_font
        ws.cell(row=row_idx, column=1).alignment = data_align_nowrap

        ws.cell(row=row_idx, column=2, value=sender).font = data_font
        ws.cell(row=row_idx, column=2).alignment = data_align_nowrap

        ws.cell(row=row_idx, column=3, value=text).font = data_font
        ws.cell(row=row_idx, column=3).alignment = data_align_wrap

        ws.cell(row=row_idx, column=4, value=attachments).font = data_font
        ws.cell(row=row_idx, column=4).alignment = data_align_wrap

    # Freeze header row
    ws.freeze_panes = 'A2'

    wb.save(output_path)
    print(f'Saved: {output_path} ({len(messages)} messages)')


BASE = os.path.dirname(os.path.abspath(__file__))

chats = [
    {
        'txt': os.path.join(BASE, 'WhatsApp Chat - Graad 2_8', '_chat.txt'),
        'out': os.path.join(BASE, 'WhatsApp Chat - Graad 2_8.xlsx'),
        'name': 'Graad 2_8'
    },
    {
        'txt': os.path.join(BASE, 'WhatsApp Chat with Gr 8k 2026', 'WhatsApp Chat with Gr 8k 2026.txt'),
        'out': os.path.join(BASE, 'WhatsApp Chat - Gr 8k 2026.xlsx'),
        'name': 'Gr 8k 2026'
    },
    {
        'txt': os.path.join(BASE, 'WhatsApp Chat with Vodacom Xander', 'WhatsApp Chat with Vodacom Xander.txt'),
        'out': os.path.join(BASE, 'WhatsApp Chat - Vodacom Xander.xlsx'),
        'name': 'Vodacom Xander'
    },
]

for chat in chats:
    messages = parse_chat(chat['txt'])
    create_xlsx(messages, chat['out'], sheet_name=chat['name'])
