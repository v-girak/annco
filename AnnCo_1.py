import re
import wave
import xml.etree.ElementTree as ET
import xml.dom.minidom as MD
import tkinter as tk

from tkinter import ttk, messagebox
from tkinter.filedialog import askopenfilename, asksaveasfilename

# move to Application class
ENCOD_MSG = ("Кодування обраного файлу не підтримується. Будь ласка, збережіть"
             " файл у кодуванні UTF-8 або оберіть інший.")

ann_file = None
file_name = None


class Interval:
    def __init__(self, start, end, text):
        self.start = start
        self.end = end
        self.text = text 


class Tier:
    def __init__(self, name, ival_list):
        self.name = name
        self.intervals = ival_list


def open_file():
    file_path = askopenfilename(
        title='Оберіть вхідний файл',
        filetypes=[
            ('Файли Praat', '*.TextGrid'),
            ('Файли Elan', '*.eaf'),
            ('Файли Transcriber', '*.trs')
        ]
    )
    if not file_path:
        return
    global ann_file
    ann_file = open(file_path, encoding='UTF-8')
    global file_name
    file_name = re.search(r'[^/]+$', file_path).group(0)
    lbl_name['text'] = file_name


def incl_empty():
    if out_format.get() == 2:
        chk_empty.config(state='active')
    else:
        chk_empty.config(state='disabled')


def convert():
    if not lbl_name['text'] or not out_format.get():
        messagebox.showinfo(
            title="Чогось не вистачає...",
            message="Оберіть вхідний файл та кінцевий формат."
        )
        return

    else:
        lbl_name['text'] = ''
        if file_name.lower().endswith('.textgrid'):
            tier_objs, length = tg_to_obj(ann_file)
        elif file_name.lower().endswith('.eaf'):
            tier_objs, length = eaf_to_obj(ann_file)
        elif file_name.lower().endswith('.trs'):
            tier_objs, length = trs_to_obj(ann_file)

        if out_format.get() == 1:
            obj_to_tg(tier_objs, length)
        elif out_format.get() == 2:
            obj_to_eaf(tier_objs)


def tg_to_obj(tg_file):
    try:
        textgrid = tg_file.read()
    except UnicodeDecodeError:
        messagebox.showerror(
            title="Ой!",
            message=ENCOD_MSG
        )
        return
    tg_file.close()
    
    # compile these
    tier_re = r'name = "(.*?)"\s+'
    start_re = r'xmin = ([\d.]+)'
    end_re = r'xmax = ([\d.]+)'
    text_re = r'text = "(.*?)"\s+'

    tg_end = float(re.search(end_re, textgrid).group(1))

    # removes all tabs and strips header
    textgrid = re.sub(r'^\s+', '', textgrid, flags=re.M)
    textgrid = textgrid[textgrid.find('item [1]:'):]

    tg_tiers = re.split(r'item \[\d+\]:', textgrid)
    del tg_tiers[0]

    tier_objs = []
    for tier in tg_tiers:
        if not re.search(r"IntervalTier", tier):
            continue
        starts = [float(start) for start in re.findall(start_re, tier)]
        ends = [float(end) for end in re.findall(end_re, tier)]
        texts = [text.strip() for text in re.findall(text_re, tier)]
        tier_name = re.findall(tier_re, tier)
        del starts[0], ends[0]
        ival_tups = zip(starts, ends, texts)
        tier_intervals = [Interval(start, end, text) for (start, end, text) in ival_tups]
        tier_objs.append(Tier(tier_name[0], tier_intervals))

    return tier_objs, tg_end


def eaf_to_obj(eaf_file):
    try:
        ann_tree = ET.parse(eaf_file)
    except UnicodeDecodeError:
        messagebox.showerror(
            title="Ой!",
            message=ENCOD_MSG
        )
        return
    ann_doc = ann_tree.getroot()
    eaf_file.close()

    def get_times(ann_list):
        for ann in ann_list:
            for tier in ann_doc.findall("TIER"):
                ref_anns = tier.findall(
                    f'''.//*[@ANNOTATION_REF='{ann.get("ANNOTATION_ID")}']'''
                )
                if len(ref_anns):
                    ref_len = (
                        (ann.get("TIME_SLOT_REF2") - ann.get("TIME_SLOT_REF1"))
                        / len(ref_anns)
                    )
                    ref_point = ann.get("TIME_SLOT_REF1")
                    for ref in ref_anns:
                        ref.set("TIME_SLOT_REF1", ref_point)
                        ref_point += ref_len
                        ref.set("TIME_SLOT_REF2", ref_point)
                    get_times(ref_anns)

    # refactor into try/except
    if not ann_doc.find("HEADER/MEDIA_DESCRIPTOR"):
        wav_path = ann_doc.find("HEADER/MEDIA_DESCRIPTOR").get("MEDIA_URL")
        try:
            wav = wave.open(wav_path[8:], 'rb')
            duration = wav.getnframes() / wav.getframerate()
            wav.close()
        except:
            duration = ann_doc.find("TIME_ORDER")[-1].get("TIME_VALUE")
            duration = int(duration) / 1000
    else:
        duration = ann_doc.find("TIME_ORDER")[-1].get("TIME_VALUE")
        duration = int(duration) / 1000

    # swap loops
    for slot in ann_doc.find("TIME_ORDER"):
        for ann in ann_doc.iter("ALIGNABLE_ANNOTATION"):
            if slot.get("TIME_SLOT_ID") == ann.get("TIME_SLOT_REF1"):
                ann.set("TIME_SLOT_REF1", int(slot.get("TIME_VALUE"))/1000)
            if slot.get("TIME_SLOT_ID") == ann.get("TIME_SLOT_REF2"):
                ann.set("TIME_SLOT_REF2", int(slot.get("TIME_VALUE"))/1000)

    # integrate into upper block
    align_anns = ann_doc.findall(".//ALIGNABLE_ANNOTATION")
    get_times(align_anns)

    tier_objs = []
    for tier in ann_doc.findall("TIER"):
        tier_intervals = []
        for ann in tier.findall("ANNOTATION/*"):
            tier_intervals.append(
                Interval(
                    ann.get("TIME_SLOT_REF1"),
                    ann.get("TIME_SLOT_REF2"),
                    ann.find("*").text
                )
            )
        tier_objs.append(Tier(tier.get("TIER_ID"), tier_intervals))
    
    return tier_objs, duration


def trs_to_obj(trs_file):
    try:
        trans_tree = ET.parse(trs_file)
    except UnicodeDecodeError:
        messagebox.showerror(
            title='Ой!',
            message=ENCOD_MSG
        )
        return
    trans = trans_tree.getroot()
    trs_file.close()

    if trans.find('Topics'):
        for topic in trans.find('Topics'):
            for section in trans.iter('Section'):
                if topic.get('id') == section.get('topic'):
                    section.set('topic', topic.get('desc'))
    
    if trans.find('Speakers'):    
        for turn in trans.iter('Turn'):
            turn.set('speaker', turn.get('speaker').replace(' ', ' + '))
            for spk in trans.find('Speakers'):
                if spk.get('id') in turn.get('speaker'):
                    turn.set(
                        'speaker',
                        turn.get('speaker').replace(spk.get('id'), spk.get('name'))
                    )

    tier_objs = [Tier('Теми', []), Tier('Мовці', []), Tier('Транскрипція', [])]
    if trans.findall(".//Background"):
        tier_objs.append(Tier('Фон', []))

    for section in trans.find('Episode'):
        if section.get('topic'):
            tier_objs[0].intervals.append(
                Interval(
                    section.get('startTime'),
                    section.get('endTime'),
                    section.get('topic')
                )
            )
        else:
            tier_objs[0].intervals.append(
                Interval(
                    section.get('startTime'),
                    section.get('endTime'),
                    section.get('type')
                )
            )

        for turn in section:
            if turn.get('speaker'):
                tier_objs[1].intervals.append(
                    Interval(
                        turn.get('startTime'),
                        turn.get('endTime'),
                        turn.get('speaker')
                    )
                )
            else:
                tier_objs[1].intervals.append(
                    Interval(
                        turn.get('startTime'),
                        turn.get('endTime'),
                        '(без мовця)'
                    )
                )

            for el in turn:
                if el.tag == 'Sync':
                    tier_objs[2].intervals.append(
                        Interval(el.get('time'), '0', el.tail.strip())
                    )

                elif el.tag == 'Who':
                    tier_objs[2].intervals[-1].text \
                    += f" {el.get('nb')}: {el.tail.strip()}"

                elif el.tag == 'Comment':
                    tier_objs[2].intervals[-1].text \
                    += f" {{{el.get('desc')}}} {el.tail.strip()}"

                elif el.tag == 'Background':
                    tier_objs[2].intervals[-1].text += f" {el.tail.strip()}"
                    if el.get('level') == 'off':
                        tier_objs[3].intervals.append(
                            Interval(el.get('time'), '0', '')
                        )
                    else:
                        tier_objs[3].intervals.append(
                            Interval(el.get('time'), '0', el.get('type'))
                        )

                elif el.tag == 'Event':
                    desc, tail = el.get('desc'), el.tail.strip()
                    if el.get('extent') == "instantaneous":
                        tier_objs[2].intervals[-1].text += f" [{desc}] {tail}"
                    elif el.get('extent') == "begin":
                        tier_objs[2].intervals[-1].text += f" [{desc}-] {tail}"
                    elif el.get('extent') == "end":
                        tier_objs[2].intervals[-1].text += f" [-{desc}] {tail}"
                    elif el.get('extent') == "next":
                        tier_objs[2].intervals[-1].text += f" [{desc}]+ {tail}"
                    elif el.get('extent') == "previous":
                        tier_objs[2].intervals[-1].text += f" +[{desc}] {tail}"
            tier_objs[2].intervals[-1].end = turn.get('endTime')

    trans_end = float(tier_objs[0].intervals[-1].end)
    
    count = 0
    while count < len(tier_objs[2].intervals)-1:
        tier_objs[2].intervals[count].end = tier_objs[2].intervals[count+1].start
        count += 1
    
    if len(tier_objs) == 4:
        count = 0
        while count < len(tier_objs[3].intervals)-1:
            tier_objs[3].intervals[count].end = tier_objs[3].intervals[count+1].start
            count += 1
        tier_objs[3].intervals[-1].end = trans_end

    for tier in tier_objs:
        for interval in tier.intervals:
            interval.start = float(interval.start)
            interval.end = float(interval.end)
            interval.text = interval.text.strip()

    return tier_objs, trans_end


def obj_to_tg(tier_objs, length):
    tg_min = 0
    tg_max = length

    for tier in tier_objs:
        count = 0
        while count < len(tier.intervals)-1:
            prev_end = tier.intervals[count].end
            next_start = tier.intervals[count + 1].start
            if prev_end < next_start:
                tier.intervals.insert(count+1, Interval(prev_end, next_start, ''))
            count += 1

        if not tier.intervals:
            tier.intervals.append(Interval(tg_min, tg_max, ''))
        else:
            if tg_min != tier.intervals[0].start:
                tier.intervals.insert(0, Interval(tg_min, tier.intervals[0].start, ''))
            if tg_max != tier.intervals[-1].end:
                tier.intervals.append(Interval(tier.intervals[-1].end, tg_max, ''))

    filepath = asksaveasfilename(
        defaultextension='TextGrid',
        filetypes=[('Файли Praat', '*.TextGrid')]
    )
    if not filepath:
        return

    tg_file = open(filepath, 'w', encoding='UTF-8')

    tg_file.write(
        'File type = "ooTextFile"\n'
        'Object class = "TextGrid"\n\n'
        f'xmin = {tg_min}\n'
        f'xmax = {tg_max}\n'
        'tiers? <exists>\n'
        f'size = {len(tier_objs)}\n'
        'item []:\n'
    )

    tier_count = 0
    for tier in tier_objs:
        tier_count += 1
        tg_file.write(
            f'    item [{tier_count}]:\n'
            '        class = "IntervalTier"\n'
            f'        name = "{tier.name}"\n'
            f'        xmin = {tg_min}\n'
            f'        xmax = {tg_max}\n'
            f'        intervals: size = {len(tier.intervals)}\n'
        )

        ival_count = 0
        for interval in tier.intervals:
            ival_count += 1
            if interval.text == None:
                interval.text = ''
            tg_file.write(
                f'        intervals [{ival_count}]:\n'
                f'            xmin = {interval.start}\n'
                f'            xmax = {interval.end}\n'
                f'            text = "{interval.text}"\n'
            )

    tg_file.close()
    messagebox.showinfo(title='Ура!', message='Файл збережено!')


def obj_to_eaf(tier_objs, duration):
    ann_doc = ET.Element(
        'ANNOTATION_DOCUMENT',
        {'AUTHOR': '', 'FORMAT': '3.0', 'VERSION': '3.0',
         'xmlns:xsi': 'http://www.w3.org/2001/XMLSchema-instance',
         'xsi:noNamespaceSchemaLocation': 'http://www.mpi.nl/tools/elan/EAFv3.0.xsd'}
    )

    ann_tree = ET.ElementTree(ann_doc)

    header = ET.SubElement(
        ann_doc,
        'HEADER',
        {'MEDIA_FILE': '', 'TIME_UNITS': 'milliseconds'}
    )

    ET.SubElement(header, 'PROPERTY', {'NAME': 'URN'}).text \
        = 'urn:nl-mpi-tools-elan-eaf:187f732a-340c-4c9e-a8c3-307ba38799fb'
    ET.SubElement(header, 'PROPERTY', {'NAME': 'lastUsedAnnotationId'}).text = '0'

    time_order = ET.SubElement(ann_doc, 'TIME_ORDER')
    time_slots = []
    for tier in tier_objs:
        for interval in tier.intervals:
            if not interval.text:
                continue
            time_slots.append(int(round(interval.start,3)*1000))
            time_slots.append(int(round(interval.end,3)*1000))

    time_slots.sort()
    time_slot_count = 0
    for item in time_slots:
        time_slot_count += 1
        ET.SubElement(
            time_order,
            'TIME_SLOT',
            {'TIME_SLOT_ID': 'ts'+str(time_slot_count), 'TIME_VALUE': str(item)}
        )

    interval_count = 0
    for tier in tier_objs:
        tier_el = ET.SubElement(
            ann_doc,
            'TIER',
            {'LINGUISTIC_TYPE_REF': 'default-lt', 'TIER_ID': tier.name}
        )

        if cvar.get():
            for interval in tier.intervals:
                interval_count += 1
                ann_el = ET.SubElement(
                    ET.SubElement(tier_el, 'ANNOTATION'),
                    'ALIGNABLE_ANNOTATION',
                    {'ANNOTATION_ID': 'a'+str(interval_count),
                     'TIME_SLOT_REF1': str(int(round(interval.start,3)*1000)),
                     'TIME_SLOT_REF2': str(int(round(interval.end,3)*1000))}
                )
                ET.SubElement(ann_el, 'ANNOTATION_VALUE').text = interval.text

        else:
            for interval in tier.intervals:
                if not interval.text:
                    continue
                interval_count += 1
                ann_el = ET.SubElement(
                    ET.SubElement(tier_el, 'ANNOTATION'),
                    'ALIGNABLE_ANNOTATION',
                    {'ANNOTATION_ID': 'a'+str(interval_count),
                     'TIME_SLOT_REF1': str(int(round(interval.start,3)*1000)),
                     'TIME_SLOT_REF2': str(int(round(interval.end,3)*1000))}
                )
                ET.SubElement(ann_el, 'ANNOTATION_VALUE').text = interval.text

    for slot in time_order:
        for ann in ann_doc.iter('ALIGNABLE_ANNOTATION'):
            if slot.get('TIME_VALUE') == ann.get('TIME_SLOT_REF1'):
                ann.set('TIME_SLOT_REF1', slot.get('TIME_SLOT_ID'))
            elif slot.get('TIME_VALUE') == ann.get('TIME_SLOT_REF2'):
                ann.set('TIME_SLOT_REF2', slot.get('TIME_SLOT_ID'))

    ET.SubElement(
        ann_doc,
        'LINGUISTIC_TYPE',
        {'GRAPHIC_REFERENCES': 'false', 'LINGUISTIC_TYPE_ID': 'default-lt',
         'TIME_ALIGNABLE': 'true'}
    )

    ET.SubElement(
        ann_doc,
        'CONSTRAINT',
        {'DESCRIPTION': ("Time subdivision of parent annotation's time "
                         "interval, no time gaps allowed within this interval"),
         'STEREOTYPE': "Time_Subdivision"}
    )

    ET.SubElement(
        ann_doc,
        'CONSTRAINT',
        {'DESCRIPTION': ("Symbolic subdivision of a parent annotation. "
                         "Annotations refering to the same parent are ordered"),
         'STEREOTYPE': "Symbolic_Subdivision"}
    )

    ET.SubElement(
        ann_doc,
        'CONSTRAINT',
        {'DESCRIPTION': "1-1 association with a parent annotation",
         'STEREOTYPE': "Symbolic_Association"}
    )

    ET.SubElement(
        ann_doc,
        'CONSTRAINT',
        {'DESCRIPTION': ("Time alignable annotations within the parent "
                         "annotation's time interval, gaps are allowed"),
         'STEREOTYPE': "Included_In"}
    )

    filepath = asksaveasfilename(
        defaultextension='eaf',
        filetypes=[('Файли Elan', '*.eaf')]
    )
    if not filepath:
        return
    
    # get rid of minidom?
    ann_tree.write(filepath, 'UTF-8')
    ann_tree = MD.parse(filepath)
    ann_tree_str = ann_tree.toprettyxml()
    root = ET.fromstring(ann_tree_str)
    ann_tree = ET.ElementTree(root)
    ann_tree.write(filepath, 'UTF-8', xml_declaration=True)
    messagebox.showinfo(title='Ура!', message='Файл збережено!')


if __name__ == '__main__':

    window = tk.Tk()
    window.title('AnnCo v1.2')
    frm_body = tk.Frame(window)

    lbl_open = tk.Label(
        frm_body,
        text="1. Оберіть файл .TextGrid, .eaf або .trs у кодуванні UTF-8:"
    )
    frm_source = tk.Frame(frm_body)
    btn_open = ttk.Button(frm_source, text="Обрати файл...", command=open_file)
    lbl_name = tk.Label(frm_source, text="")

    lbl_dest = tk.Label(frm_body, text="2. Оберіть кінцевий формат:")
    frm_dest = tk.Frame(frm_body)
    out_format = tk.IntVar(frm_dest)
    rad_tg = ttk.Radiobutton(
        frm_dest,
        text=".TextGrid",
        value=1,
        variable=out_format,
        command=incl_empty
    )
    rad_eaf = ttk.Radiobutton(
        frm_dest,
        text=".eaf",
        value=2,
        variable=out_format,
        command=incl_empty
    )

    cvar = tk.BooleanVar(frm_dest)
    chk_empty = ttk.Checkbutton(
        frm_dest,
        text="Включити порожні інтервали до анотації",
        variable=cvar,
        onvalue=1,
        offvalue=0,
        state='disabled'
    )

    btn_convert = ttk.Button(
        window,
        text="Конвертувати",
        command=convert,
        default='active'
    )

    frm_body.columnconfigure(0, minsize=500)

    lbl_open.grid(row=0, column=0, sticky='w', padx=10, pady=10)
    btn_open.grid(row=0, column=0, ipadx=7, padx=3)
    lbl_name.grid(row=0, column=1, sticky='w')
    frm_source.grid(row=1, column=0, sticky='ew', padx=10)

    lbl_dest.grid(row=2, column=0, sticky='w', padx=10, pady=10)
    rad_tg.grid(row=0, column=0, sticky='w')
    rad_eaf.grid(row=1, column=0, sticky='w', pady=5)
    chk_empty.grid(row=1, column=1, padx=10)
    frm_dest.grid(row=3, column=0, sticky='ew', padx=10)

    frm_body.pack()

    btn_convert.pack(side=tk.RIGHT, padx=10, pady=10, ipadx=5)

    window.mainloop()