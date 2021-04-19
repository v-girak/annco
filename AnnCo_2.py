import re
import wave
import xml.etree.ElementTree as ET
import xml.dom.minidom as MD
import tkinter as tk

from tkinter import ttk, messagebox
from tkinter.filedialog import askopenfilename, asksaveasfilename


class Interval:
    """Represents annotation interval."""

    def __init__(self, start: float, end: float, text):
        self.start = start
        self.end = end
        self.text = text

    def __repr__(self):
        return f'Interval({self.start}, {self.end}, {self.text})'

    def __str__(self):
        return self.text

    def __len__(self):
        return self.end - self.start


class Tier:
    """Represents annotation tier containing its intervals."""

    def __init__(self, name, intervals=None):
        self.name = name
        if intervals is None:
            self.intervals = []
        else:
            self.intervals = intervals

    def __repr__(self):
        return f'Tier({self.name}, intervals)'

    def __len__(self):
        return len(self.intervals)


class Annotation:
    """Represents entire annotation."""

    def __init__(self, tiers, duration: float):
        self.tiers = tiers
        self.duration = duration
    
    @classmethod
    def from_tg(cls, contents):
        """Creates annotation from .TextGrid contents."""

        RE_TIER = re.compile(r'name = "(.*?)"\s+')
        RE_START = re.compile(r'xmin = ([\d.]+)')
        RE_END = re.compile(r'xmax = ([\d.]+)')
        RE_TEXT = re.compile(r'text = "(.*?)"\s+')

        duration = float(RE_END.search(textgrid).group(1))

        textgrid = textgrid[textgrid.find('item [1]:'):]
        tg_tiers = re.split(r'item \[\d+\]:', textgrid)[1:]

        tiers = []
        for t in tg_tiers:
            if not re.search(r"IntervalTier", tier):
                continue

            starts = [float(start) for start in RE_START.findall(t)][1:]
            ends = [float(end) for end in RE_END.findall(t)][1:]
            texts = [text.strip() for text in RE_TEXT.findall(t)]
            name = RE_TIER.search(t).group(1)
            tups = zip(starts, ends, texts)

            intervals = [Interval(start, end, text) for (start, end, text) in tups]
            tiers.append(Tier(name, intervals))

        return cls(tiers, duration)


class Converter:
    """Represents AnnCo interface."""
    
    RE_TXT = re.compile(r'\.textgrid$', re.I)
    RE_XML = re.compile(r'\.(eaf|trs)$', re.I)
    ENCOD_MSG = ("Кодування обраного файлу не підтримується. Будь ласка, "
                 "збережіть файл у кодуванні UTF-8 або оберіть інший.")

    def __init__(self):
        self.ann_file = None
        self.file_name = None

    def open_file(self):
        """Opens annotation file and returns its contents."""

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

        self.file_name = re.search(r'[^/]+$', file_path).group(0)

        if self.RE_TXT.search(self.file_name):
            self.file_format = 'txt'
        elif self.RE_XML.search(self.file_name):
            self.file_format = 'xml'

        # lbl_name['text'] = file_name

        with open(file_path, encoding='UTF-8') as ann_file:
            try:
                if self.file_format == 'txt':
                    contents = ann_file.read()
                elif self.file_format == 'xml':
                    contents = ET.parse(ann_file)
                return contents
            except UnicodeDecodeError:
                messagebox.showerror(title="Ой!", message=self.ENCOD_MSG)
                return