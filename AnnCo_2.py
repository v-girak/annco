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

    @classmethod
    def from_eaf(cls, contents):
        """Creates annotation from .eaf file contents."""

        ann_doc = contents.getroot()
        duration = cls.get_duration(ann_doc)

        align_anns = ann_doc.findall('.//ALIGNABLE_ANNOTATION')
        ann_doc = cls.insert_align_ann_times(ann_doc, align_anns)
        ann_doc = cls.insert_ref_ann_times(ann_doc, align_anns)

        tiers = cls.get_tiers(ann_doc)

        return cls(tiers, duration)

    @classmethod
    def from_trs(cls, contents):
        """Creates annotation from .trs file contents."""

        trans = contents.getroot()
        # insert sections topics
        # insert turns speakers
        # get sections
        # get turns
        # get transcription & background
        # set_ends(transcription)
        # set_ends(background)
        # create Tiers from lists above
        # create Annotation from tiers & duration

    @staticmethod
    def insert_ref_ann_times(root, annotations) -> ET.Element:
        """Assigns time boundaries to referring annotations."""

        for ann in annotations:
            for t in root.findall('TIER'):
                ref_anns = t.findall(
                    f".//*[@ANNOTATION_REF='{ann.get('ANNOTATION_ID')}']"
                )
                if ref_anns:
                    ref_dur = (
                        (ann.get('TIME_SLOT_REF2') - ann.get('TIME_SLOT_REF1'))
                        / len(ref_anns)
                    )
                    ref_time = ann.get('TIME_SLOT_REF1')

                    for ref in ref_anns:
                        ref.set('TIME_SLOT_REF1', ref_time)
                        ref_time += ref_len
                        ref.set("TIME_SLOT_REF2", ref_time)

                    insert_ref_ann_times(root, ref_anns)

        return root

    @staticmethod
    def get_duration(root) -> float:
        """Gets annotation duration from .eaf file root.
        
        Extracts duration from a media file. If media file is absent,
        sets value of the last time slot as duration.
        """

        try:
            wav_path = root.find('HEADER/MEDIA_DESCRIPTOR').get('MEDIA_URL')
            with wave.open(wav_path[8:], 'rb') as wav:
                duration = wav.getnframes() / wav.getframerate()
        except:
            duration = int(root.find('TIME_ORDER')[-1].get('TIME_VALUE')) / 1000

        return duration

    @staticmethod
    def insert_align_ann_times(root, annotations) -> ET.Element:
        """Replaces .eaf alignable annotations' time references with times."""

        for ann in annotations:
            for slot in root.find('TIME_ORDER'):

                if ann.get('TIME_SLOT_REF1') == slot.get('TIME_SLOT_ID'):
                    ann.set('TIME_SLOT_REF1', int(slot.get('TIME_VALUE')) / 1000)

                if ann.get('TIME_SLOT_REF2') == slot.get('TIME_SLOT_ID'):
                    ann.set('TIME_SLOT_REF2', int(slot.get('TIME_VALUE')) / 1000)

        return root

    @staticmethod
    def get_tiers(root) -> list:
        """Returns tiers and their intervals from .eaf file root."""

        tiers = []
        for t in root.findall('TIER'):
            name = t.get('TIER_ID')
            intervals = []

            for ann in t.findall('ANNOTATION/*'):
                start = ann.get('TIME_SLOT_REF1')
                end = ann.get('TIME_SLOT_REF2')
                text = ann.find("*").text

                intervals.append(Interval(start, end, text))

            tiers.append(Tier(name, intervals))

        return tiers

    @staticmethod
    def insert_topics(root) -> ET.Element:
        """Sets sections' topics to descriptions in .trs file root."""

        if root.find('Topics'):
            for sect in trans.iter('Section'):
                for topic in root.find('Topics'):
                    if sect.get('topic') == topic.get('id'):
                        sect.set('topic', topic.get('desc'))

        return root

    @staticmethod
    def insert_speakers(root) -> ET.Element:
        """Sets turns' speakers to names in .trs file root."""

        if trans.find('Speakers'):    
            for turn in trans.iter('Turn'):
                turn.set('speaker', turn.get('speaker').replace(' ', ' + '))
                for spk in trans.find('Speakers'):
                    if spk.get('id') in turn.get('speaker'):
                        turn.set(
                            'speaker',
                            turn.get('speaker').replace(spk.get('id'),
                                                        spk.get('name'))
                        )

        return root

    @staticmethod
    def get_sections(root) -> list:
        """Returns list of Interval instances for sections from .trs file root."""

        sections = []
        for sect in root.iter('Section'):
            start = float(sect.get('startTime'))
            end = float(sect.get('endTime'))
            text = sect.get('topic') if sect.get('topic') else sect.get('type')

            sections.append(Interval(start, end, text))

        return sections

    @staticmethod
    def get_turns(root) -> list:
        """Returns list of Interval instances for turns from .trs file root."""

        turns = []
        for turn in root.iter('Turn'):
            start = float(turn.get('startTime'))
            end = float(turn.get('endTime'))
            text = turn.get('speaker') if turn.get('speaker') else '(без мовця)'

            turns.append(Interval(start, end, text))

        return turns

    @staticmethod
    def get_transcription(root) -> list:
        """Returns list of Interval instances for transcription and background
        from .trs file root.
        """

        transcription = []
        background = []

        for el in root.findall('.//Turn/*'):
            if el.tag == 'Sync':
                start = float(el.get('time'))
                end = 0.0
                text = el.tail.strip()
                transcription.append(Interval(start, end, text))

            elif el.tag == 'Who':
                nb = el.get('nb')
                text = el.tail.strip()
                transcription[-1].text += f" {nb}: {text}"

            elif el.tag == 'Comment':
                desc = el.get('desc')
                text = el.tail.strip()
                transcription[-1].text += f" {{{desc}}} {text}"

            elif el.tag == 'Background':
                text = el.tail.strip()
                transcription[-1].text += f" {text}"
                start = float(el.get('time'))
                end = 0.0
                text = el.get('type') if el.get('level') == 'off' else ''
                background.append(Interval(start, end, text))

            elif el.tag == 'Event':
                desc, text = el.get('desc'), el.tail.strip()
                if el.get('extent') == 'instantaneous':
                    transcription[-1].text += f" [{desc}] {tail}"
                if el.get('extent') == 'begin':
                    transcription[-1].text += f" [{desc}-] {tail}"
                if el.get('extent') == 'end':
                    transcription[-1].text += f" [-{desc}] {tail}"
                if el.get('extent') == 'next':
                    transcription[-1].text += f" [{desc}]+ {tail}"
                if el.get('extent') == 'previous':
                    transcription[-1].text += f" +[{desc}] {tail}"

            # set_ends(transcription)
            # set_ends(background)

            return transcription, background


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