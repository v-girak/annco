import re
import wave
import xml.etree.ElementTree as ET
import xml.dom.minidom as MD
import tkinter as tk

from tkinter import ttk, messagebox
from tkinter.filedialog import askopenfilename, asksaveasfilename
from AnnCo_1 import obj_to_tg, obj_to_eaf


class Interval:
    """Represents annotation interval."""

    def __init__(self, start: float, end: float, text=None):
        self.start = start
        self.end = end
        if text is None:
            self.text = ''
        else:
            self.text = text

    def __repr__(self):
        return f'Interval({self.start}, {self.end}, {self.text})'

    def __str__(self):
        return self.text

    def __len__(self):
        return self.end - self.start

    def to_tg(self, i) -> str:
        "Returns a string representing interval in a .TextGrid file"

        tg_interval = (
            f"        intervals [{i}]:\n"
            f"            xmin = {self.start}\n"
            f"            xmax = {self.end}\n"
            f"            text = \"{self.text}\"\n"
        )

        return tg_interval


class Tier:
    """Represents annotation tier containing its intervals."""

    def __init__(self, name, intervals=None):
        self.name = name
        if intervals is None:
            self.intervals = []
        else:
            self.intervals = intervals
        self._index = 0

    def __repr__(self):
        return f'Tier({self.name}, intervals)'

    def __str__(self):
        return self.name

    def __len__(self):
        return len(self.intervals)

    def __iter__(self):
        self._index = 0
        return self

    def __next__(self):
        if self._index >= len(self.intervals):
            raise StopIteration
        i = self._index
        self._index += 1
        return self.intervals[i]

    def __getitem__(self, index):
        return self.intervals[index]

    def to_tg(self, t) -> str:
        "Returns a string representing tier in a .TextGrid file"

        tg_tier = (
            f"    item [{t}]:\n"
            "        class = \"IntervalTier\"\n"
            f"        name = \"{self.name}\"\n"
            f"        xmin = {self[0].start}\n"
            f"        xmax = {self[-1].end}\n"
            f"        intervals: size = {len(self)}\n"
        )

        for i, interval in enumerate(self, start=1):
            tg_tier += interval.to_tg(i)

        return tg_tier


class Annotation:
    """Represents entire annotation."""

    def __init__(self, tiers, duration: float):
        self.tiers = tiers
        self.duration = duration
        self._index = 0

    def __str__(self):
        return f"Annotation contains {len(self.tiers)} tiers."

    def __len__(self):
        return len(self.tiers)

    def __iter__(self):
        self._index = 0
        return self

    def __next__(self):
        if self._index >= len(self.tiers):
            raise StopIteration
        i = self._index
        self._index += 1
        return self.tiers[i]

    def __getitem__(self, index):
        return self.tiers[index]
    
    @classmethod
    def from_tg(cls, contents):
        """Creates Annotation instance from .TextGrid file contents."""

        RE_TIER = re.compile(r'name = "(.*?)"\s+')
        RE_START = re.compile(r'xmin = ([\d.]+)')
        RE_END = re.compile(r'xmax = ([\d.]+)')
        RE_TEXT = re.compile(r'text = "(.*?)"\s+', re.S)

        textgrid = contents[contents.find('item [1]:'):]
        tg_tiers = re.split(r'item \[\d+\]:', textgrid)[1:]

        duration = float(RE_END.search(textgrid).group(1))

        tiers = []
        for t in tg_tiers:
            if not re.search(r"IntervalTier", t):
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
        """Creates Annotation instance from .eaf file contents."""

        ann_doc = contents.getroot()
        duration = cls._get_duration(ann_doc)

        align_anns = ann_doc.findall('.//ALIGNABLE_ANNOTATION')
        cls._insert_align_ann_times(ann_doc, align_anns)
        cls._insert_ref_ann_times(ann_doc, align_anns)

        tiers = cls._get_tiers(ann_doc)

        return cls(tiers, duration)

    @classmethod
    def from_trs(cls, contents):
        """Creates Annotation instance from .trs file contents."""

        trans = contents.getroot()
        cls._insert_topics(trans)
        cls._insert_speakers(trans)

        sections = cls._get_sections(trans)
        turns = cls._get_turns(trans)
        transcription, background = cls._get_transcription(trans)

        duration = sections[-1].end

        cls._set_ends(transcription, duration)
        cls._set_ends(background, duration)

        tiers = [Tier('Теми', sections), Tier('Мовці', turns),
                 Tier('Транскрипція', transcription)]
        if background:
            tiers.append(Tier('Фон', background))

        return cls(tiers, duration)

    @staticmethod
    def _get_duration(root) -> float:
        """Gets annotation duration from .eaf file root.
        
        Extracts duration from a media file. If media file is absent,
        sets value of the last time slot as duration.
        """

        try:
            wav_path = root.find('HEADER/MEDIA_DESCRIPTOR').get('MEDIA_URL')
            with wave.open(wav_path[8:], 'rb') as wav:
                duration = wav.getnframes() / wav.getframerate()
        except:
            last_time = int(root.find('TIME_ORDER')[-1].get('TIME_VALUE')) / 1000
            duration = last_time if last_time > 300.0 else 300.0

        return duration

    @staticmethod
    def _insert_align_ann_times(root, annotations) -> None:
        """Replaces .eaf alignable annotations' time references with times."""

        for ann in annotations:
            for slot in root.find('TIME_ORDER'):

                if ann.get('TIME_SLOT_REF1') == slot.get('TIME_SLOT_ID'):
                    ann.set('TIME_SLOT_REF1', int(slot.get('TIME_VALUE')) / 1000)

                if ann.get('TIME_SLOT_REF2') == slot.get('TIME_SLOT_ID'):
                    ann.set('TIME_SLOT_REF2', int(slot.get('TIME_VALUE')) / 1000)

        return

    @staticmethod
    def _insert_ref_ann_times(root, annotations) -> None:
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
                        ref_time += ref_dur
                        ref.set("TIME_SLOT_REF2", ref_time)

                    Annotation._insert_ref_ann_times(root, ref_anns)

        return

    @staticmethod
    def _get_tiers(root) -> list:
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
    def _insert_topics(root) -> None:
        """Sets sections' topics to descriptions in .trs file root."""

        if root.find('Topics'):
            for sect in root.iter('Section'):
                for topic in root.find('Topics'):
                    if sect.get('topic') == topic.get('id'):
                        sect.set('topic', topic.get('desc'))

        return

    @staticmethod
    def _insert_speakers(root) -> None:
        """Sets turns' speakers to names in .trs file root."""

        if root.find('Speakers'):    
            for turn in root.iter('Turn'):
                turn.set('speaker', turn.get('speaker').replace(' ', ' + '))
                for spk in root.find('Speakers'):
                    if spk.get('id') in turn.get('speaker'):
                        turn.set(
                            'speaker',
                            turn.get('speaker').replace(spk.get('id'),
                                                        spk.get('name'))
                        )

        return

    @staticmethod
    def _get_sections(root) -> list:
        """Returns list of Interval instances for sections from .trs file root."""

        sections = []
        for sect in root.iter('Section'):
            start = float(sect.get('startTime'))
            end = float(sect.get('endTime'))
            text = sect.get('topic') if sect.get('topic') else sect.get('type')

            sections.append(Interval(start, end, text))

        return sections

    @staticmethod
    def _get_turns(root) -> list:
        """Returns list of Interval instances for turns from .trs file root."""

        turns = []
        for turn in root.iter('Turn'):
            start = float(turn.get('startTime'))
            end = float(turn.get('endTime'))
            text = turn.get('speaker') if turn.get('speaker') else '(без мовця)'

            turns.append(Interval(start, end, text))

        return turns

    @staticmethod
    def _get_transcription(root) -> list:
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
                text = '' if el.get('level') == 'off' else el.get('type')
                background.append(Interval(start, end, text))

            elif el.tag == 'Event':
                desc, text = el.get('desc'), el.tail.strip()
                if el.get('extent') == 'instantaneous':
                    transcription[-1].text += f" [{desc}] {text}"
                if el.get('extent') == 'begin':
                    transcription[-1].text += f" [{desc}-] {text}"
                if el.get('extent') == 'end':
                    transcription[-1].text += f" [-{desc}] {text}"
                if el.get('extent') == 'next':
                    transcription[-1].text += f" [{desc}]+ {text}"
                if el.get('extent') == 'previous':
                    transcription[-1].text += f" +[{desc}] {text}"

        # if base interval text was empty & formatted str was appended:
        for interval in transcription:
            interval.text = interval.text.strip()

        return transcription, background

    @staticmethod
    def _set_ends(intervals, duration) -> None:
        """Sets ends for intervals."""

        i = 0
        while i < len(intervals) - 1:
            intervals[i].end = intervals[i+1].start
            i += 1
        if intervals: intervals[i].end = duration

        return

    def fill_spaces(self) -> None:
        "Fills empty spaces between intervals and tier boundaries on all tiers."

        for tier in self:
            for i in range(len(tier)-1):
                if tier[i].end < tier[i+1].start:
                    tier.intervals.insert(
                        i+1, Interval(tier[i].end, tier[i+1].start, '')
                    )

            if tier.intervals:
                if tier[0].start > 0:
                    tier.intervals.insert(0, Interval(0, tier[0].start, ''))
                if tier[-1].end < self.duration:
                    tier.intervals.append(Interval(tier[-1].end, self.duration, ''))

        return

    def to_tg(self) -> str:
        "Returns a string representing Annotation to be written into .TextGrid"

        self.fill_spaces()

        tg_ann = (
            "File type = \"ooTextFile\"\n"
            "Object class = \"TextGrid\"\n\n"
            f"xmin = {0}\n"
            f"xmax = {self.duration}\n"
            "tiers? <exists>\n"
            f"size = {len(self)}\n"
            "item []:\n"
        )

        for t, tier in enumerate(self, start=1):
            tg_ann += tier.to_tg(t)

        return tg_ann


class Converter:
    """Represents AnnCo interface."""
    
    RE_TXT = re.compile(r'\.textgrid$', re.I)
    RE_XML = re.compile(r'\.(eaf|trs)$', re.I)
    ENCOD_MSG = ("Кодування обраного файлу не підтримується. Будь ласка, "
                 "збережіть файл у кодуванні UTF-8 або оберіть інший.")

    def __init__(self):
        self.ann_file = None
        self.file_name = None
        self.file_format = None

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

    def test_convert(self, ext):
        """Test function to check functionality."""

        contents = self.open_file()

        if ext == 'tg':
            ann = Annotation.from_tg(contents)
        elif ext == 'eaf':
            ann = Annotation.from_eaf(contents)
        elif ext == 'trs':
            ann = Annotation.from_trs(contents)

        tg_ann = ann.to_tg()
    
        filepath = asksaveasfilename(
            defaultextension='TextGrid',
            filetypes=[('Файли Praat', '*.TextGrid')]
        )

        with open(filepath, 'w', encoding='UTF-8') as tg_file:
            tg_file.write(tg_ann) 

        messagebox.showinfo(title='Ура!', message='Файл збережено!')
        # obj_to_eaf(ann.tiers, ann.duration)


if __name__ == '__main__':

    annco = Converter()
    # annco.test_convert('tg')
    # annco.test_convert('eaf')
    annco.test_convert('trs')