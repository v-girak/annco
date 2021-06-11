# For license, see LICENSE.txt

import re
import wave
import random
import xml.etree.ElementTree as ET
import tkinter as tk

from tkinter import ttk, messagebox
from tkinter.filedialog import askopenfilenames, asksaveasfilename


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

    @property
    def eaf_start(self) -> int:
        "Returns interval start value formatted for .eaf"

        return int(round(self.start, 3) * 1000)

    @property
    def eaf_end(self) -> int:
        "Returns interval end value formatted for .eaf"

        return int(round(self.end, 3) * 1000)

    @property
    def antx_start(self) -> str:
        """Returns interval start value formatted for .antx"""

        return str(self.start * 44100)
    
    @property
    def antx_dur(self) -> str:
        """Returns interval duration value formatted for .antx"""

        return str(44100 * (self.end - self.start))

    def to_tg(self, i) -> str:
        "Returns a string representing interval in .TextGrid file"

        if self.start != self.end:
            tg_interval = (
                f"        intervals [{i}]:\n"
                f"            xmin = {self.start}\n"
                f"            xmax = {self.end}\n"
                f"            text = \"{self.text}\"\n"
            )
        else:
            tg_interval = (
                f"        points [{i}]:\n"
                f"            number = {self.start}\n"
                f"            mark = \"{self.text}\"\n"
            )

        return tg_interval

    def to_eaf(self, i, tier_el) -> None:
        "Creates ANNOTATION element representing interval in .eaf file"

        ann_el = ET.SubElement(tier_el, 'ANNOTATION')

        align_ann = ET.SubElement(ann_el, 'ALIGNABLE_ANNOTATION',
                                  {'ANNOTATION_ID': 'a' + str(i),
                                   'TIME_SLOT_REF1': str(self.eaf_start),
                                   'TIME_SLOT_REF2': str(self.eaf_end)})

        ET.SubElement(align_ann, 'ANNOTATION_VALUE').text = self.text

    def to_antx(self, root, segment_id: str, layer_id: str) -> None:
        """Creates Segment element representing interval in .antx file."""

        segment = ET.SubElement(root, 'Segment')

        id_el = ET.SubElement(segment, 'Id')
        id_el.text = segment_id

        layer_id_el = ET.SubElement(segment, 'IdLayer')
        layer_id_el.text = layer_id

        label = ET.SubElement(segment, 'Label')
        label.text = self.text

        fore_color = ET.SubElement(segment, 'ForeColor')
        fore_color.text = '-16777216'

        back_color = ET.SubElement(segment, 'BackColor')
        back_color.text = '-1'

        border_color = ET.SubElement(segment, 'BorderColor')
        border_color.text = '-16777216'

        start = ET.SubElement(segment, 'Start')
        start.text = self.antx_start

        duration = ET.SubElement(segment, 'Duration')
        duration.text = self.antx_dur

        is_sel = ET.SubElement(segment, 'IsSelected')
        is_sel.text = 'false'

        feat = ET.SubElement(segment, 'Feature')
        lang = ET.SubElement(segment, 'Language')
        group = ET.SubElement(segment, 'Group')
        name = ET.SubElement(segment, 'Name')
        param_1 = ET.SubElement(segment, 'Parameter1')
        param_2 = ET.SubElement(segment, 'Parameter2')
        param_3 = ET.SubElement(segment, 'Parameter3')

        is_marker = ET.SubElement(segment, 'IsMarker')
        is_marker.text = 'false'

        marker = ET.SubElement(segment, 'Marker')
        r_script = ET.SubElement(segment, 'RScript')

        vid_off = ET.SubElement(segment, 'VideoOffset')
        vid_off.text = '0'


class Tier:
    """Represents annotation tier containing its intervals."""

    def __init__(self, name, intervals=None, is_point=False):
        self.name = name
        if intervals is None:
            self.intervals = []
        else:
            self.intervals = intervals
        self.is_point = is_point
        self._antx_id = None
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

    def extend_points(self, duration) -> None:
        """If the tier is not empty, extends intervals ends
        to starts of intervals following them or to duration"""

        if self.intervals:
            for i in range(len(self) - 1):
                self[i].end = self[i+1].start
            self[-1].end = duration

    def fill_gaps(self, duration) -> None:
        "Fills gaps between intervals and tier boundaries with empty text intervals"

        if self.intervals:

            for i in range(len(self) - 1):
                if self[i].end < self[i+1].start:
                    self.intervals.insert(
                        i+1, Interval(self[i].end, self[i+1].start)
                    )

            if self[0].start > 0:
                self.intervals.insert(0, Interval(0, self[0].start))
            if self[-1].end < duration:
                self.intervals.append(Interval(self[-1].end, duration))

        else:
            self.intervals.append(Interval(0, duration))

    def to_tg(self, t, end) -> str:
        "Returns a string representing tier in a .TextGrid file"

        if not self.is_point:
            tg_tier = (
                f"    item [{t}]:\n"
                "        class = \"IntervalTier\"\n"
                f"        name = \"{self.name}\"\n"
                f"        xmin = 0\n"
                f"        xmax = {end}\n"
                f"        intervals: size = {len(self)}\n"
            )
        else:
            tg_tier = (
                f"    item [{t}]:\n"
                "        class = \"TextTier\"\n"
                f"        name = \"{self.name}\"\n"
                f"        xmin = 0\n"
                f"        xmax = {end}\n"
                f"        points: size = {len(self)}\n"
            )

        for i, interval in enumerate(self, start=1):
            tg_tier += interval.to_tg(i)

        return tg_tier

    def to_eaf(self, root) -> None:
        "Creates TIER element representing tier in .eaf file"

        tier_el = ET.SubElement(root, 'TIER', {'LINGUISTIC_TYPE_REF': 'default-lt',
                                               'TIER_ID': self.name})

        if interface.body.output_frame.incl_empty_var.get():
            for i, interval in enumerate(self, start=1):
                interval.to_eaf(i, tier_el)
        else:
            for i, interval in enumerate(self, start=1):
                if not interval.text:
                    continue
                interval.to_eaf(i, tier_el)

    def to_antx(self, root, layer_id: str) -> None:
        """Creates Layer element representing tier in .antx file."""

        self._antx_id = layer_id
        layer = ET.SubElement(root, 'Layer')

        id_el = ET.SubElement(layer, 'Id')
        id_el.text = layer_id

        name = ET.SubElement(layer, 'Name')
        name.text = self.name

        forecolor = ET.SubElement(layer, 'ForeColor')
        forecolor.text = '-16777216'

        backcolor = ET.SubElement(layer, 'BackColor')
        backcolor.text = '-1'

        is_sel = ET.SubElement(layer, 'IsSelected')
        is_sel.text = 'false'

        height = ET.SubElement(layer, 'Height')
        height.text = '70'

        ccs = ET.SubElement(layer, 'CoordinateControlStyle')
        ccs.text = '0'

        is_locked = ET.SubElement(layer, 'IsLocked')
        is_locked.text = 'false'

        is_closed = ET.SubElement(layer, 'IsClosed')
        is_closed.text = 'false'

        sos = ET.SubElement(layer, 'ShowOnSpectrogram')
        sos.text = 'false'

        sac = ET.SubElement(layer, 'ShowAsChart')
        sac.text = 'false'

        chart_min = ET.SubElement(layer, 'ChartMinimum')
        chart_min.text = '-50'

        chart_max = ET.SubElement(layer, 'ChartMaximum')
        chart_max.text = '50'

        show_bounds = ET.SubElement(layer, 'ShowBoundaries')
        show_bounds.text = 'true'

        iif = ET.SubElement(layer, 'IncludeInFrequency')
        iif.text = 'true'

        param_1 = ET.SubElement(layer, 'Parameter1Name')
        param_1.text = 'Parameter 1'

        param_2 = ET.SubElement(layer, 'Parameter2Name')
        param_2.text = 'Parameter 2'

        param_3 = ET.SubElement(layer, 'Parameter3Name')
        param_3.text = 'Parameter 3'

        is_vis = ET.SubElement(layer, 'IsVisible')
        is_vis.text = 'true'

        font_size = ET.SubElement(layer, 'FontSize')
        font_size.text = '10'

        vpi = ET.SubElement(layer, 'VideoPlayerIndex')
        vpi.text = '0'


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

        RE_NAME = re.compile(r'name = "(.*?)"\s+')
        RE_XMIN = re.compile(r'xmin = ([\d.]+)')
        RE_XMAX = re.compile(r'xmax = ([\d.]+)')
        RE_NUMB = re.compile(r'number = ([\d.]+)')
        RE_TEXT = re.compile(r'text = "(.*?)"\s+', re.S)
        RE_MARK = re.compile(r'mark = "(.*?)"\s+', re.S)

        textgrid = contents[contents.find('item [1]:'):]
        tg_tiers = re.split(r'item \[\d+\]:', textgrid)[1:]

        duration = float(RE_XMAX.search(textgrid).group(1))

        tiers = []
        for t in tg_tiers:
            if re.search(r"IntervalTier", t):
                starts = [float(start) for start in RE_XMIN.findall(t)][1:]
                ends = [float(end) for end in RE_XMAX.findall(t)][1:]
                texts = [text.strip() for text in RE_TEXT.findall(t)]
                name = RE_NAME.search(t).group(1)
                tups = zip(starts, ends, texts)

                intervals = [Interval(start, end, text) for (start, end, text) in tups]
                tiers.append(Tier(name, intervals))

            elif re.search(r"TextTier", t):
                times = [float(time) for time in RE_NUMB.findall(t)]
                texts = [text.strip() for text in RE_MARK.findall(t)]
                name = RE_NAME.search(t).group(1)
                tups = zip(times, texts)

                intervals = [Interval(time, time, text) for time, text in tups]
                tiers.append(Tier(name, intervals, is_point=True))

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

        tiers = [
            Tier('Теми', sections),
            Tier('Мовці', turns),
            Tier('Транскрипція', transcription)
        ]
        if background:
            tiers.append(Tier('Фон', background))

        return cls(tiers, duration)

    @classmethod
    def from_antx(cls, contents):
        """Creates Annotation instance from .antx file contents."""

        ann = contents.getroot()
        namespace = {'ns': 'http://tempuri.org/AnnotationSystemDataSet.xsd'}
        samplerate = cls._get_samplerate(ann, namespace)
        layers, max_end = cls._get_layers(ann, namespace, samplerate)
        duration = cls._get_duration_antx(max_end)

        return cls(layers, duration)

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
        except Exception:
            try:
                last_time = int(root.find('TIME_ORDER')[-1].get('TIME_VALUE')) / 1000
                duration = last_time if last_time > 300.0 else 300.0
            except IndexError:
                duration = 300.0

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

        # if the initial interval text was empty and a formatted string
        # with leading space was appended
        for interval in transcription:
            interval.text = interval.text.strip()

        return transcription, background

    @staticmethod
    def _set_ends(intervals, duration) -> None:
        """Sets ends for intervals."""

        for i in range(len(intervals)-1):
            intervals[i].end = intervals[i+1].start

        if intervals: intervals[i].end = duration

    @staticmethod
    def _get_samplerate(root, namespace: dict) -> int:
        """Extracts sample rate from .antx file root"""

        value = root.find(".//*[ns:Key='Samplerate']/ns:Value", namespace)
        samplerate = int(value.text)

        return samplerate

    @staticmethod
    def _get_layers(root, namespace: dict, samplerate: int) -> list:
        """Return list of Tier objects and their Intervals from .antx root."""

        max_end = 0  # used to determine duration of annotation

        layers = []
        for layer in root.findall('ns:Layer', namespace):
            layer_id = layer.find('ns:Id', namespace).text
            name = layer.find('ns:Name', namespace).text
            segments = []

            for seg in root.findall(f".//*[ns:IdLayer='{layer_id}']", namespace):
                samp_start = float(seg.find('ns:Start', namespace).text)
                samp_duration = float(seg.find('ns:Duration', namespace).text)
                text = seg.find('ns:Label', namespace).text

                samp_end = samp_start + samp_duration
                start = samp_start / samplerate
                end = samp_end / samplerate

                if not max_end or end > max_end:
                    max_end = end

                segments.append(Interval(start, end, text))

            layers.append(Tier(name, segments))

        return layers, max_end

    @staticmethod
    def _get_duration_antx(max_end: float) -> float:
        """Gets annotation duration from .antx file root."""

        if max_end > 15.0:
            return max_end
        else:
            return 15.0

    def to_tg(self) -> str:
        "Returns a string representing Annotation to be written into .TextGrid"

        for tier in self:
            if not tier.is_point:
                tier.fill_gaps(self.duration)

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
            tg_ann += tier.to_tg(t, self.duration)

        return tg_ann

    def to_eaf(self) -> ET.ElementTree:
        "Returns an Element Tree representing Annotation to be written into .eaf"

        # if include point intervals checkbutton is toggled
        if interface.body.output_frame.incl_point_var.get():
            for tier in self:
                if tier.is_point:
                    tier.extend_points(self.duration)
        else:
            self.tiers = [tier for tier in self if not tier.is_point]
 
        ann_doc = self._eaf_root()
        ann_tree = ET.ElementTree(ann_doc)

        self._eaf_header(ann_doc)
        self._time_slots(ann_doc, self._time_values())
        for tier in self: tier.to_eaf(ann_doc)
        self._time_slot_refs(ann_doc)

        self._default_lt(ann_doc)
        self._time_sub(ann_doc)
        self._symb_sub(ann_doc)
        self._symb_assoc(ann_doc)
        self._incl_in(ann_doc)

        return ann_tree

    def to_antx(self) -> ET.ElementTree:
        """Returns an Element Tree representing Annotation to be written into .antx"""

        if interface.body.output_frame.incl_point_var.get():
            for tier in self:
                if tier.is_point:
                    tier.extend_points(self.duration)
        else:
            self.tiers = [tier for tier in self if not tier.is_point]

        namespace = {'ns': 'http://tempuri.org/AnnotationSystemDataSet.xsd'}

        ann = self._antx_root()
        ann_tree = ET.ElementTree(ann)

        for tier in self:
            tier.to_antx(ann, self._generate_id())

        if interface.body.output_frame.incl_empty_var.get():
            for tier in self:
                for interval in tier:
                    interval.to_antx(ann, self._generate_id(), tier._antx_id)
        else:
            for tier in self:
                for interval in tier:
                    if not interval.text:
                        continue
                    interval.to_antx(ann, self._generate_id(), tier._antx_id)

        self._configs(ann)

        return ann_tree

    @staticmethod
    def _eaf_root() -> ET.Element:
        "Returns root ANNOTATION_DOCUMENT element for .eaf tree"

        root = ET.Element(
            'ANNOTATION_DOCUMENT',
            {'AUTHOR': '', 'FORMAT': '3.0', 'VERSION': '3.0',
            'xmlns:xsi': 'http://www.w3.org/2001/XMLSchema-instance',
            'xsi:noNamespaceSchemaLocation': 'http://www.mpi.nl/tools/elan/EAFv3.0.xsd'}
        )

        return root

    @staticmethod
    def _eaf_header(root) -> None:
        "Creates HEADER element in .eaf tree"

        header = ET.SubElement(root, 'HEADER', {'MEDIA_FILE': '',
                                                'TIME_UNITS': 'milliseconds'}
        )

        urn = ET.SubElement(header, 'PROPERTY', {'NAME': 'URN'})
        last_ann = ET.SubElement(header, 'PROPERTY', {'NAME': 'lastUsedAnnotationId'})

        urn.text = 'urn:nl-mpi-tools-elan-eaf:187f732a-340c-4c9e-a8c3-307ba38799fb'
        last_ann.text = '0'

    def _time_values(self) -> list:
        "Returns a list of time values of all annotation intervals."

        time_values = []

        # if include empty intervals checkbutton is toggled
        if interface.body.output_frame.incl_empty_var.get():
            for tier in self:
                for interval in tier:
                    time_values.append(interval.eaf_start)
                    time_values.append(interval.eaf_end)
        else:
            for tier in self:
                for interval in tier:
                    if not interval.text:
                        continue
                    time_values.append(interval.eaf_start)
                    time_values.append(interval.eaf_end)
        
        time_values.sort()

        return time_values

    @staticmethod
    def _time_slots(root, time_values) -> None:
        "Creates TIME_ORDER and TIME_SLOT elements in .eaf tree from time values"

        time_order = ET.SubElement(root, 'TIME_ORDER')

        for i, tv in enumerate(time_values, start=1):
            ET.SubElement(time_order, 'TIME_SLOT', {'TIME_SLOT_ID' : 'ts' + str(i),
                                                    'TIME_VALUE': str(tv)}
            )

    @staticmethod
    def _time_slot_refs(root) -> None:
        "Replaces TIME_SLOT_REF1/2 attribute values with references"

        for ann in root.iter('ALIGNABLE_ANNOTATION'):
            for ts in root.find('TIME_ORDER'):

                if ann.get('TIME_SLOT_REF1') == ts.get('TIME_VALUE'):
                    ann.set('TIME_SLOT_REF1', ts.get('TIME_SLOT_ID'))

                if ann.get('TIME_SLOT_REF2') == ts.get('TIME_VALUE'):
                    ann.set('TIME_SLOT_REF2', ts.get('TIME_SLOT_ID'))

    @staticmethod
    def _default_lt(root) -> None:
        "Creates LINGUISTIC_TYPE element for default-lt type in .eaf tree"

        ET.SubElement(root, 'LINGUISTIC_TYPE', {'GRAPHIC_REFERENCES': 'false',
                                                'LINGUISTIC_TYPE_ID': 'default-lt',
                                                'TIME_ALIGNABLE': 'true'}
        )

    @staticmethod
    def _time_sub(root) -> None:
        "Creates CONSTRAINT element for Time_Subdivision in .eaf tree"

        DESC = (
            "Time subdivision of parent annotation's time interval, no time "
            "gaps allowed within this interval"
        )

        ET.SubElement(root, 'CONSTRAINT', {'DESCRIPTION': DESC,
                                           'STEREOTYPE': 'Time_Subdivision'}
        )

    @staticmethod
    def _symb_sub(root) -> None:
        "Creates CONSTRAINT element for Symbolic_Subdivision in .eaf tree"

        DESC = (
            "Symbolic subdivision of a parent annotation. "
            "Annotations refering to the same parent are ordered"
        )

        ET.SubElement(root, 'CONSTRAINT', {'DESCRIPTION': DESC,
                                           'STEREOTYPE': 'Symbolic_Subdivision'}
        )

    @staticmethod
    def _symb_assoc(root) -> None:
        "Creates CONSTRAINT element for Symbolic_Association in .eaf tree"

        DESC = "1-1 association with a parent annotation"

        ET.SubElement(root, 'CONSTRAINT', {'DESCRIPTION': DESC,
                                           'STEREOTYPE': 'Symbolic_Association'}
        )

    @staticmethod
    def _incl_in(root) -> None:
        "Creates CONSTRAINT element for Included_In in .eaf tree"

        DESC = (
            "Time alignable annotations within the parent annotation's "
            "time interval, gaps are allowed"
        )

        ET.SubElement(root, 'CONSTRAINT', {'DESCRIPTION': DESC,
                                           'STEREOTYPE': 'Included_In'}
        )

    @staticmethod
    def _antx_root() -> ET.Element:
        """Returns root AnnotationSystemDataSet element for .antx tree"""

        root = ET.Element(
            'AnnotationSystemDataSet',
            {'xmlns': 'http://tempuri.org/AnnotationSystemDataSet.xsd'}
        )

        return root

    @staticmethod
    def _generate_id() -> str:
        """Generates a unique id for .antx layer/segment."""

        CHARS = '0123456789abcdef'

        generated = (
            ''.join(random.choices(CHARS, k=8)) + '-'
            + ''.join(random.choices(CHARS, k=4)) + '-'
            + ''.join(random.choices(CHARS, k=4)) + '-'
            + ''.join(random.choices(CHARS, k=4)) + '-'
            + ''.join(random.choices(CHARS, k=12))
        )

        return generated

    @staticmethod
    def _configs(root) -> None:
        """Creates Configuration elements in .antx file tree."""

        version = ET.SubElement(root, 'Configuration')
        key = ET.SubElement(version, 'Key'); key.text = 'Version'
        value = ET.SubElement(version, 'Value'); value.text = '5'

        created = ET.SubElement(root, 'Configuration')
        key = ET.SubElement(created, 'Key'); key.text = 'Created'
        value = ET.SubElement(created, 'Value')

        modified = ET.SubElement(root, 'Configuration')
        key = ET.SubElement(modified, 'Key'); key.text = 'Modified'
        value = ET.SubElement(modified, 'Value')

        samplerate = ET.SubElement(root, 'Configuration')
        key = ET.SubElement(samplerate, 'Key'); key.text = 'Samplerate'
        value = ET.SubElement(samplerate, 'Value'); value.text = '44100'

        file_vers = ET.SubElement(root, 'Configuration')
        key = ET.SubElement(file_vers, 'Key'); key.text = 'FileVersion'
        value = ET.SubElement(file_vers, 'Value'); value.text = '5'

        author = ET.SubElement(root, 'Configuration')
        key = ET.SubElement(author, 'Key'); key.text = 'Author'
        value = ET.SubElement(author, 'Value')

        title = ET.SubElement(root, 'Configuration')
        key = ET.SubElement(title, 'Key'); key.text = 'ProjectTitle'
        value = ET.SubElement(title, 'Value')

        environ = ET.SubElement(root, 'Configuration')
        key = ET.SubElement(environ, 'Key'); key.text = 'ProjectEnvironment'
        value = ET.SubElement(environ, 'Value')

        noises = ET.SubElement(root, 'Configuration')
        key = ET.SubElement(noises, 'Key'); key.text = 'ProjectNoises'
        value = ET.SubElement(noises, 'Value')

        collect = ET.SubElement(root, 'Configuration')
        key = ET.SubElement(collect, 'Key'); key.text = 'ProjectCollection'
        value = ET.SubElement(collect, 'Value')

        corpus_type = ET.SubElement(root, 'Configuration')
        key = ET.SubElement(corpus_type, 'Key'); key.text = 'ProjectCorpusType'
        value = ET.SubElement(corpus_type, 'Value')

        corpus_own = ET.SubElement(root, 'Configuration')
        key = ET.SubElement(corpus_own, 'Key'); key.text = 'ProjectCorpusOwner'
        value = ET.SubElement(corpus_own, 'Value')

        lic = ET.SubElement(root, 'Configuration')
        key = ET.SubElement(lic, 'Key'); key.text = 'ProjectLicense'
        value = ET.SubElement(lic, 'Value')

        desc = ET.SubElement(root, 'Configuration')
        key = ET.SubElement(desc, 'Key'); key.text = 'ProjectDescription'
        value = ET.SubElement(desc, 'Value')


class InputFrame(ttk.Labelframe):
    "Labelframe containing"

    RE_NAME = re.compile(r'[^/]+$')
    RE_TXT = re.compile(r'\.(textgrid|txt)$', re.I)
    RE_XML = re.compile(r'\.(eaf|trs|antx)$', re.I)
    ENCOD_MSG = ("Кодування файлу(ів) {} не підтримується. Будь ласка, "
                 "збережіть файл(и) у кодуванні UTF-8 та спробуйте ще раз.")

    def __init__(self, master, *args, **kwargs):
        super().__init__(master, *args, **kwargs)

        self.names, self.contents = [], []
        self.names_var = tk.StringVar(self, value=self.names)
        self.lb_files = tk.Listbox(self, height=10, width=45, activestyle='none',
                                   listvariable=self.names_var)
        self.lb_files.bind('<<ListboxSelect>>', self.btn_remove_state)

        self.sb_files = ttk.Scrollbar(self, orient=tk.VERTICAL,
                                      command=self.lb_files.yview)
        self.lb_files['yscrollcommand'] = self.sb_files.set

        self.btn_select = ttk.Button(self, text="Обрати файли", width=14,
                                     command=self.select_files)

        self.btn_clear = ttk.Button(self, text="Очистити все", width=14,
                                    command=self.clear_files)

        self.btn_remove = ttk.Button(self, text="Видалити",
                                     command=self.remove_files, width=14,
                                     state='disabled')

        self._layout()

    def _layout(self) -> None:
        "Lays widgets out"

        self.lb_files.pack(side=tk.LEFT)
        self.sb_files.pack(side=tk.LEFT, fill=tk.Y)
        self.btn_select.pack(padx=10)
        self.btn_clear.pack(padx=10, pady=3)
        self.btn_remove.pack(padx=10)
    
    def select_files(self) -> None:
        "Extracts contents from user selected files and updates file names"

        paths = self._get_paths()
        names = self._get_names(paths)
        formats = self._get_formats(names)
        names, contents = self._read_files(paths, names, formats)
        self.names.extend(names)
        self.contents.extend(contents)
        self.names_var.set(self.names)

    def clear_files(self) -> None:
        "Clear all contents and names of all files"

        self.contents.clear()
        self.names.clear()
        self.names_var.set(self.names)
        self.btn_remove.state(['disabled'])

    def remove_files(self) -> None:
        "Remove contents and names of the selected file in lb_files"

        i = self.lb_files.curselection()[0]
        if i == len(self.names) - 1:
            self.btn_remove.state(['disabled'])
        del self.contents[i], self.names[i]
        self.names_var.set(self.names)

    def btn_remove_state(self, *args) -> None:
        "Changes state of file remove button"

        if self.names: self.btn_remove.state(['!disabled'])

    @staticmethod
    def _get_paths() -> list:
        "Returns a list of strings representing paths for user-selected files"

        paths = askopenfilenames(
            title='Оберіть вхідний(і) файл(и) у кодуванні UTF-8',
            filetypes=[('Файли Praat', '*.TextGrid'),
                       ('Файли Elan', '*.eaf'),
                       ('Файли Transcriber', '*.trs'),
                       ('Файли Annotation Pro', '*.antx')]
        )

        return paths

    @staticmethod
    def _get_names(paths) -> list:
        "Returns a list of file names extracted from paths"

        return [InputFrame.RE_NAME.search(path).group() for path in paths]

    @staticmethod
    def _get_formats(names) -> list:
        "Returns a list of file formats extracted from names"

        formats = []

        for name in names:
            if InputFrame.RE_TXT.search(name):
                formats.append('txt')
            elif InputFrame.RE_XML.search(name):
                formats.append('xml')

        return formats

    @staticmethod
    def _read_files(paths, names, formats):
        contents = []
        unsupported = []

        for p, n, f in zip(paths, names, formats):
            with open(p, encoding='UTF-8') as inp_file:
                try:
                    if f == 'txt':
                        contents.append(inp_file.read())
                    elif f == 'xml':
                        contents.append(ET.parse(inp_file))
                except UnicodeDecodeError:
                    unsupported.append(n)
        
        for u in unsupported: names.remove(u)

        if unsupported:
            messagebox.showerror(
                title="Кодування не підтримується",
                message=InputFrame.ENCOD_MSG.format(', '.join(unsupported))
            )

        return names, contents


class OutputFrame(ttk.Labelframe):

    def __init__(self, master, *args, **kwargs):
        super().__init__(master, *args, **kwargs)

        self.format_var = tk.IntVar(self)
        self.rb_tg = ttk.Radiobutton(self, text=".TextGrid (Praat)", value=1,
                                     variable=self.format_var,
                                     command=self.cb_state)
        self.rb_eaf = ttk.Radiobutton(self, text=".eaf (Elan)", value=2,
                                      variable=self.format_var,
                                      command=self.cb_state)
        self.rb_antx = ttk.Radiobutton(self, text=".antx (Annotation Pro)", value=3,
                                      variable=self.format_var,
                                      command=self.cb_state)

        self.incl_empty_var = tk.BooleanVar(self)
        self.cb_incl_empty = ttk.Checkbutton(self, variable=self.incl_empty_var,
                                             offvalue=0, onvalue=1, state='disabled',
                                             text="Порожні інтервали")

        self.incl_point_var = tk.BooleanVar(self)
        self.cb_incl_point = ttk.Checkbutton(self, variable=self.incl_point_var,
                                             offvalue=0, onvalue=1, state='disabled',
                                             text="Точкові рівні") 

        self._layout()

    def cb_state(self) -> None:
        "Changes state of checkbuttons for .eaf output format"
        
        if self.format_var.get() == 1:
            self.cb_incl_empty.config(state='disabled')
            self.cb_incl_point.config(state='disabled')
        else:
            self.cb_incl_empty.config(state='active')
            self.cb_incl_point.config(state='active')

    def _layout(self) -> None:

        self.rb_tg.grid(row=0, column=0, sticky='w', padx=5, pady=2)
        self.rb_eaf.grid(row=1, column=0, sticky='w', padx=5, pady=2)
        self.rb_antx.grid(row=2, column=0, sticky='w', padx=5, pady=2)
        self.cb_incl_empty.grid(row=1, column=1, sticky='w', padx=5)
        self.cb_incl_point.grid(row=1, column=2, sticky='w', padx=5)

    
class ConvertFrame(tk.Frame):

    def __init__(self, master, *args, **kwargs):
        super().__init__(master, *args, **kwargs)

        self.btn_convert = ttk.Button(self, default='active', width=18,
                                      command=self.convert,
                                      text="Конвертувати все")

        self._layout()

    def convert(self) -> None:
        "Converts, duh"

        names = self.master.input_frame.names
        contents = self.master.input_frame.contents
        sel_fmt = self.master.output_frame.format_var.get()

        if names and sel_fmt:
            for name, contents in zip(names, contents):

                if name.lower().endswith('.textgrid'):
                    ann = Annotation.from_tg(contents)
                elif name.lower().endswith('.eaf'):
                    ann = Annotation.from_eaf(contents)
                elif name.lower().endswith('.trs'):
                    ann = Annotation.from_trs(contents)
                elif name.lower().endswith('.antx'):
                    ann = Annotation.from_antx(contents)

                if sel_fmt == 1:
                    save_path = asksaveasfilename(
                        title="Збережіть результат конвертації " + name,
                        defaultextension="TextGrid",
                        filetypes=[("Файли Praat", "*.TextGrid")],
                    )
                elif sel_fmt == 2:
                    save_path = asksaveasfilename(
                        title="Збережіть результат конвертації " + name,
                        defaultextension="eaf",
                        filetypes=[("Файли Elan", "*.eaf")]
                    )
                elif sel_fmt == 3:
                    save_path = asksaveasfilename(
                        title="Збережіть результат конвертації " + name,
                        defaultextension="antx",
                        filetypes=[("Файли Annotation Pro", "*.antx")]
                    )

                if save_path.endswith(".TextGrid"):
                    with open(save_path, 'w', encoding='UTF-8') as sf:
                        sf.write(ann.to_tg())
                elif save_path.endswith(".eaf"):
                    ann.to_eaf().write(save_path, 'UTF-8', xml_declaration=True)
                elif save_path.endswith(".antx"):
                    ann.to_antx().write(save_path, 'UTF-8', xml_declaration=True)

            messagebox.showinfo(title="Готово!", message="Готово!")

        elif not names and sel_fmt:
            messagebox.showerror(
                title="Чогось не вистачає...",
                message="Оберіть вхідний(і) файл(и)."
            )

        elif names and not sel_fmt:
            messagebox.showerror(
                title="Чогось не вистачає...",
                message="Оберіть кінцевий формат."
            )

        else:
            messagebox.showerror(
                title="Чогось не вистачає...",
                message="Оберіть вхідний(і) файл(и) та кінцевий формат."
            )

    def _layout(self):
        "Lays out"

        self.btn_convert.grid(row=0, column=1, sticky='e')
        self.btn_convert.pack()


class Body(tk.Frame):

    def __init__(self, master, *args, **kwargs):
        super().__init__(master, *args, **kwargs)
        self.input_frame = InputFrame(self, text="Вхідні файли",
                                      padding=7)
        self.output_frame = OutputFrame(self, text="Кінцевий формат",
                                        padding=12)
        self.convert_frame = ConvertFrame(self)

        self.input_frame.pack()
        self.output_frame.pack(pady=15)
        self.convert_frame.pack(side=tk.RIGHT)


class Interface(tk.Tk):
    """Represents AnnCo interface."""

    def __init__(self, *args, **kwargs) -> None:
        super().__init__(*args, **kwargs)
        self.title("AnnCo v2.0")
        self.body = Body(self, padx=15, pady=15)

        self.body.pack()


if __name__ == '__main__':
    interface = Interface()
    interface.mainloop()