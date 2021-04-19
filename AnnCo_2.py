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