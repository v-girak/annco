# module used for storing test functions

import xml.etree.ElementTree as ET
import xml.dom.minidom as MD

from tkinter import ttk, messagebox
from tkinter.filedialog import askopenfilename, asksaveasfilename


def readable(tree: ET.ElementTree) -> None:
    "Create toprettyxml representation of .eaf tree"

    filepath = asksaveasfilename(
        defaultextension='eaf',
        filetypes=[('Файли Elan', '*.eaf')]
    )
    if not filepath:
        return

    tree.write(filepath, 'UTF-8')
    tree = MD.parse(filepath)
    tree_str = tree.toprettyxml()
    root = ET.fromstring(tree_str)
    tree = ET.ElementTree(root)
    tree.write(filepath, 'UTF-8', xml_declaration=True)
    messagebox.showinfo(title='Ура!', message='Файл збережено!')