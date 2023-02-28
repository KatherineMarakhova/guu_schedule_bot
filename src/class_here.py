from openpyxl import *
from openpyxl.utils import range_boundaries
from pathlib import Path
import selenium_fcs as sf
import time


class direct:
    path = ''
    inst = ''
    napr = ''
    def __init__(self, path, inst):
        self.path = path
