# coding: utf-8

import logging
import sys
from com.sun.star.awt.MessageBoxButtons import BUTTONS_OK, BUTTONS_YES_NO
from com.sun.star.awt.MessageBoxResults import OK, YES, NO


OS = sys.platform
NAME_EXT = 'EasyDev'

FORMAT = '%(asctime)s - %(levelname)s - %(message)s'
DATE = '%d/%m/%Y %H:%M:%S'
logging.basicConfig(level=logging.DEBUG, format=FORMAT, datefmt=DATE)
LOG = logging.getLogger(NAME_EXT)

VERSION = '2.0.0'
ID_EXT = 'org.universolibre.EasyDev'
NODE = '/{}.Configuration/Settings'.format(ID_EXT)
NODE_CONFIG = 'Values'
LINUX = 'linux'
WIN = 'win32'
DESKTOP = 'com.sun.star.frame.Desktop'
TOOLKIT = 'com.sun.star.awt.Toolkit'
SRV_JOB = ('com.sun.star.task.Job',)
CALC = 'scalc'
WRITER = 'swriter'
EXT_PDF = 'pdf'
CLIPBOARD_FORMAT_TEXT = 'text/plain;charset=utf-16'
TIMEOUT = 10
