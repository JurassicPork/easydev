# coding: utf-8
import logging
import sys


OS = sys.platform

FORMAT = '%(asctime)s - %(levelname)s - %(message)s'
DATE = '%d/%m/%Y %H:%M:%S'
logging.basicConfig(level=logging.DEBUG, format=FORMAT, datefmt=DATE)
LOG = logging.getLogger('EasyDev')

VERSION = '0.1.0'
ID_EXT = 'org.universolibre.util.EasyDev'
NAME_EXT = 'EasyDev'
LINUX = 'linux'
WIN = 'win32'
DESKTOP = 'com.sun.star.frame.Desktop'
TOOLKIT = 'com.sun.star.awt.Toolkit'
CALC = 'scalc'
WRITER = 'swriter'
EXT_PDF = 'pdf'
