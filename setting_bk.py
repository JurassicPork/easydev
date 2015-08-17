import logging


NAME_EXTENSION = 'EasyDev'
ADDIN_ID = "org.universolibre.util.EasyDev"
ADDIN_VERSION = "0.1.0"
ADDIN_DESCRIPTION = 'Tools for developer in LibreOffice'
ADDIN_AUTHOR_WEB = "http://www.universolibre.org"
ADDIN_AUTHOR_NAME = "Universo Libre Mexico, A.C."
#~ ADDIN_DATA = ({
    #~ 'name': 'aletras',
    #~ 'description': 'Convierte una cantidad en letras',
    #~ 'parameters': {
        #~ 'cantidad':'La cantidad entera a convertir',},
    #~ },{
    #~ 'name': 'amoneda',
    #~ 'description': 'Convierte una cantidad en letras usando moneda',
    #~ 'parameters': {
        #~ 'cantidad':'La cantidad entera a convertir',
        #~ 'moneda':'La moneda usada, si se omite se asume pesos',
        #~ 'fraccion':'Fraccion de moneda, si se establece, se asume que también se convierten a letras',},
    #~ },
#~ )

FORMAT = '%(asctime)s - %(levelname)s - %(message)s'
DATE = '%d/%m/%Y %H:%M:%S'
logging.basicConfig(level=logging.DEBUG, format=FORMAT, datefmt=DATE)
LOG = logging.getLogger('MAKE_OXT')

PATH_IDLC = '/usr/lib/libreoffice/sdk/bin/idlc'
PATH_IDLC_INCLUDE = '/usr/share/idl/libreoffice'
PATH_REGMERGE = '/usr/lib/libreoffice/ure/bin/regmerge'

FILE_PY = '{}.py'.format(NAME_EXTENSION)
FILE_IDL = 'X{}.idl'.format(NAME_EXTENSION)
FILE_URD = 'X{}.urd'.format(NAME_EXTENSION)
FILE_RDB = 'X{}.rdb'.format(NAME_EXTENSION)
FILE_PNG = '{}.png'.format(NAME_EXTENSION.lower())
FILE_OXT = '{}_v{}.oxt'.format(NAME_EXTENSION, ADDIN_VERSION)
FILE_MANIFEST = 'manifest.xml'
FILE_DESCRIPTION = 'description.xml'
DIR_META = 'META-INF'
DIR_SOURCE = 'source'
DIR_BIN = 'bin'

XML_MANIFEST = """<?xml version="1.0" encoding="UTF-8"?>
<manifest:manifest>
    <manifest:file-entry manifest:full-path="{}" manifest:media-type="application/vnd.sun.star.uno-component;type=Python"/>
    <manifest:file-entry manifest:full-path="{}" manifest:media-type="application/vnd.sun.star.uno-typelibrary;type=RDB"/>
</manifest:manifest>
""".format(FILE_PY, FILE_RDB)

XML_DESCRIPTION = """<?xml version="1.0" encoding="UTF-8"?>
<description
    xmlns="http://openoffice.org/extensions/description/2006"
    xmlns:d="http://openoffice.org/extensions/description/2006"
    xmlns:xlink="http://www.w3.org/1999/xlink">

    <identifier value="{}" />
    <version value="{}" />
    <display-name><name lang="es">{}</name></display-name>
    <publisher><name xlink:href="{}" lang="es">{}</name></publisher>
    <icon><default xlink:href="icons/{}"/></icon>
    <extension-description>
        <src lang="es" xlink:href="descriptions/desc.es"/>
    </extension-description>
</description>
""".format(ADDIN_ID, ADDIN_VERSION, ADDIN_DESCRIPTION, ADDIN_AUTHOR_WEB, ADDIN_AUTHOR_NAME, FILE_PNG)
