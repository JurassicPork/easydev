import argparse
import os
import sys
from subprocess import call
import zipfile
from conf import *


def exists(path):
    return os.path.exists(path)


def join(*paths):
    return os.path.join(*paths)


def validate():
    path = join(DIR_SOURCE, FILE_PY)
    if not exists(path):
        LOG.error('No se encontro el archivo PY')
        return False
    path = join(DIR_SOURCE, FILE_IDL)
    if not exists(path):
        LOG.error('No se encontro el archivo IDL')
        return False
    if not exists(PATH_IDLC):
        LOG.error('No se encontro el archivo IDLC necesario para compilar')
        return False
    if not exists(PATH_REGMERGE):
        LOG.error('No se encontro el archivo REGMERGE necesario para compilar')
        return False
    return True


def compilate_idl():
    LOG.info('Compilate IDL...')
    path_rdb = join(DIR_SOURCE, FILE_RDB)
    path_urd = join(DIR_SOURCE, FILE_URD)
    if exists(path_rdb):
        os.remove(path_rdb)
    if exists(path_urd):
        os.remove(path_urd)
    path = join(DIR_SOURCE, FILE_IDL)
    call([PATH_IDLC, '-I', PATH_IDLC_INCLUDE, path])
    #~ call([PATH_IDLC, path])
    call([PATH_REGMERGE, path_rdb, '/UCR',  path_urd])
    os.remove(path_urd)
    LOG.info('Compilate IDL sucesfully...')
    return True


def compress_zip():
    LOG.info('Compress zip extension...')
    if not exists(DIR_FILES):
        os.mkdir(DIR_FILES)
    path_oxt = join(DIR_FILES, FILE_OXT)
    z = zipfile.ZipFile(path_oxt, 'w', compression=zipfile.ZIP_DEFLATED)
    root_len = len(os.path.abspath(DIR_SOURCE))
    for root, dirs, files in os.walk(DIR_SOURCE):
        relative = os.path.abspath(root)[root_len:]
        for f in files:
            fullpath = join(root, f)
            file_name = join(relative, f)
            if file_name == FILE_IDL:
                continue
            z.write(fullpath, file_name, zipfile.ZIP_DEFLATED)
    z.close()
    LOG.info('Extension zip sucesfully...')
    return


def install():
    LOG.info('Install extension...')
    path_oxt = join(DIR_FILES, FILE_OXT)
    call(['unopkg', 'add', '-v', '-f', '-s', path_oxt])
    call(['soffice', '--calc'])
    LOG.info('Install extension sucesfully...')
    return


def make_xml():
    LOG.info('Update files XML...')
    path = join(DIR_SOURCE, DIR_META)
    if not exists(path):
        os.mkdir(path)

    path = join(path, FILE_MANIFEST)
    with open(path, 'w') as f:
        f.write(XML_MANIFEST)

    path = join(DIR_SOURCE, FILE_DESCRIPTION)
    with open(path, 'w') as f:
        f.write(XML_DESCRIPTION)

    path = join(DIR_FILES, FILE_UPDATE)
    with open(path, 'w') as f:
        f.write(XML_UPDATE)

    LOG.info('Files XML update sucesfully...')
    return


def process_command_line_arguments():
    parser = argparse.ArgumentParser(
        description='Make AddIn extension OXT')
    parser.add_argument('-i', '--install', dest='install', action='store_true',
        default=False, required=False)
    parser.add_argument('-o', '--only_compress', dest='only_compress',
        action='store_true', default=False, required=False)
    return parser.parse_args()


if __name__ == '__main__':
    args = process_command_line_arguments()
    if not validate():
        sys.exit(0)

    if args.only_compress:
        compress_zip()
        if args.install:
            install()
        sys.exit(0)

    if not compilate_idl():
        sys.exit(0)

    make_xml()
    compress_zip()
    if args.install:
        install()
    LOG.info('Extension make sucesfully...')

