import shutil


def sc_zip(path):
    shutil.make_archive(path, 'zip', path)