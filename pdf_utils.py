from __future__ import unicode_literals, print_function

import os

from six.moves import tkinter
from six.moves.tkinter_filedialog import askdirectory
from comtypes import client
import PyPDF2
import progressbar


COMTYPES_PDF_FORMAT = 17


class PDFError(Exception):
    pass


def get_files_of_type(path, types):
    """
    Gets all files in `path` that have extension `type`.
    """
    if not isinstance(types, (list, tuple)):
        types = tuple(types)
    dirpath, dirnames, filenames = next(os.walk(path), (None, None, []))
    filenames = [i for i in filenames if any(i.lower().endswith(j.lower()) for j in types)]
    return filenames


def get_dir():
    window = tkinter.Tk()
    window.withdraw()
    target_dir = askdirectory()
    window.destroy()
    return os.path.normpath(target_dir)


def check_tmp_path(tmp_path):
    """
    Ensures that a path is an existing empty folder.
    """
    if not os.path.isdir(tmp_path):
        os.mkdir(tmp_path)
    elif not os.listdir(tmp_path):
        raise PDFError('Temporary directory must be empty.')


def word_to_pdf(path, tmp_path=None):
    """
    Converts all Word docs in `path` to PDF, collected in a temporary folder.

    If `tmp_path` is provided and it is an existing folder, it must be empty.
    Otherwise the folder is created.
    """
    if tmp_path is None:
        tmp_path = os.path.join(path, '.tmp')

    check_tmp_path(tmp_path)

    # Use Word itself to 'save as' each file as a PDF via the comtypes library
    word = client.CreateObject('Word.Application')
    word.Visible = False
    filenames = get_files_of_type(path, ('.doc', '.docx'))
    
    bar = progressbar.ProgressBar(max_value=len(filenames))
    for i, fn in enumerate(filenames):
        doc = word.Documents.Open(os.path.join(path, fn))
        out_fn = '{}.pdf'.format(fn.split('.')[0])
        doc.SaveAs(os.path.join(tmp_path, out_fn), FileFormat=COMTYPES_PDF_FORMAT)
        doc.close()
        bar.update(i + 1)
    word.Quit()


def merge_pdfs(path, out_path=None, use_outlines=False):
    """
    Combines all PDF's found at `path` into a single document.

    To set output filename and/or location use `out_path`
    """
    print('\nCombining pages')
    files = []
    if out_path is None:
        out_path = os.path.join(path, 'combined.pdf')
    filenames = get_files_of_type(path, '.pdf')
    merger = PyPDF2.PdfFileMerger()
    bar = progressbar.ProgressBar(max_value=len(filenames), redirect_stdout=True)
    for i, fn in enumerate(filenames):
        bk_txt = fn.split('.')[0]
        curr_path = os.path.join(path, fn)
        # Purposefully NOT using `with`, see http://stackoverflow.com/q/6773631
        f = open(curr_path, 'rb')
        # Can't close input files until output file is saved.
        # Instead, move to list to close later
        files.append(f)
        merger.append(f, bookmark=bk_txt, import_bookmarks=False)
        bar.update(i + 1)
    if use_outlines:
        merger.setPageMode('/UseOutlines')
    with open(out_path, 'wb') as f:
        merger.write(f)
    # Now we can close our files
    for f in files:
        f.close()
