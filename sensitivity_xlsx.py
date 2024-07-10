"""Change sensitivity label of xlsx files in folder."""
import os
import shutil
import tempfile
import zipfile

xml_content_confidential = """<?xml version="1.0"
"""

xml_content_internal = """<?xml version="1.0"
"""

xml_content_public = """<?xml version="1.0"
"""

xml_content_restricted = """<?xml version="1.0"
"""


def process_xlsx_files(
    folder_path: str,
    filenames_to_remove: list[str],
    xml_content: str,
        ) -> None:
    """
    Process XLSX files in the specified folder by replacing 'custom.xml'.

    Args:
        folder_path (str): Path to the folder containing XLSX files.
        filenames_to_remove (list[str]): List of filenames to remove from XLSX.
        xml_content (str): New sensitivity label to insert.
    """
    for filename in os.listdir(folder_path):
        if filename.endswith('.xlsx'):
            file_path = os.path.join(folder_path, filename)
            remove_from_zip(file_path, *filenames_to_remove)
            with zipfile.ZipFile(file_path, 'a') as zipwrite:
                zipwrite.writestr('docProps/custom.xml', xml_content)


def remove_from_zip(zipfname: str, *filenames: str) -> None:
    """
    Remove specified files from a ZIP archive.

    Args:
        zipfname (str): Path to the ZIP archive.
        filenames (str): Filenames to remove from the archive.
    """
    tempdir = tempfile.mkdtemp()
    tempname = os.path.join(tempdir, 'new.zip')
    copy_files_excluding(zipfname, tempname, filenames)
    shutil.move(tempname, zipfname)
    shutil.rmtree(tempdir)


def copy_files_excluding(
    zipfname: str,
    tempname: str,
    filenames: tuple[str, ...],
        ) -> None:
    """
    Copy files from one ZIP archive to another, excluding specified filenames.

    Args:
        zipfname (str): Path to the source ZIP archive.
        tempname (str): Path to the destination ZIP archive.
        filenames (tuple): Filenames to exclude from copying.
    """
    with zipfile.ZipFile(zipfname, 'r') as zipread:
        with zipfile.ZipFile(tempname, 'w') as zipwrite:
            for file_in_zip in zipread.infolist():
                if file_in_zip.filename not in filenames:
                    data_to_copy = zipread.read(file_in_zip.filename)
                    zipwrite.writestr(file_in_zip, data_to_copy)


# Example usage:
file_to_insert = xml_content_restricted
folder_path = os.path.abspath(os.getcwd())
filenames_to_remove = ['docProps/custom.xml']
process_xlsx_files(folder_path, filenames_to_remove, file_to_insert)
