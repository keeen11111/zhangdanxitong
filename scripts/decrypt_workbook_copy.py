from __future__ import annotations

import sys
import shutil
from pathlib import Path

import msoffcrypto


source = Path(sys.argv[1])
destination = Path(sys.argv[2])
password = sys.argv[3]
with source.open("rb") as encrypted:
    try:
        workbook = msoffcrypto.OfficeFile(encrypted)
        if not workbook.is_encrypted():
            raise ValueError("not encrypted")
        workbook.load_key(password=password)
        with destination.open("wb") as decrypted:
            workbook.decrypt(decrypted)
    except (ValueError, msoffcrypto.exceptions.FileFormatError):
        shutil.copyfile(source, destination)
