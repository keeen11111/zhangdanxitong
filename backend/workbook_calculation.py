"""Calculate personnel exports and retain calculated values in review copies."""
from __future__ import annotations

from pathlib import Path
from xml.etree import ElementTree as ET
from zipfile import ZipFile


def recalculate_personnel_workbook(master: Path, path: Path, insertions: list[dict]) -> dict:
    from core._personnel_formula_cache import refresh_personnel_formula_caches
    return refresh_personnel_formula_caches(master, path, insertions)


def preserve_review_formula_caches(calculated: Path, review: Path) -> None:
    """Review annotations do not change formulas; retain their calculated values."""
    ns = '{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'
    temporary = review.with_suffix('.cache.tmp')
    with ZipFile(calculated) as source, ZipFile(review) as annotations, ZipFile(temporary, 'w') as target:
        for entry in annotations.infolist():
            payload = annotations.read(entry.filename)
            if entry.filename.startswith('xl/worksheets/sheet') and entry.filename.endswith('.xml') and entry.filename in source.namelist():
                original = ET.fromstring(source.read(entry.filename))
                cached = {cell.attrib['r']: cell for cell in original.iter(ns + 'c') if cell.find(ns + 'f') is not None}
                root = ET.fromstring(payload)
                changed = False
                for cell in root.iter(ns + 'c'):
                    previous = cached.get(cell.attrib['r'])
                    formula = cell.find(ns + 'f')
                    if previous is None or formula is None or formula.text != previous.find(ns + 'f').text:
                        continue
                    value = previous.find(ns + 'v')
                    if value is None:
                        continue
                    existing = cell.find(ns + 'v')
                    if existing is not None:
                        cell.remove(existing)
                    cell.append(ET.fromstring(ET.tostring(value)))
                    if 't' in previous.attrib:
                        cell.attrib['t'] = previous.attrib['t']
                    else:
                        cell.attrib.pop('t', None)
                    changed = True
                if changed:
                    payload = ET.tostring(root, encoding='utf-8', xml_declaration=True)
            target.writestr(entry, payload)
    temporary.replace(review)
