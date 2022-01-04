import os
import sys

import simplekml
from openpyxl import load_workbook


def _dms2dd(d: float, m: float, s: float) -> float:
    return d + m / 60 + s / 3600


def _parse_dms(cell) -> tuple:
    string = str(cell.value)
    return int(string[:2]), int(string[2:4]), int(string[4:6])


def xlsx2kml(xlsx_path: str, kml_path: str) -> None:
    wb = load_workbook(filename=xlsx_path)
    sheet = wb.active  # uses the active sheet

    kml = simplekml.Kml()

    coordinates = []

    for i, row in enumerate(sheet.rows):
        if row[0].value is None:
            break
        lat = _dms2dd(*_parse_dms(row[0]))
        lng = _dms2dd(*_parse_dms(row[1]))

        coordinates.append((lat, lng))

    kml.newpolygon(coords=coordinates)
    kml.save(kml_path)


if __name__ == '__main__':
    if not len(sys.argv) == 2:
        sys.exit("Invalid usage, please use: python xlsx2kml.py <input path>")
    argument_path = os.path.abspath(sys.argv[1])
    if os.path.isdir(argument_path):
        files_to_translate = [
            os.path.join(argument_path, filename)
            for filename in os.listdir(argument_path)]
        output_path = os.path.join(argument_path, os.path.basename(
            argument_path))
    else:
        files_to_translate = [argument_path]
        output_path, extension = os.path.splitext(argument_path)
    output_path += ".kml"
    for input_path in files_to_translate:
        output_path, extension = os.path.splitext(input_path)
        output_path += ".kml"
        xlsx2kml(input_path, output_path)