#!/usr/bin/python3.9

# Copyright (c) 2021, Valentin Svet
#
# Permission is hereby granted, free of charge, to any person obtaining a copy
# of this software and associated documentation files (the "Software"), to deal
# in the Software without restriction, including without limitation the rights
# to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
# copies of the Software, and to permit persons to whom the Software is
# furnished to do so, subject to the following conditions:
#
# The above copyright notice and this permission notice shall be included in all
# copies or substantial portions of the Software.
#
# THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
# IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
# FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
# AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
# LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
# OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
# SOFTWARE.


import os
import logging

import numpy as np
import pandas as pd

from zipfile import ZipFile
from warnings import simplefilter
from re import findall

FORMAT = "utf-8"

# autofill options for remarks part
REMARK_FILL_OPTIONS: dict = {
    # warning levels
    "1a": "Die Versetzung ist zurzeit gefährdet.\n",
    "1b": "Die Versetzung ist zurzeit stark gefährdet.\n",
    "1c": "Die Versetzung ist zurzeit ausgeschlossen.\n",

    # probationary period
    "2a": "<form_of_address> hat die Probezeit bestanden.\n",
    "2b": "Die allgemeine Schulpflicht ist erfüllt.\n",

    # repetition, resignation, skipping
    "3a": "<form_of_address> hat die Berufsbildungsreife erworben.\n",
    "3b": "Dieses Zeugnis ist der Berufsbildungsreife / der erweiterten Berufsbildungsreife gleichwertig.\n",

    # dyslexia
    "4a": "Aufgrund von festgestellten Lese- und Rechtschreibschwierigkeiten wurden "
          "die Lese- und Rechtschreibleistungen nicht in vollem Umfang bewertet.\n",

    # german
    "5a": "<form_of_address> hat an Fördermaßnahmen zur Verbesserung der deutschen Sprachkenntnisse teilgenommen.\n",

    # theology class
    "6a": "<form_of_address> hat am Religionsunterricht der Evangelischen Kirche teilgenommen.\n"
          "Der Träger kann eine eigene Teilnahmebescheinigung bzw. Beurteilung erteilen.\n"
    # Feel free to add your templates!
}

FILL_OPTIONS: dict = {
    # general
    "fill_subject": "……………………………………………*)",
    "first_semester": "Schulhalbjahr /S̶c̶h̶u̶l̶j̶a̶h̶r̶",
    "second_semester": "1̶.̶ ̶S̶c̶h̶u̶l̶h̶a̶l̶b̶j̶a̶h̶r̶ /Schuljahr",
    "male_pronounce": "S̶i̶e̶ /Er",
    "female_pronounce": "Sie /E̶r̶",
    "male_form_of_address": "D̶i̶e̶ ̶S̶c̶h̶ü̶l̶e̶r̶i̶n̶ /Der Schüler",
    "female_form_of_address": "Die Schülerin /D̶e̶r̶ ̶S̶c̶h̶ü̶l̶e̶r̶",
    "cross_out_not": "n̶i̶c̶h̶t̶",
    "not": "nicht",

    # boxes
    "picked_box": "☒",
    "unpicked_box": "☐"
}

# set template as working directory
os.chdir("./template")

# do not print FutureWarnings from numpy
simplefilter(action='ignore', category=FutureWarning)

# create and configure logger
logging.basicConfig(filename="../log_file.log",
                    format="[%(levelname)s]\t[%(asctime)s]\t%(message)s",
                    filemode="w")
logger = logging.getLogger()
logger.setLevel(logging.DEBUG)

# unzip template document
with ZipFile("../template.docx") as docx_template:
    docx_template.extractall("./")
    with open(file="./word/document.xml", mode="r", encoding=FORMAT) as document_template:
        # get document.xml from template
        TEMPLATE_CONTENT: str = document_template.read()

        # get attributes for certificate
        keys = findall(r"{+[\w]+}", TEMPLATE_CONTENT)
        KEY_LIST = [string[1:-1] for string in keys]
        KEY_LIST.append("geschlecht")
        KEY_LIST_LEN = len(KEY_LIST)


def get_all_file_paths() -> np.array:
    file_paths: np.array = np.array([])

    # iterate through all files through all directories
    for root, directories, files in os.walk("./"):
        for filename in files:
            # add filepath with root to array
            filepath = os.path.join(root, filename)
            file_paths = np.append(file_paths, filepath)

    return file_paths


def get_csv_table_data() -> np.array:
    student_data: np.array = np.array([])

    # iterate through all files
    for root, directories, files in os.walk("../tables"):
        for csv in files:
            # get path to file
            csv_path = os.path.join(root, csv)

            # load data to dictionary
            data: dict = {}
            try:
                # if the user instead decides to use a .csv
                data = pd.read_csv(csv_path, dtype=str, encoding=FORMAT).to_dict(orient="list")
            except ValueError:
                try:
                    # the application expects from user to use an .xls by default
                    data = pd.read_excel(csv_path, dtype=str, encoding=FORMAT).to_dict(orient="list")
                except ValueError:
                    # if the file should not be read
                    pass

            # add path to dataset for debug message
            data["path"] = csv_path

            # add dictionary to array
            student_data = np.append(student_data, data)

    return student_data


def generate_docx(docx_file_paths: np.array, student: dict) -> None:
    religion: str = student["religion"]
    if (religion != "nan") and (religion != ""):
        student["religion_label"] = "Religion"

    # modify values based on sex
    if student["geschlecht"] == "w":
        student["pronomen"] = FILL_OPTIONS["female_pronounce"]
        student["form_of_address"] = FILL_OPTIONS["female_form_of_address"]
    elif student["geschlecht"] == "m":
        student["pronomen"] = FILL_OPTIONS["male_pronounce"]
        student["form_of_address"] = FILL_OPTIONS["male_form_of_address"]
    else:
        logging.error(msg=f"[Schüler ID: {student['schueler_id']}]\t[Falsch Geschlechtsangabe: {student['geschlecht']}]")
        return None

    klasse: str = student["klasse"]
    semester: str = student["semester"]
    remarks: str = student["bemerkungen"]

    # modify values based on semester
    if semester == "1":
        student["semester"] = FILL_OPTIONS["first_semester"]
        student["kreuz1"] = FILL_OPTIONS["unpicked_box"]
        student["kreuz2"] = FILL_OPTIONS["picked_box"]
        student["kreuz3"] = FILL_OPTIONS["unpicked_box"]
        student["kreuz4"] = FILL_OPTIONS["picked_box"]

        student["jahr"] = f"{student['jahr'] }/{student['jahr'] + 1}"

        student["neue_jahrgangsstuffe"] = "/"
        student["bestanden"] = FILL_OPTIONS["not"]

    elif semester == "2":
        student["semester"] = FILL_OPTIONS["second_semester"]
        student["kreuz1"] = FILL_OPTIONS["picked_box"]
        student["kreuz2"] = FILL_OPTIONS["unpicked_box"]
        student["kreuz3"] = FILL_OPTIONS["picked_box"]
        student["kreuz4"] = FILL_OPTIONS["unpicked_box"]

        student["jahr"] = f"{student['jahr'] - 1}/{student['jahr']}"

        student["neue_jahrgangsstuffe"] = str(int(klasse[:-1]))

        # modify certificate for the case the student fails
        if "<1c>" in remarks:
            student["bestanden"] = FILL_OPTIONS["not"]
        else:
            student["bestanden"] = FILL_OPTIONS["cross_out_not"]
    else:
        logging.error(msg=f"[Schüler ID: {student['schueler_id']}]\t[Falsche Semesterangabe: {student['semester']}]")
        return None

    # insert text for keywords
    for option in REMARK_FILL_OPTIONS.keys():
        remarks = remarks.replace(f"<{option}>", REMARK_FILL_OPTIONS[option])

    # insert correct form of address
    remarks = remarks.replace("<form_of_address>", student["form_of_address"])
    student["bemerkungen"] = remarks

    # write document.xml with student data
    with open(file="./word/document.xml", mode="w", encoding=FORMAT) as document:
        document_content: str = TEMPLATE_CONTENT
        for key in KEY_LIST:
            document_content = document_content.replace("{" + key + "}", student[key])
        document.write(document_content)

    # build docx
    docx_path: str = f"../certificate/" \
                     f"[{student['schueler_id']}] " \
                     f"[{student['klasse']}] " \
                     f"{student['vorname']} " \
                     f"{student['familienname']}.docx"
    with ZipFile(docx_path, "w") as docx:
        for file in docx_file_paths:
            docx.write(file)
    logging.info(msg=f"Zeugnis fertig: {docx_path[15:-1]}")


def main() -> None:
    print(
        f"{'#' * 25} ZEUGNISS GENERAROR {'#' * 25} \n \n",
        f"{'#' * 25} von Valentin Svet  {'#' * 25} \n"
    )

    # create array with files of docx
    docx_file_paths: np.array = get_all_file_paths()

    # read tables
    logging.info(msg="Tabellen werden gelesen")
    student_data = get_csv_table_data()

    # create array with student IDs
    logging.info(msg="Schüler IDs werden zusammengetragen")
    student_id_list: np.array = np.array([])
    for data in student_data:
        try:
            for student_id in data["schueler_id"]:
                if student_id and (not np.isin(student_id, student_id_list)):
                    student_id_list = np.append(student_id_list, student_id)
        except KeyError:
            pass

    logging.info(msg="Datensätze werden gelesen und Zeugnisse geschrieben.")

    # get date for certificate
    date: str = str(input("Bitte geben Sie das Datum für die Zeugnisse an: "))
    year: int = int(date[-4:])

    # generate dictionary with all attributes from all files for all students
    for student_id in student_id_list:
        student: dict = {
            "schueler_id": str(student_id),
            "datum": date,
            "jahr": year,

            # default values
            "religion_label": FILL_OPTIONS["fill_subject"],
            "religion": "",
            "wpu1_name": FILL_OPTIONS["fill_subject"],
            "wpu1_note": "",
            "wpu2_name": FILL_OPTIONS["fill_subject"],
            "wpu2_note": "",
        }

        # iterate through all tables
        for data in student_data:
            try:
                # iterate through all records
                for i in range(len(data["schueler_id"])):
                    if str(data["schueler_id"][i]) == student_id:
                        # iterate through all attributes
                        for key in KEY_LIST:
                            if not (key in ["kreuz1", "kreuz2", "kreuz3", "kreuz4",
                                            "pronomen", "bestanden", "neue_jahrgangsstuffe"]):
                                try:
                                    value: str = str(data[key][i])
                                    if value != "nan":
                                        student[key] = value
                                        logging.info(msg=f"[Schüler ID: {student['schueler_id']}]\t"
                                                         f"[Attribut: {key}]\t"
                                                         f"[Wert: {value}]\t"
                                                         f"[Pfad: {data['path'][3:]}]"
                                                     )
                                except KeyError:
                                    pass
            except KeyError:
                pass

        # checks whether all attributes are available
        if KEY_LIST_LEN - len(student.keys()) == 6:
            generate_docx(docx_file_paths, student)
        else:
            for key in KEY_LIST:
                if not (key in student.keys()):
                    if not (key in ("kreuz1", "kreuz2", "kreuz3", "kreuz4",
                                    "pronomen", "bestanden", "neue_jahrgangsstuffe")):
                        logging.error(msg=f"[Schüler ID: {student['schueler_id']}]\t[Fehlender Attribut: {key}]")

    # remove student data from document.xml
    with open(file="./word/document.xml", mode="w") as document:
        document.write("")

    logging.info(msg="Zeugnisse sind fertig")


if "__main__" == __name__:
    main()
