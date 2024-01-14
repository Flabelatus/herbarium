import time
import hashlib
import sqlite3
from datetime import datetime
import pandas as pd
import docx2txt
import platform
import flet
import os, shutil
from flet import Page, Column, Container, Card, Divider, Text

EXCEL_FIELDS = ["No", "Fam", "Nam", "Loc", "Col", "Det", "Col_Dat", "Det_Dat"]
DB = "hashes.db"


def get_path():
    path = ""
    os_name = platform.system()

    if os_name == "Darwin":
        path = os.path.expanduser("~/Desktop")
    elif os_name == "Windows":
        user_profile = os.environ["USERPROFILE"]
        path = os.path.join(user_profile, "Desktop")

    if "Herbarium" not in os.listdir(path):
        os.mkdir(os.path.join(path, "Herbarium"))

    root = os.path.join(path, "Herbarium")
    return root


def create_table():
    path = get_path()

    connection = sqlite3.connect(os.path.join(path, DB))
    cursor = connection.cursor()

    # Create the table if it doesn't exist
    cursor.execute(
        """
        CREATE TABLE IF NOT EXISTS hashes (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            hash TEXT,
            created_at TIMESTAMP
        )
    """
    )

    connection.commit()
    connection.close()


def get_file_type(file):
    extension = file.split(".")
    if len(extension) > 1 and extension[0] != "":
        return extension[1]


def parse_record(record):
    lines = record.split("\n")
    record_dict = {"No": "", "Fam": "", "Nam": "", "Loc": ""}
    for line in lines:
        if ": " in line:
            key, value = line.split(": ", 1)
            record_dict[key.strip()] = value.strip()

    record_dict["Col_Dat"] = ""
    record_dict["Det_Dat"] = ""

    for k, v in record_dict.items():
        c = None

        if "Dat" in v:
            c = v.split("Dat:", -1)
        if k == "Col":
            record_dict[k] = c[0].rstrip()
            record_dict["Col_Dat"] = c[1].rstrip()
        elif k == "Det":
            record_dict[k] = c[0].rstrip()
            record_dict["Det_Dat"] = c[1].rstrip()

    return record_dict


def main(page: Page):
    create_table()
    page.theme_mode = "light"
    page.title = "Herbarium GUI"
    page.description = "A simple tool to convert the data from docx to csv"
    page.window_width = 750
    page.window_height = 500
    page.theme = flet.Theme(
        visual_density=flet.ThemeVisualDensity.COMFORTABLE,
        use_material3=True,
        color_scheme_seed="#4336f5",
    )

    info_text = Text(color=flet.colors.BROWN, size=15, weight=flet.FontWeight.W_600)

    def convert(file):
        directory_path = get_path()
        default_path = os.path.join(directory_path, file)

        my_text = docx2txt.process(default_path)
        records = my_text.strip().split("\n\n\n")

        parsed_records = [parse_record(record) for record in records]
        df = pd.DataFrame(parsed_records)

        path_to_save = os.path.join(directory_path, "excel files", "excel-data.csv")

        if not os.path.isfile(path_to_save):
            initial_excel_df = pd.DataFrame(columns=EXCEL_FIELDS)
            initial_excel_df.to_csv(path_to_save, index=False)

        existing_df = pd.read_csv(path_to_save)
        df_without_header = df.iloc[1:]

        updated_df = pd.concat([existing_df, df_without_header], ignore_index=True)
        updated_df.to_csv(path_to_save, index=False)

    def message(msg, timeout):
        time.sleep(0.2)

        info_text.value = msg
        page.update()

        time.sleep(timeout)

        info_text.value = ""
        page.update()

    def execute(e):
        connection = sqlite3.connect(os.path.join(get_path(), DB))
        cursor = connection.cursor()

        has_docx = False
        directory_path = get_path()
        if "saved" not in os.listdir(directory_path):
            os.mkdir(os.path.join(directory_path, "saved"))
        if "excel files" not in os.listdir(directory_path):
            os.mkdir(os.path.join(directory_path, "excel files"))

        for file in os.listdir(directory_path):
            extensions = get_file_type(file)
            if extensions and extensions == "docx":
                has_docx = True

        if has_docx:
            for index, file in enumerate(os.listdir(directory_path)):
                if get_file_type(file) == "docx":
                    tail = file + str(
                        os.path.getsize(os.path.join(directory_path, file))
                    )

                    hashed_values = hashlib.md5(tail.encode()).hexdigest()
                    created_at = datetime.now()

                    hash_from_db = cursor.execute(
                        "SELECT * FROM hashes WHERE hash = ?", (hashed_values,)
                    ).fetchall()

                    if len(hash_from_db) > 0:
                        message(f"{file} is already saved in the Excel!", 1)
                        try:
                            shutil.move(
                                os.path.join(directory_path, file),
                                os.path.join(directory_path, "saved"),
                            )
                        except shutil.Error as err:
                            print(err)
                        continue

                    cursor.execute(
                        "INSERT INTO hashes (hash, created_at) VALUES (?, ?)",
                        (hashed_values, created_at),
                    )
                    connection.commit()

                    message("Processing File...", 1)
                    convert(file)
                    message("File Saved!", 0.6)
                    message(f"{file} now moved to the directory called 'saved'", 2.5)

                    try:
                        shutil.move(
                            os.path.join(directory_path, file),
                            os.path.join(directory_path, "saved"),
                        )
                    except shutil.Error as err:
                        print(err)
        else:
            message("No files with .docx format found", 1)

        connection.close()

    gui = Container(
        bgcolor="#CCFFCC",
        content=Column(
            horizontal_alignment=flet.CrossAxisAlignment.CENTER,
            controls=[
                Divider(height=25, color="#CCFFCC"),
                Card(
                    width=580,
                    content=Container(
                        padding=40,
                        content=Column(
                            horizontal_alignment=flet.CrossAxisAlignment.CENTER,
                            controls=[
                                Text(
                                    "Make sure to place all your word (.docx) files inside a folder called"
                                    " 'Herbarium' in the Desktop",
                                    size=16,
                                    text_align=flet.TextAlign.CENTER,
                                    color=flet.colors.GREEN,
                                )
                            ],
                        ),
                    ),
                ),
                Card(
                    width=580,
                    color="#FFFFFF",
                    content=Container(
                        padding=40,
                        content=Column(
                            horizontal_alignment=flet.CrossAxisAlignment.CENTER,
                            controls=[
                                Text(
                                    "File Converter Herbarium",
                                    size=30,
                                    weight=flet.FontWeight.W_600,
                                    color=flet.colors.GREEN,
                                ),
                                flet.ElevatedButton(
                                    "Convert",
                                    on_click=execute,
                                    bgcolor=flet.colors.GREEN,
                                    color=flet.colors.WHITE,
                                ),
                                info_text,
                            ],
                        ),
                    ),
                ),
                flet.Column(
                    alignment=flet.MainAxisAlignment.END,
                    controls=[
                        flet.IconButton(
                            icon=flet.icons.FEEDBACK,
                            icon_color=flet.colors.GREEN,
                        )
                    ],
                ),
            ],
        ),
    )

    page.views.append(gui)
    page.update()


flet.app(target=main)
