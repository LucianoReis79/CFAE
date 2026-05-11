import os
import shutil
import tempfile
import zipfile

from pathlib import Path
from datetime import datetime

from openpyxl import load_workbook

from src.file_utils import clean_filename


def load_workbook_data(file_path):

    workbook = load_workbook(
        file_path,
        data_only=False,
        keep_vba=False
    )

    return workbook


def get_sheet_names(workbook):

    return workbook.sheetnames


def get_columns(
    workbook,
    sheet_name,
    return_unique=False,
    filter_column=None
):

    sheet = workbook[sheet_name]

    headers = [
        cell.value
        for cell in sheet[1]
    ]

    if not return_unique:
        return headers

    column_index = headers.index(filter_column)

    unique_values = set()

    for row in sheet.iter_rows(
        min_row=2,
        values_only=True
    ):

        value = row[column_index]

        if value is not None:

            unique_values.add(
                str(value)
            )

    return sorted(unique_values)


def remove_excel_tables(sheet):

    try:

        table_names = list(sheet.tables.keys())

        for table_name in table_names:

            del sheet.tables[table_name]

    except Exception:
        pass


def keep_only_selected_sheet(
    workbook,
    selected_sheet
):

    for sheet_name in workbook.sheetnames:

        if sheet_name != selected_sheet:

            workbook.remove(
                workbook[sheet_name]
            )


def remove_non_matching_rows(
    sheet,
    filter_column,
    filter_value
):

    headers = [
        cell.value
        for cell in sheet[1]
    ]

    col_index = headers.index(filter_column) + 1

    rows_to_delete = []

    for row in range(2, sheet.max_row + 1):

        value = sheet.cell(
            row=row,
            column=col_index
        ).value

        if str(value) != str(filter_value):

            rows_to_delete.append(row)

    for row in reversed(rows_to_delete):

        sheet.delete_rows(row)


def generate_filtered_files(
    source_file,
    sheet_name,
    filter_column,
    values,
    progress_bar,
    status_text
):

    output_dir = Path("output")

    output_dir.mkdir(exist_ok=True)

    total = len(values)

    base_temp = tempfile.NamedTemporaryFile(
        delete=False,
        suffix=".xlsx"
    )

    base_temp.close()

    shutil.copy(
        source_file,
        base_temp.name
    )

    generated_files = []

    for index, value in enumerate(values):

        status_text.text(
            f"Gerando arquivo: {value}"
        )

        temp_copy = tempfile.NamedTemporaryFile(
            delete=False,
            suffix=".xlsx"
        )

        temp_copy.close()

        shutil.copy(
            base_temp.name,
            temp_copy.name
        )

        workbook = load_workbook(
            temp_copy.name,
            data_only=True,
            keep_vba=False
        )

        keep_only_selected_sheet(
            workbook,
            sheet_name
        )

        sheet = workbook[sheet_name]

        remove_non_matching_rows(
            sheet,
            filter_column,
            value
        )

        remove_excel_tables(sheet)

        current_date = datetime.now()

        month = current_date.strftime("%m")

        year = current_date.strftime("%Y")

        safe_value = clean_filename(
            str(value)
        )

        filename = (
            f"{filter_column}_{safe_value}_{month}_{year}.xlsx"
        )

        final_path = output_dir / filename

        workbook.save(final_path)

        workbook.close()

        generated_files.append(final_path)

        os.remove(temp_copy.name)

        progress = int(
            ((index + 1) / total) * 100
        )

        progress_bar.progress(progress)

    os.remove(base_temp.name)

    zip_path = output_dir / "arquivos_filtrados.zip"

    with zipfile.ZipFile(
        zip_path,
        "w",
        zipfile.ZIP_DEFLATED
    ) as zipf:

        for file in generated_files:

            zipf.write(
                file,
                arcname=file.name
            )

    status_text.text(
        "Processo finalizado!"
    )

    return zip_path