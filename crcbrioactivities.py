import openpyxl
import logging

FILENAME = "tests/data/brio.xlsx"
INITIAL_SHEET_NAME = "Sheet1"

RECEIVED_100_PALETTES_RECEIVED_COLUMN_NAME = "Palettes sol 100*120 réellement reçues"
RECEIVED_80_PALETTES_RECEIVED_COLUMN_NAME = "Palettes sol 80*120 réellement reçues"
STATUS_COLUMN_NAME = "tStatut"

STATUS_OK = "livrée"
STATUS_NOT_OK = "annulée"

log = logging.getLogger(__name__)
formatter = logging.Formatter(
    "%(levelname)s %(asctime)s %(filename)s:%(lineno)d(%(funcName)s) %(message)s"
)
handler = logging.StreamHandler()
handler.setFormatter(formatter)
log.addHandler(handler)
log.setLevel(logging.DEBUG)


if __name__ == "__main__":
    file = openpyxl.load_workbook(FILENAME, data_only=True)
    print(file.sheetnames)
    sheet = file[INITIAL_SHEET_NAME]

    # Display sheet
    for row in sheet.iter_rows(min_row=1, max_row=1):
        for cell in row:
            print(cell.value, end=", ")
        print()
    for row in sheet.iter_rows(min_row=2, max_row=sheet.max_row):
        for cell in row:
            print(cell.value, end=", ")
        print()

    # Get columns id from header
    header_list = []
    columns_id = {}
    for row in sheet.iter_rows(min_row=1, max_row=1):
        header_list = list(row)
        columns_id = {cell.value: id for id, cell in enumerate(row)}

    print(columns_id)

    # Analyse 1 paletettes ans status
    try:
        received_100_palettes_column = columns_id[RECEIVED_100_PALETTES_RECEIVED_COLUMN_NAME]
        received_80_palettes_column = columns_id[RECEIVED_80_PALETTES_RECEIVED_COLUMN_NAME]
        status_column = columns_id[STATUS_COLUMN_NAME]
    except KeyError as error:
        log.error("The following column is nowhere to be found:", error)
        # TODO: stop function

    erroneous_rows = []
    for row in sheet.iter_rows(min_row=2, max_row=sheet.max_row):
        # TODO: test type of values
        total_received_palettes = (
            row[received_80_palettes_column].value + row[received_100_palettes_column].value
        )

        if total_received_palettes == 0 and row[status_column].value != STATUS_NOT_OK:
            erroneous_rows.append(row)
        elif total_received_palettes != 0 and row[status_column].value == STATUS_NOT_OK:
            erroneous_rows.append(row)

    print("erroneous_rows")
    for row in erroneous_rows:
        for cell in row:
            print(cell.value, end=", ")
        print()

    # Creation of a workbook.
    output_workbook = openpyxl.Workbook(write_only=True)

    # Add the worsheet for palettes analysis.
    palettes_worksheet = output_workbook.create_sheet("palettes")

    palettes_worksheet.append(["Analyse palettes"])
    palettes_worksheet.append([])
    palettes_worksheet.append(["Lignes avec incohérences :"])
    palettes_worksheet.append(header_list)
    for row in erroneous_rows:
        palettes_worksheet.append(row)

    # save workbook.
    output_workbook.save("analyse.xlsx")
