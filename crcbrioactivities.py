import openpyxl

#TODO: Add log system

FILENAME = "tests/data/brio.xlsx"
INITIAL_SHEET_NAME = "Sheet1"

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

    # TODO: get columns id from header
    columns_id = {}
    for row in sheet.iter_rows(min_row=1, max_row=1):
        columns_id = {cell.value: id for id, cell in enumerate(row)}

    print(columns_id)
