import pandas as pd
import openpyxl as pyxl
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.styles import numbers
from openpyxl import load_workbook
import re


# 查找某列所在的索引
def find_column_index(worksheet, column_value):
    idx = 1
    for cell in next(worksheet.iter_rows(min_row=1, max_row=1, values_only=True)):
        if cell == column_value:
            return (idx)
        idx = idx + 1


# 查找某列所在的列名
def find_column_letter(worksheet, column_value):
    return pyxl.utils.get_column_letter(find_column_index(worksheet, column_value))


def generate_test_file(filepath):
    sheet_formula = pd.read_excel(filepath, sheet_name="Formula")
    sheet_value = pd.read_excel(filepath, sheet_name="Value")
    value_types = ["Text", "Integer", "Decimal", "Date", "Time"]
    types_to_format = {"Text": numbers.FORMAT_TEXT,
                       "Integer": numbers.FORMAT_NUMBER,
                       "Decimal": numbers.FORMAT_NUMBER_00,
                       "Date": numbers.FORMAT_DATE_DATETIME,
                       "Time": numbers.FORMAT_DATE_TIME6}
    init_values = {}
    row_counts = 0
    for value_type in value_types:
        init_values[value_type] = sheet_value[value_type]
        row_counts += len(sheet_value[value_type])
    writer = pd.ExcelWriter('./generate.xlsx', engine='xlsxwriter')

    for i in range(len(sheet_formula)):
        formula_type = sheet_formula.loc[i]["Type"]
        formula_max = sheet_formula.loc[i]["Max"]
        row_counts *= formula_max;
        row_datas = []
        if formula_max == 1:
            for value_type in value_types:
                for init_value in init_values[value_type]:
                    row_data = {"Formula": formula_type,
                                "ExpectedResult": 0,
                                "FormulaString": formula_type+"(["+value_type+"1])",
                                value_type+"1": init_value}
                    row_datas.append(row_data)
        df = pd.DataFrame(row_datas)
        df.to_excel(writer, sheet_name=formula_type, index=False)
    writer._save()

    workbook = load_workbook('./generate.xlsx')
    for i in range(len(sheet_formula)):
        formula_type = sheet_formula.loc[i]["Type"]
        formula_max = sheet_formula.loc[i]["Max"]
        worksheet = workbook[formula_type]
        if formula_max == 1:
            for value_type in value_types:
                column_letter = find_column_letter(worksheet, value_type+'1')
                for cell in worksheet[column_letter]:
                    cell.number_format = types_to_format[value_type]
            column_expected_result_letter = find_column_letter(worksheet, "ExpectedResult")
            column_formula_string_letter = find_column_letter(worksheet, "FormulaString")
            for col_idx in range(len(worksheet[column_letter])):
                if col_idx == 0:
                    continue
                col_idx_string = str(col_idx + 1)
                formula_string = worksheet[column_formula_string_letter + col_idx_string].value
                expected_result = "="
                results = re.findall(r"\[.*?\]", formula_string)
                for result in results:
                    formula_letter = find_column_letter(worksheet, result[1:-1])
                    expected_result = "=" + formula_string.replace(result, formula_letter + col_idx_string)
                    print(result[1:-1])
                    print(formula_letter)
                    print(expected_result)
                worksheet[column_expected_result_letter + col_idx_string] = expected_result

    # 保存Excel文件
    workbook.save('./TestCases.xlsx')


if __name__ == '__main__':
    generate_test_file('./init.xlsx')
