import pandas as pd
import openpyxl as pyxl
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.styles import numbers
from openpyxl import load_workbook
import re
import numpy as np

global sheet_formula
global value_types
value_types = ["Text", "Integer", "Decimal", "Date", "Time"]
global testFilePath
testFilePath = "./TestCases.xlsx"


# 查找某列所在的索引
def find_column_index(worksheet, column_value):
    idx = 1
    for cell in next(worksheet.iter_rows(min_row=1, max_row=1, values_only=True)):
        if cell == column_value:
            return idx
        idx = idx + 1


# 查找某列所在的列名
def find_column_letter(worksheet, column_value):
    return pyxl.utils.get_column_letter(find_column_index(worksheet, column_value))


def my_range(start, end):
    while start <= end:
        yield start
        start += 1


# 生成Formula、FormulaString、Type列
def generate_formula_string(sheet_value):
    init_values = {}
    for value_type in value_types:
        sheet_value_with_nan = sheet_value[value_type].dropna()
        sheet_value_with_nan.loc[len(sheet_value_with_nan)] = np.nan
        init_values[value_type] = sheet_value_with_nan

    writer = pd.ExcelWriter('./generate.xlsx', engine='xlsxwriter')

    for i in range(len(sheet_formula)):
        formula_type = sheet_formula.loc[i]["Type"]
        print("...begin generate formula string: " + formula_type)
        formula_min = sheet_formula.loc[i]["Min"]
        formula_max = sheet_formula.loc[i]["Max"]
        row_datas = []
        for num in my_range(formula_min, formula_max):
            if num == 1:
                for value_type in value_types:
                    for init_value in init_values[value_type]:
                        row_data = {"Formula": formula_type,
                                    "ExpectedResult": 0,
                                    "FormulaString": formula_type + "([" + value_type + "1])",
                                    value_type + "1": init_value}
                        row_datas.append(row_data)
            elif num == 2:
                for value_type1 in value_types:
                    for value_type2 in value_types:
                        for init_value1 in init_values[value_type1]:
                            for init_value2 in init_values[value_type2]:
                                row_data = {"Formula": formula_type,
                                            "ExpectedResult": 0,
                                            "FormulaString": formula_type + "([" + value_type1 + "1], [" + value_type2 + "2])",
                                            value_type1 + "1": init_value1,
                                            value_type2 + "2": init_value2}
                                row_datas.append(row_data)
            elif num == 3:
                for value_type1 in value_types:
                    for value_type2 in value_types:
                        for value_type3 in value_types:
                            for init_value1 in init_values[value_type1]:
                                for init_value2 in init_values[value_type2]:
                                    for init_value3 in init_values[value_type3]:
                                        row_data = {"Formula": formula_type,
                                                    "ExpectedResult": 0,
                                                    "FormulaString": formula_type + "([" + value_type1 + "1], ["
                                                                     + value_type2 + "2], [" + value_type3 + "3])",
                                                    value_type1 + "1": init_value1,
                                                    value_type2 + "2": init_value2,
                                                    value_type3 + "3": init_value3}
                                        row_datas.append(row_data)
        df = pd.DataFrame(row_datas)
        df.to_excel(writer, sheet_name=formula_type, index=False)
    writer._save()


# 填充ExpectedResult数据
def generate_expected_result():
    types_to_format = {"Text": numbers.FORMAT_TEXT,
                       "Integer": numbers.FORMAT_NUMBER,
                       "Decimal": numbers.FORMAT_NUMBER_00,
                       "Date": numbers.FORMAT_DATE_DATETIME,
                       "Time": numbers.FORMAT_DATE_TIME6}
    workbook = load_workbook('./generate.xlsx')
    for i in range(len(sheet_formula)):
        formula_type = sheet_formula.loc[i]["Type"]
        print("...begin generate expected result: " + formula_type)
        formula_max = sheet_formula.loc[i]["Max"]
        worksheet = workbook[formula_type]

        if formula_max == 0:
            continue

        # 修改单元格格式
        for idx in range(int(formula_max)):
            for value_type in value_types:
                column_letter = find_column_letter(worksheet, value_type + str(idx + 1))
                for cell in worksheet[column_letter]:
                    cell.number_format = types_to_format[value_type]
        # 填充ExpectedResult
        column_name_to_letter = {}
        column_expected_result_letter = find_column_letter(worksheet, "ExpectedResult")
        column_formula_string_letter = find_column_letter(worksheet, "FormulaString")
        for col_idx in range(len(worksheet['A'])):
            if col_idx == 0:
                continue
            col_idx_string = str(col_idx + 1)
            formula_string = worksheet[column_formula_string_letter + col_idx_string].value
            results = re.findall(r"\[.*?\]", formula_string)
            expected_result = formula_string
            for result in results:
                if result[1:-1] in column_name_to_letter:
                    formula_letter = column_name_to_letter[result[1:-1]]
                else:
                    formula_letter = find_column_letter(worksheet, result[1:-1])
                    column_name_to_letter[result[1:-1]] = formula_letter
                expected_result = expected_result.replace(result, formula_letter + col_idx_string)
            worksheet[column_expected_result_letter + col_idx_string] = "=" + expected_result
    workbook.save(testFilePath)


# 删除公式结果为#VALUE!的行
# 这两个库好像计算不出正确数据
def delete_invalid_rows():
    workbook = load_workbook(testFilePath, data_only=True)
    for i in range(len(sheet_formula)):
        formula_type = sheet_formula.loc[i]["Type"]
        print("...begin delete invalid rows: " + formula_type)
        worksheet = workbook[formula_type]
        column_expected_result_index = find_column_index(worksheet, "ExpectedResult")

        rows_to_delete = []
        for row in worksheet.iter_rows(min_row=2):  # 从第二行开始，忽略标题行
            if row[column_expected_result_index-1].value is None:
                rows_to_delete.append(row)

        for row_to_delete in rows_to_delete:
            worksheet.delete_rows(row_to_delete[0].row)


def generate_test_file(filepath):
    global sheet_formula
    sheet_formula = pd.read_excel(filepath, sheet_name="Formula")
    sheet_value = pd.read_excel(filepath, sheet_name="Value")

    # generate_formula_string(sheet_value)
    # generate_expected_result()
    abs = pd.read_excel(testFilePath)


if __name__ == '__main__':
    generate_test_file('./init.xlsx')
