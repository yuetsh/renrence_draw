import openpyxl
import pandas as pd


def draw(filepath):
    wb = openpyxl.load_workbook(filepath)
    course_count = len(wb.sheetnames)

    result = pd.DataFrame()
    for i in range(course_count):
        df = pd.read_excel(filepath, sheet_name=i, skiprows=2)
        df = pd.concat([df, result]).drop_duplicates("身份证号", keep=False)

        a = df[df["等级"] == "A"]
        b = df[df["等级"] == "B"]
        c = df[df["等级"] == "C"]

        random_a = a.sample(n=5)
        random_b = b.sample(n=5)
        random_c = c.sample(n=5)

        course = pd.concat([random_a, random_b, random_c])
        result = pd.concat([result, course])

    result["序号"] = range(1, 1 + len(result))
    return result


"""
users 是已经在语数英中抽到的学生 pf
filepath 是专业课文件路径
"""


def draw_zhuanyeke(users, filepath):
    wb = openpyxl.load_workbook(filepath)
    course_count = len(wb.sheetnames)

    result = pd.DataFrame()
    for i in range(course_count):
        df = pd.read_excel(filepath, sheet_name=i, skiprows=2)
        # 把 df 中出现在 users 里面的数据去掉
        df = df[~df["身份证号"].isin(users["身份证号"])]

        a = df[df["等级"] == "A"]
        b = df[df["等级"] == "B"]
        c = df[df["等级"] == "C"]

        random_a = a.sample(n=2)
        random_b = b.sample(n=2)
        random_c = c.sample(n=2)

        course = pd.concat([random_a, random_b, random_c])
        result = pd.concat([result, course])

    result["序号"] = range(1, 1 + len(result))
    return result
