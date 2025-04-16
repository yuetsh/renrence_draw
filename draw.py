import openpyxl
import pandas as pd


def draw_one(filepath, count):
    wb = openpyxl.load_workbook(filepath)
    course_count = len(wb.sheetnames)

    result = pd.DataFrame()

    first_sheet = pd.read_excel(filepath, sheet_name=0, skiprows=1)
    top_school = first_sheet["学校"].value_counts().idxmax()

    for i in range(course_count):
        df = pd.read_excel(filepath, sheet_name=i, skiprows=1)
        df = pd.concat([df, result]).drop_duplicates("身份证号", keep=False)

        a = df[df["等级"] == "A"]
        b = df[df["等级"] == "B"]
        c = df[df["等级"] == "C"]

        random_a = a.sample(n=count)
        random_b = b.sample(n=count)
        random_c = c.sample(n=count)

        course = pd.concat([random_a, random_b, random_c])
        result = pd.concat([result, course])

    result["序号"] = range(1, 1 + len(result))
    return result


def draw_two(filepath):
    wb = openpyxl.load_workbook(filepath)
    course_count = len(wb.sheetnames)

    result = pd.DataFrame()

    first_sheet = pd.read_excel(filepath, sheet_name=0, skiprows=1)
    top_school = first_sheet["学校"].value_counts().idxmax()

    for i in range(course_count):
        df = pd.read_excel(filepath, sheet_name=i, skiprows=1)
        df = pd.concat([df, result]).drop_duplicates("身份证号", keep=False)

        """
        学校（人数多的） ———— 15
        学校（人数少的） ———— 9

        A8 B8 C8
        """

        for school, group in df.groupby("学校"):
            a = group[group["等级"] == "A"]
            b = group[group["等级"] == "B"]
            c = group[group["等级"] == "C"]

            n = 5 if school == top_school else 3

            random_a = a.sample(n=n)
            random_b = b.sample(n=n)
            random_c = c.sample(n=n)

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
        df = pd.read_excel(filepath, sheet_name=i, skiprows=1)
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
