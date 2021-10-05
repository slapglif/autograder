import warnings
import pandas as pd
from datetime import datetime
import os

warnings.simplefilter(action='ignore', category=FutureWarning)
solution_file = "submissions/Accounting Simulation Solution For Automated Grading.xlsm"
# solution_file = "submissions/hhf_ers1.xlsm"


def parse_files():
    for file in os.listdir(r'./solutions'):
        if file.endswith(".xlsm"):
            # TODO: parse logic here
            pass


class ParseWorkbooks:
    def __init__(self):
        self.workbooks = self.build_workbooks_dict()
        wb = pd.read_excel(solution_file, sheet_name="Solution")
        self.solution_workbook = self.parse_workbook(wb, 4)
        self.student = self.extract_student_info(wb)
        self.compare_dates = self.build_workbook_dates()

    def parse_workbook(self, workbook: pd.DataFrame, header_row: int) -> pd.DataFrame:
        headers = workbook.iloc[header_row]
        return pd.DataFrame(workbook.values[header_row:], columns=headers)

    def extract_student_info(self, workbook: pd.DataFrame) -> dict:
        df_data = workbook.T.values.tolist()
        first_name = df_data[1][0]
        last_name = list(workbook.to_dict().keys())[1]
        student_id = df_data[1][1]
        data = {
            "first_name": first_name,
            "last_name": last_name,
            "student_id": student_id
        }
        return data

    def build_workbook(self, year: int) -> pd.DataFrame:
        wb = pd.read_excel(solution_file, sheet_name=f"Dec 31 {year}")
        rebuilt_workbook = self.parse_workbook(wb, 0)
        return rebuilt_workbook

    def build_workbook_dates(self) -> dict:
        return {x: datetime(x, 12, 31) for x in range(2013, 2019)}

    def build_workbooks_dict(self) -> dict:
        return {x: self.build_workbook(x) for x in range(2013, 2019)}

    def build_age_groups(self, year: int) -> list:
        age_groups = list()
        workbook_date = self.compare_dates[year]
        dates_paid = self.workbooks[year]["date purchased"].values.tolist()[1:]

        deltas = list()
        for index, date_paid in enumerate(dates_paid):
            if pd.isna(date_paid):
                age_groups.append(4)
                continue
            time_delta =  workbook_date - date_paid
            deltas.append(time_delta)
            if time_delta.days <= 30:
                age_groups.append(1)
            if time_delta.days >= 30 and time_delta.days < 60:
                age_groups.append(2)
            if time_delta.days >= 60 and time_delta.days < 90:
                age_groups.append(3)
            if time_delta.days >= 90:
                age_groups.append(4)
        return age_groups

    def percentage_of_sales(self, year: int) -> dict:
        workbook = self.workbooks[year]
        age_groups = self.build_age_groups(year)
        workbook["Age Groups"] = age_groups[len(workbook)]
        workbook_na = workbook[workbook["date paid"].isna()].groupby("Age Group")[["amount"]].sum()
        workbook_total = workbook.groupby("Age Group")[["amount"]].sum()
        sales = sum(workbook["total sales to customer"].values.tolist()[1:])
        total_percent = sum(workbook_na["amount"]) / sales
        data = {"total": round(total_percent, 3)}
        for x in range(1, 5):
            try:
                percentage = round(workbook_na["amount"][x] / workbook_total["amount"][x], 3)
                data.update({x: percentage})
            except KeyError as e:
                data.update({x: 0.0})
        return data

    def total_receivables(self, year: int) -> dict:
        return self.workbooks[year].groupby("Age Group")["amount"].sum().to_dict()

    def total_written_off(self, year: int) -> dict:
        workbook = self.workbooks[year]
        workbook_isna = workbook[workbook["date paid"].isna()].groupby("Age Group")[["amount"]].sum().to_dict()
        return workbook_isna["amount"]

    def total_outstanding(self) -> dict:
        data = dict()
        for year in range(2013, 2019):
            workbook = self.workbooks[year]
            total_by_year = {year: workbook.groupby("Age Group")[["amount"]].sum().to_dict()["amount"]}
            data.update(total_by_year)
        year_total = dict()
        for year in data:
            yearly_data = list()
            for x in data[year]:
                if isinstance(data[year][x], int):
                    yearly_data.append(data[year][x])
            year_total.update({year: sum(yearly_data)})
        return year_total


# Tests
parser = ParseWorkbooks()

print(parser.student)
column_headers = [
          "Year",
          "Total outstanding receivables",
          "Percent of sales on account written off",
          "0-30",
          "30-60",
          "61-90",
          "over 90"
      ]
workbook = parser.solution_workbook.drop(1)
workbook.columns = column_headers
solution = workbook.iloc[1: , :]

# build student data dict
student_answers = solution.groupby("Year")[
    "Total outstanding receivables",
    "Percent of sales on account written off",
    "0-30",
    "30-60",
    "61-90",
    "over 90"
].sum().to_dict()

# clean student data from df
print(student_answers)
for key in student_answers:
    student_answers[key][2018] = student_answers[key].pop("2018 (predicted)")
    for year in range(2013, 2019):
        student_answers[key][year] = round(student_answers[key][year], 3)

# default grades to fail
header_to_check = column_headers[1:]
student_grades = {
    "name": parser.student,
    "grades": {
        n: {
            header: "Fail" for header in header_to_check
        } for n in range(2013, 2019)
    }
}

# grade workbooks
for year in range(2013, 2018):
    percentage_of_sales = parser.percentage_of_sales(year)
    yearly_total_outstanding = parser.total_outstanding()
    if student_answers["Total outstanding receivables"][year] == yearly_total_outstanding[year]:
        student_grades["grades"][year]["Total outstanding receivables"] = "Pass"
    if student_answers["Percent of sales on account written off"][year] == percentage_of_sales["total"]:
        student_grades["grades"][year]["Percent of sales on account written off"] = "Pass"
    if student_answers["0-30"][year] == percentage_of_sales[1]:
        student_grades["grades"][year]["0-30"] = "Pass"
    if student_answers["30-60"][year] == percentage_of_sales[2]:
        student_grades["grades"][year]["30-60"] = "Pass"
    if student_answers["61-90"][year] == percentage_of_sales[3]:
        student_grades["grades"][year]["61-90"] = "Pass"
    if student_answers["over 90"][year] == percentage_of_sales[4]:
        student_grades["grades"][year]["over 90"] = "Pass"
if student_answers["Total outstanding receivables"][2018] == parser.total_outstanding()[2018]:
    student_grades["grades"][2018]["Total outstanding receivables"] = "Pass"

pd.set_option("display.max_rows", None, "display.max_columns", None)
print(pd.DataFrame(student_grades["grades"]))