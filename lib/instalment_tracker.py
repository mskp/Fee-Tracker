import pandas as pd
from .course_instalments import get_instalments_per_course


class InstalmentTracker:

    YEAR_DICT = {
        "YR1": 1,
        "YR2": 2,
        "YR3": 3
    }

    SEM_DICT = {
        "SEM1": 'I',
        "SEM2": 'II',
        "SEM3": 'III',
        "SEM4": 'IV',
        "SEM5": 'V',
        "SEM6": 'VI',
        "SEM7": 'VII',
        "SEM8": 'VIII'
    }

    def __init__(self, year: int) -> None:
        self.INSTALMENTS_PER_COURSE = get_instalments_per_course(year)

    @staticmethod
    def check_substring_in_list(list_of_strings, substring) -> bool:
        for string in list_of_strings:
            if substring in string:
                return True
        return False

    @staticmethod
    def get_not_paid(payable_instalments: list[str], paid: list[str]) -> list[str]:
        return [i for i in payable_instalments if not InstalmentTracker.check_substring_in_list(paid, i)]

    def find_defaulter(self, df: pd.DataFrame, course: str, year_sem: str, instalment_num: int) -> pd.DataFrame:
        """
        Function that returns a dataframe of defaulters."""

        RESULTANT_COLUMNS = ["EXAM RNO", "MCODE", "NAME", "PH1",
                             "PH2", "COURSE", "YEAR/SEM", "UNPAID", 'PAID']

        payable_instalments = self.INSTALMENTS_PER_COURSE.get(
            course).get(year_sem)

        new_df = df.groupby('MCODE')['TRANSACTION'].apply(
            list).reset_index(name='PAID')

        not_paid = new_df.PAID.apply(
            lambda x: InstalmentTracker.get_not_paid(payable_instalments, x))

        new_df["UNPAID"] = not_paid.apply(", ".join)
        new_df["YEAR/SEM"] = InstalmentTracker.YEAR_DICT.get(year_sem) if InstalmentTracker.YEAR_DICT.get(
            year_sem) else InstalmentTracker.SEM_DICT.get(year_sem)
        new_df['PAID'] = new_df['PAID'].str.join(", ")
        df = df.merge(new_df, on="MCODE", how="right",
                      suffixes=('', '_remove'))

        condition1 = df.COURSE == course
        condition2 = df['YEAR/SEM'] == InstalmentTracker.SEM_DICT.get(year_sem)
        condition3 = df['YEAR/SEM'] == InstalmentTracker.YEAR_DICT.get(
            year_sem)
        condition4 = df["UNPAID"].str.contains(
            payable_instalments[instalment_num])

        final_df = df[condition1 & (condition2 | condition3) & condition4][RESULTANT_COLUMNS].drop_duplicates(
            subset='MCODE').sort_values("NAME")
        return final_df

    def find_defaulter_by_year_sem(self, df: pd.DataFrame, year_sem: str) -> pd.DataFrame:
        dfs = list()
        courses = {
            k: v for (k, v) in self.INSTALMENTS_PER_COURSE.items() if year_sem in v}

        try:
            for course in courses.keys():
                payable_instalments = self.INSTALMENTS_PER_COURSE.get(
                    course).get(year_sem)
                for instalment_number in range(len(payable_instalments)):
                    defaulters_df = self.find_defaulter(
                        df, course, year_sem, instalment_number)
                    dfs.append(
                        defaulters_df) if not defaulters_df.empty else ...

            dfs = pd.concat(dfs).drop_duplicates(
            ).sort_values("COURSE")
            return dfs
        except Exception as e:
            print(e)
            return None

    def find_defaulter_by_course(self, df: pd.DataFrame, course: str) -> pd.DataFrame:
        dfs = list()
        years_sems = self.INSTALMENTS_PER_COURSE.get(course).keys()
        try:
            for year_sem in years_sems:
                payable_instalments = self.INSTALMENTS_PER_COURSE.get(
                    course).get(year_sem)
                for instalment_number in range(len(payable_instalments)):
                    defaulters_df = self.find_defaulter(
                        df, course, year_sem, instalment_number)
                    dfs.append(
                        defaulters_df) if not defaulters_df.empty else ...
            dfs = pd.concat(dfs).drop_duplicates(
            ).sort_values("YEAR/SEM")
            return dfs
        except Exception as e:
            print(e)
            return None

    def find_all_defaulters(self, df: pd.DataFrame) -> pd.DataFrame:
        dfs = list()
        courses = self.INSTALMENTS_PER_COURSE.keys()

        for course in courses:
            years_sems = self.INSTALMENTS_PER_COURSE.get(course).keys()
            for year_sem in years_sems:
                payable_instalments = self.INSTALMENTS_PER_COURSE.get(
                    course).get(year_sem)
                for instalment_number in range(len(payable_instalments)):
                    defaulters_df = self.find_defaulter(
                        df, course, year_sem, instalment_number)
                    dfs.append(
                        defaulters_df) if not defaulters_df.empty else ...

        dfs = pd.concat(dfs).drop_duplicates(
        ).sort_values("COURSE")
        return dfs
