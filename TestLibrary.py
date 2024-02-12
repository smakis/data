import os
import requests
import datetime
import json
import pandas as pd
import nameparser as np
from nameparser.config import CONSTANTS
from pprint import pprint
from pandas import ExcelWriter
import openpyxl


class TestLibrary:
    def get_data(self, request_url: str) -> str:
        if os.path.exists("data.json"):
            with open("data.json", "r") as data_file:
                data = json.load(data_file)
            self.add_firstname_lastname(data)
            return data
        else:
            response = requests.get(request_url)
            data = response.json()
            json_data = json.dumps(data)
            with open("data.json", "w") as data_file:
                data_file.write(json_data)
            self.add_firstname_lastname(data)
            return data

    def add_firstname_lastname(self, data: list):
        """Add firstname and lastname to employee data. Suffixes are included in lastname.
        Titles are removed."""
        CONSTANTS.string_format = "{title} {first} {last} {suffix}"
        name_nr = 0
        for item in data:
            if item := item["name"]:
                name = np.HumanName(item)
                data[name_nr]["firstname"] = name.first
                lastname = "".join(name.last + " " + name.suffix).rstrip()
                data[name_nr]["lastname"] = lastname
                name_nr += 1

    def create_dataframe(self, data: list) -> pd.DataFrame:
        """Create a dataframe with wanted information and return dataframe"""
        df = pd.DataFrame(data)
        column_titles = [
            "lastname",
            "firstname",
            "email",
            "street",
            "city",
            "zipcode",
            "phone",
            "website",
        ]
        df = df.drop(columns=["company", "username", "id"])

        # Address column is removed and new columns are created from address columns data
        df_flat = df.join(pd.json_normalize(df["address"])).drop(
            "address", axis="columns"
        )
        df_flat = df_flat.drop(columns=["geo.lat", "suite", "geo.lng"])
        df_final = df_flat.drop(columns=["name"])
        # Order columns based on column_titles
        df_final = df_final.reindex(columns=column_titles)
        df_final = df_final.sort_values(
            ["lastname", "firstname"], ascending=[True, True]
        )
        return df_final

    def save_to_excel(self, dataframe: pd.DataFrame, filepath=None):
        """By default excel file is saved to project folder.
        Excel can be saved to specified filepath using optional filepath argument.
        If folder does not exist it is created.
        """
        timestamp = self.timestamp_now()
        filename = f"employees_{timestamp}.xlsx"
        if filepath is None:
            self.write_excel(filename, dataframe)

        elif not os.path.exists(filepath):
            os.makedirs(filepath)
            file = os.path.join(filepath, filename)
            self.write_excel(file, dataframe)

        else:
            file = os.path.join(filepath, filename)
            self.write_excel(file, dataframe)

    def write_excel(self, file, dataframe):
        writer = ExcelWriter(file)
        dataframe.to_excel(writer, index=False)
        writer.save()

    def timestamp_now(self) -> str:
        now = datetime.datetime.now()
        format = "%Y%m%d%H%M%S"
        return now.strftime(format)


if __name__ == "__main__":
    test_library = TestLibrary()
    data = test_library.get_data("https://jsonplaceholder.typicode.com/users")
    data = test_library.create_dataframe(data)
    # Change the specified folder if wanting to save elsewhere
    excel_save_file_to = r".\Data"
    # Give save to excel "excel_with_file_path" if want to save to folder.
    # Otherwise will save to current directory
    test_library.save_to_excel(data, excel_save_file_to)
