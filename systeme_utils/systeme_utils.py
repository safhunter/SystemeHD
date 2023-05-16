import pandas as pd
import os
import json
class JsonConverter:
    def __init__(self, file_path: str, file_name: str):
        self.file_path = file_path
        if file_name:
            self.file_name = file_name
        else:
            raise ValueError

    def convert(self):
        try:
            df_config = pd.read_excel(os.path.join(self.file_path, self.file_name), header=1, usecols="A:B")
        except(IOError, SystemError) as ex:
            print(f"Can't open a file: {self.file_name}. Cause:\n{ex}")
            return None

        name, ext = os.path.splitext(self.file_name)
        df_dict = {}
        for row in df_config.itertuples(index=False, name=None):
            df_dict[row[0]] = row[1]

        with open(f'{os.path.join(self.file_path, name)}.json', 'w') as outfile:
            json.dump(df_dict, outfile)

class PlatformConverter:
    def __init__(self, file_path: str, file_name: str):
        self.file_path = file_path
        if file_name:
            self.file_name = file_name
        else:
            raise ValueError

    def convert(self):
        try:
            df_config = pd.read_excel(os.path.join(self.file_path, self.file_name), header=1, usecols="A:D")
        except(IOError, SystemError) as ex:
            print(f"Can't open a file: {self.file_name}. Cause:\n{ex}")
            return None

        df_list = []
        for row in df_config.itertuples(index=False, name=None):
            df_list.extend(self.parce_row(row))

        name, ext = os.path.splitext(self.file_name)
        df = pd.DataFrame(df_list, index=None, columns=["Сигнал", "Описание", "Тип", "Привязка", "Тип объекта", "Экземпляр объекта", "Свойство объекта", "Индекс", "Протокольный тип"])
        df.to_excel(f'{os.path.join(self.file_path, name)}_plat.xlsx', sheet_name='BACnetAddressMap', index=False)

    @staticmethod
    def parce_row(row :tuple):
        area = row[1][:2]
        line_1 = []
        line_2 = []
        try:
            int(row[1][2:])
        except ValueError:
            print(f"Incorrect row: {row[0]} {row[1]}")
            return tuple()

        if area == "BI":
            line_1.append(row[0] + ".PresentValue")
            line_1.append(row[3])
            line_1.append("uint4")
            line_1.append("непосредственно")
            line_1.append("Binary Input")
            line_1.append(row[1][2:])
            line_1.append("PRESENT_VALUE")
            line_1.append("")
            line_1.append("Enum")

            line_2.append(row[0] + ".OutOfService")
            line_2.append(row[3])
            line_2.append("bool")
            line_2.append("непосредственно")
            line_2.append("Binary Input")
            line_2.append(row[1][2:])
            line_2.append("OUT_OF_SERVICE")
            line_2.append("")
            line_2.append("BOOLEAN")
        elif area == "BO":
            line_1.append(row[0] + ".PresentValue")
            line_1.append(row[3])
            line_1.append("uint4")
            line_1.append("непосредственно")
            line_1.append("Binary Output")
            line_1.append(row[1][2:])
            line_1.append("PRESENT_VALUE")
            line_1.append("")
            line_1.append("Enum")

            line_2.append(row[0] + ".OutOfService")
            line_2.append(row[3])
            line_2.append("bool")
            line_2.append("непосредственно")
            line_2.append("Binary Output")
            line_2.append(row[1][2:])
            line_2.append("OUT_OF_SERVICE")
            line_2.append("")
            line_2.append("BOOLEAN")
        elif area == "BV":
            line_1.append(row[0] + ".PresentValue")
            line_1.append(row[3])
            line_1.append("uint4")
            line_1.append("непосредственно")
            line_1.append("Binary Value")
            line_1.append(row[1][2:])
            line_1.append("PRESENT_VALUE")
            line_1.append("")
            line_1.append("Enum")
        elif area == "AI":
            line_1.append(row[0] + ".PresentValue")
            line_1.append(row[3])
            line_1.append("float")
            line_1.append("непосредственно")
            line_1.append("Analog Input")
            line_1.append(row[1][2:])
            line_1.append("PRESENT_VALUE")
            line_1.append("")
            line_1.append("REAL")

            line_2.append(row[0] + ".OutOfService")
            line_2.append(row[3])
            line_2.append("bool")
            line_2.append("непосредственно")
            line_2.append("Analog Input")
            line_2.append(row[1][2:])
            line_2.append("OUT_OF_SERVICE")
            line_2.append("")
            line_2.append("BOOLEAN")
        elif area == "AO":
            line_1.append(row[0] + ".PresentValue")
            line_1.append(row[3])
            line_1.append("float")
            line_1.append("непосредственно")
            line_1.append("Analog Output")
            line_1.append(row[1][2:])
            line_1.append("PRESENT_VALUE")
            line_1.append("")
            line_1.append("REAL")

            line_2.append(row[0] + ".OutOfService")
            line_2.append(row[3])
            line_2.append("bool")
            line_2.append("непосредственно")
            line_2.append("Analog Output")
            line_2.append(row[1][2:])
            line_2.append("OUT_OF_SERVICE")
            line_2.append("")
            line_2.append("BOOLEAN")
        elif area == "AV":
            line_1.append(row[0] + ".PresentValue")
            line_1.append(row[3])
            line_1.append("float")
            line_1.append("непосредственно")
            line_1.append("Analog Value")
            line_1.append(row[1][2:])
            line_1.append("PRESENT_VALUE")
            line_1.append("")
            line_1.append("REAL")

        if line_2:
            return line_1, line_2
        return (line_1, )