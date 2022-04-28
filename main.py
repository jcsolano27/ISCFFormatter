import os
from ExcelManager import ExcelManager


class ISCFFormatter:

    input = "iscf"
    output = "output"
    opened_bracket = False
    begin_flag = "BEGIN_"
    end_flag = "END_"
    separator = ","
    excel_manager = None

    def __init__(self):
        self.excel_manager = ExcelManager(self.output, self.separator)
        if not os.path.exists(self.input):
            os.mkdir(self.input)
        if not os.path.exists(self.output):
            os.mkdir(self.output)

    @staticmethod
    def get_bracket_name(bracket, flag):
        return (bracket.strip())[len(flag):]

    def read_path(self):
        path = self.input
        for f in os.listdir(path):
            if f.split(".")[1] == "ignore":
                continue
            self.excel_manager.open_file(f)
            self.read_file(f"{path}\\{f}")
            self.excel_manager.close_file()

    def read_file(self, file):
        title = ""
        with open(file, 'r') as f:
            for line in f:
                if self.begin_flag in line:
                    if self.opened_bracket:
                        raise Exception(f"ERROR: {self.begin_flag} found before closing the previous flag")
                    else:
                        self.opened_bracket = True
                        bracket_name = self.get_bracket_name(line, self.begin_flag)
                        self.excel_manager.add_sheet(bracket_name)
                        title = f.readline()
                        if self.end_flag in title:
                            self.opened_bracket = False
                            title = F"NO RECORDS FOUND BETWEEN {bracket_name} BRACKETS!"
                        self.excel_manager.add_row(title)
                elif self.end_flag in line:
                    if not self.opened_bracket:
                        raise Exception(f"ERROR: {self.end_flag} found before closing the previous flag!")
                    else:
                        self.opened_bracket = False
                else:
                    self.excel_manager.add_row(line)


def main():
    formatter = ISCFFormatter()
    formatter.read_path()


if __name__ == '__main__':
    main()
