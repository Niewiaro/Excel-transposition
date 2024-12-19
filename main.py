from openpyxl import load_workbook


class Config:
    def __init__(self) -> None:
        from pathlib import Path

        self.input_name = "Profile 2025 GD.xlsx"
        self.input_path = Path.cwd() / self.input_name


def main() -> None:
    config = Config()
    workbook = load_workbook(config.input_path)
    sheets_names = workbook.sheetnames

    for sheet in sheets_names:
        print(sheet)


if __name__ == "__main__":
    main()
