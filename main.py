from openpyxl import load_workbook
import logging

logging.basicConfig(
    level=logging.DEBUG, format="%(asctime)s - %(name)s - %(levelname)s - %(message)s"
)


class Config:
    def __init__(self) -> None:
        from pathlib import Path

        self.input_name = "Profile 2025 GD.xlsx"
        self.input_path = Path.cwd() / self.input_name


def df_transform():
    pass


def main() -> None:
    logging.info("Start")
    config = Config()
    workbook = load_workbook(config.input_path)
    sheets_names = workbook.sheetnames

    for sheet in sheets_names:
        print(sheet)


if __name__ == "__main__":
    main()
