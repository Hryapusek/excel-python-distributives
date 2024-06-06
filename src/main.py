
import pandas as pd
from loguru import logger

from distributives.distributive import Distributive
from distributives.distributive_searcher import DistributiveSearcher
from saver.saver import Saver


def read_name_surname() -> str:
    name = input("Введите ФИО пользователя: ").strip()
    return name

def read_distributives() -> list[str]:
    while True:
        distrs = input("Введите дистрибутивы: ").strip().split()
        if not distrs:
            print("Вы не ввели ни одного дистрибутива. Попробуйте снова")
            continue
        break
    distrs = [x.upper() for x in distrs]
    return distrs

def main():
    excel_file_name = "base.xlsx"
    while True:
        excel_file = pd.ExcelFile(excel_file_name)
        searcher = DistributiveSearcher(excel_file)
        print("\n" + "-"*15 + "\n")
        distrs = read_distributives()
        name_surname = read_name_surname()
        main_distr_str = distrs[0]
        dop_distrs_str = distrs[1:]
        print("Начинаю искать дистрибутивы в таблице")
        main_distr = searcher.get_main_distributive(main_distr_str)
        if not main_distr:
            logger.warning(f"Основной дистрибутив с именем \"{main_distr_str}\" не найден")
            continue
        logger.info(f"Нашел главный дистрибутив {main_distr.number} в листе {main_distr.sheet_name}")
        dop_distrs: list[Distributive] = []
        for dop_distr_str in dop_distrs_str:
            dop_distr = searcher.get_dop_distributive(dop_distr_str)
            if not dop_distr:
                logger.warning(f"Дополнительный дистрибутив с именем \"{dop_distr_str}\" не найден")
                continue
            dop_distrs.append(dop_distr)

        main_distr.user_name = name_surname
        saver = Saver(excel_file)
        saver.save_to_excel_file(excel_file_name, main_distr, dop_distrs)

        print(f"ФИО: {name_surname}")
        print(main_distr.number, *[x.number for x in dop_distrs], sep='; ')
        print("Файл успешно обновлен!")


if __name__ == "__main__":
    main()
