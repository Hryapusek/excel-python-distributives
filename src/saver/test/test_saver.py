import unittest

import pandas as pd

from distributives.distributive import Distributive
from distributives.distributive_searcher import DistributiveSearcher
from saver.saver import Saver


class TestSaver(unittest.TestCase):
    def test_save(self):
        excel_file = pd.ExcelFile("base.xlsx")
        searcher = DistributiveSearcher(excel_file)
        main_distr = searcher.get_main_distributive("СКУО")
        main_distr.user_name = "Автор"
        self.assertIsNotNone(main_distr)
        saver = Saver(excel_file)
        saver.save_to_excel_file("test_file.xlsx", main_distr, [Distributive("123123", name="Dop distr"), Distributive("234234", name="Dop distr")])

    def test_save_mz(self):
        excel_file = pd.ExcelFile("base.xlsx")
        searcher = DistributiveSearcher(excel_file)
        main_distr = searcher.get_main_distributive("СКЗО")
        main_distr.user_name = "Автор"
        self.assertIsNotNone(main_distr)
        saver = Saver(excel_file)
        saver.save_to_excel_file("test_file.xlsx", main_distr, [Distributive(3436, name="ПРАС"), Distributive(3424, name="ПРАС")])

