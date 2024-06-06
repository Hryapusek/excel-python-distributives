import unittest

import pandas as pd

from distributives.distributive_searcher import DistributiveSearcher


class TestDistributiveSearcher(unittest.TestCase):
    def test_main_distr(self):
        excel_file = pd.ExcelFile("base.xlsx")
        searcher = DistributiveSearcher(excel_file)
        searcher.get_main_distributive("СКЮО")

    def test_dop_distr(self):
        excel_file = pd.ExcelFile("base.xlsx")
        searcher = DistributiveSearcher(excel_file)
        searcher.get_dop_distributive("ПРСОЮ")