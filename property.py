from collections import defaultdict

from permutation import PermutationFactory
from permutation import ExceededMaxDownpaymentPermutation
from permutation import Permutation
from exceptions import ExceededMaxDownpayment

from constants import *


class Property(object):

    def __init__(self, price, rent, area, extra_onetime_expense, name, url):
        self.price = float(price)
        self.rent = float(rent)
        self.area = float(area)
        self.extra_onetime_expense = float(extra_onetime_expense)
        self.name = name
        self.url = url

        self.annual_rent = self.rent * MONTHS

        self.permutation_stats = defaultdict(lambda: defaultdict(Permutation))

        self.local_max_roi = MINIMUM_DECIMAL
        self.local_max_stats = None

        self.local_max_roi_x_years = MINIMUM_DECIMAL
        self.local_max_stats_x_years = None

    def process(self, min_num_years, max_num_year, min_downpayment_percent, max_downpayment_percent):

        factory = PermutationFactory()

        for num_years in range(min_num_years, max_num_year+1):
            for downpayment_percent in range(min_downpayment_percent, max_downpayment_percent+1):

                permutation_stats = factory.create(
                    parent_prop=self,
                    num_years=num_years,
                    downpayment_percent=downpayment_percent,
                )

                try:
                    roi = permutation_stats.calculate()
                    self.permutation_stats[num_years][downpayment_percent] = permutation_stats

                    if roi > self.local_max_roi:
                        self.local_max_roi = roi
                        self.local_max_stats = permutation_stats

                    if permutation_stats.x_years_avg_annual_roi > self.local_max_roi_x_years:
                        self.local_max_roi_x_years = permutation_stats.x_years_avg_annual_roi
                        self.local_max_stats_x_years = permutation_stats

                except ExceededMaxDownpayment:
                    self.permutation_stats[num_years][downpayment_percent] = ExceededMaxDownpaymentPermutation(
                        parent_prop=self,
                        num_years=num_years,
                        downpayment_percent=downpayment_percent,
                    )
                    break

        print(self.local_max_stats)


    def __repr__(self):
        return '\n=======Property {}/{} {}=======\n{}\n'.format(self.price, self.rent, self.name, self.url)
