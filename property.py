from collections import defaultdict

from permutation import ExceededMaxDownpaymentPermutation
from permutation import ExceededMaxMonthlyInstallmentPermutation
from permutation import LessThanMinmumLoanPrincipalPermutation
from permutation import Permutation
from permutation import PermutationFactory
from exceptions import ExceededMaxDownpayment
from exceptions import ExceededMaxMonthlyInstallment
from exceptions import LessThanMinmumLoanPrincipal

from constants import *


class Property(object):

    def __init__(self, price, rent, reletting_factor, gov_tax_discount, area, extra_onetime_expense, name, url):
        self.price = float(price)
        self.rent = float(rent)
        self.reletting_factor = float(reletting_factor)
        self.gov_tax_discount = float(gov_tax_discount)
        self.area = float(area)
        self.extra_onetime_expense = float(extra_onetime_expense)
        self.name = name
        self.url = url

        self.annual_rent = self.rent * MONTHS

        self.permutation_stats = defaultdict(lambda: defaultdict(Permutation))

        self.max_roi = MINIMUM_DECIMAL
        self.max_stats = None

        self.max_roi_x_years = MINIMUM_DECIMAL
        self.max_stats_x_years = None

    def process(self, min_num_years, max_num_year, min_downpayment_percent, max_downpayment_percent):

        factory = PermutationFactory()
        self.calc_expenses()

        for num_years in range(min_num_years, max_num_year+1):
            for downpayment_percent in range(min_downpayment_percent, max_downpayment_percent+1):

                permutation = factory.create(
                    parent_prop=self,
                    num_years=num_years,
                    downpayment_percent=downpayment_percent,
                )

                try:
                    roi = permutation.calculate()
                    self.permutation_stats[num_years][downpayment_percent] = permutation

                    if roi > self.max_roi:
                        self.max_roi = roi
                        self.max_stats = permutation

                    if permutation.x_years_avg_annual_roi > self.max_roi_x_years:
                        self.max_roi_x_years = permutation.x_years_avg_annual_roi
                        self.max_stats_x_years = permutation

                except ExceededMaxDownpayment:
                    self.permutation_stats[num_years][downpayment_percent] = ExceededMaxDownpaymentPermutation(
                        parent_prop=self,
                        num_years=num_years,
                        downpayment_percent=downpayment_percent,
                    )
                    break
                except ExceededMaxMonthlyInstallment:
                    self.permutation_stats[num_years][downpayment_percent] = ExceededMaxMonthlyInstallmentPermutation(
                        parent_prop=self,
                        num_years=num_years,
                        downpayment_percent=downpayment_percent,
                    )
                except LessThanMinmumLoanPrincipal:
                    self.permutation_stats[num_years][downpayment_percent] = LessThanMinmumLoanPrincipalPermutation(
                        parent_prop=self,
                        num_years=num_years,
                        downpayment_percent=downpayment_percent,
                    )

    def calc_expenses(self):
        self.expense_annual_govt_tax = (self.annual_rent * 0.80) * 0.15 * (1.0 - self.gov_tax_discount)
        self.expense_monthly_govt_tax = self.expense_annual_govt_tax / MONTHS
        self.expense_annual_reletting = self.rent * self.reletting_factor
        self.expense_monthly_reletting = self.expense_annual_reletting / MONTHS

        self.expense_total_monthly = (MONTHLY_OPERATING_EXPENSES + self.expense_monthly_govt_tax
            + self.expense_monthly_reletting + MONTHLY_MAINTAINENCE)

        self.afterloan_monthly_income = self.rent - self.expense_total_monthly
        self.afterloan_annual_income = self.afterloan_monthly_income * MONTHS


    def __repr__(self):
        return '\n=======Property {}/{} {}=======\n{}\n'.format(
            self.price, self.rent, self.name, self.url,
        )
