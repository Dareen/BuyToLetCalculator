import math

from exceptions import ExceededMaxDownpayment
from exceptions import ExceededMaxMonthlyInstallment

from constants import *


INTEREST_RATE_DECIMAL = BANK_INTEREST_RATE/100.00
MONTHLY_INTEREST_RATE = INTEREST_RATE_DECIMAL/ MONTHS # 0.00625


class Permutation(object):
    
    def __init__(self, parent_prop, num_years, downpayment_percent):

        self.parent_prop = parent_prop
        self.num_years = int(num_years)
        self.downpayment_percent = int(downpayment_percent)

        if downpayment_percent == 100:
            self.num_years = 0
            self.monthly_installments = 0.0

    def calculate(self):

        self.downpayment = self.parent_prop.price * (self.downpayment_percent/100.00)

        if self.downpayment > MAX_DOWNPAYMENT:
            raise ExceededMaxDownpayment()

        self.loan_principal = self.parent_prop.price - self.downpayment
        self.num_payments = self.num_years * MONTHS

        self.calc_loan_fees()
        self.calc_monthly_installment()

        self.total_monthly_outcome = self.parent_prop.expense_total_monthly + self.monthly_installments
        self.total_monthly_income = self.parent_prop.rent - self.total_monthly_outcome

        self.total_annual_income = self.total_monthly_income * MONTHS

        self.equity = self.downpayment + self.total_fees
        self.monthly_ROI = self.total_monthly_income / self.equity
        self.annual_ROI = self.total_annual_income / self.equity

        self.calculate_afterloan_roi()
        self.calc_x_years_roi(period_years=10)

        return self.annual_ROI

    def calculate_afterloan_roi(self):

        self.equity_with_loan = self.equity
        if self.total_annual_income <= 0.00:
            # minus because total_annual_income is already negative
            self.equity_with_loan = self.equity - (self.total_annual_income * self.num_years)

        self.afterloan_annual_roi = self.parent_prop.afterloan_annual_income / self.equity_with_loan
        return self.afterloan_annual_roi

    def calc_loan_fees(self):
        self.loan_giving_fee = self.loan_principal * BANK_LOAN_GIVING_FEE
        self.mortgage_fee = self.loan_principal * BANK_MORTGAGE_FEE
        self.loan_stamps_fee = self.loan_principal * BANK_LOAN_STAMPS
        self.total_fees = self.loan_giving_fee + self.mortgage_fee + self.loan_stamps_fee

    def calc_monthly_installment(self):
        # https://mortgage.lovetoknow.com/Calculate_Mortgage_Payments_Formula
        if self.downpayment_percent == 100:
            return

        self.term = math.pow(1 + MONTHLY_INTEREST_RATE, self.num_payments)
        self.monthly_installments = self.loan_principal * ((MONTHLY_INTEREST_RATE * self.term) / (self.term - 1))

        if self.monthly_installments > MAX_MONTHLY_INSTALLMENT:
            raise ExceededMaxMonthlyInstallment()

    def calc_x_years_roi(self, period_years):

        self.equity_x_years = self.equity
        loan_period_income = 0.0
        if self.total_annual_income <= 0.00:
            # minus because total_annual_income is already negative
            self.equity_x_years = self.equity - (self.total_annual_income * self.num_years)
        else:
            loan_period_income = self.total_annual_income * min(period_years, self.num_years)

        after_loan_years = period_years - self.num_years
        after_loan_income_within_x_years = self.parent_prop.afterloan_annual_income * after_loan_years

        if after_loan_income_within_x_years < 0:
            after_loan_income_within_x_years = 0

        self.total_income_x_years = after_loan_income_within_x_years + loan_period_income

        self.x_years_avg_annual_roi = (self.total_income_x_years / period_years) / self.equity_x_years


    def __repr__(self):
        return '''num_years: {num_years}      |      downpayment_percent: {downpayment_percent}%
monthly_installments: {monthly_installments:.2f}      |      downpayment: {downpayment:.2f}
loan_principal: {loan_principal:.2f}      |      num_payments: {num_payments:.2f}
total_loan_fees: {total_fees:.2f}      |      total_monthly_outcome: {total_monthly_outcome:.2f}
total_monthly_income: {total_monthly_income:.2f}      |      total_annual_income: {total_annual_income:.2f}
equity: {equity:.2f}      |      annual_ROI: {annual_ROI:.2f}
  ***
afterloan_annual_roi: {afterloan_annual_roi:.2f}      |      equity_with_loan: {equity_with_loan:.2f}
  ***
equity_x_years: {equity_x_years:.2f}      |      x_years_avg_annual_roi: {x_years_avg_annual_roi:.2f}
'''.format(**self.__dict__)


class ShortTermPermutation(Permutation):
    def calculate(self):
        return super(ShortTermPermutation, self).calculate()
        # return self.afterloan_annual_roi


# TODO - remove this
class PermutationFactory(object):

    def create(self, parent_prop, num_years, downpayment_percent):
        if num_years <= 5:
            return ShortTermPermutation(parent_prop, num_years, downpayment_percent)
        else:
            return Permutation(parent_prop, num_years, downpayment_percent)


class ExceededMaxDownpaymentPermutation(Permutation):
    def __repr__(self):
        return 'Exceeded Max Downpayment {}'.format(MAX_DOWNPAYMENT)


class ExceededMaxMonthlyInstallmentPermutation(Permutation):
    def __repr__(self):
        return 'Exceeded Max Monthly Installment {}'.format(MAX_MONTHLY_INSTALLMENT)

