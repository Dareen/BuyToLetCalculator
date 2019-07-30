import math
from collections import defaultdict
from pprint import pprint
import csv

import xlsxwriter


MONTHS = 12.00

MAX_DOWNPAYMENT = 30000

# BANK PARAMS
BANK_INTEREST_RATE = 7.50
BANK_LOAN_GIVING_FEE = 0.01
BANK_MORTGAGE_FEE = 1.20 * 0.01
BANK_LOAN_STAMPS = 0.003


INTEREST_RATE_DECIMAL = BANK_INTEREST_RATE/100.00
MONTHLY_INTEREST_RATE = INTEREST_RATE_DECIMAL/ MONTHS # 0.00625

MONTHLY_OPERATING_EXPENSES = 30.00 # WIFI, Water
ANNUAL_MAINTAINENCE = 100.00
MONTHLY_MAINTAINENCE = ANNUAL_MAINTAINENCE / MONTHS

ROI_YEARS = [1, 3, 5, 10, 15, 20, 25, 30]
VISUALIZING_ROI_YEARS = 10

MINIMUM_DECIMAL = -99999999.99

class ExceededMaxDownpayment(Exception):
    pass


class PropStats(object):

    def __init__(self, price, rent, area, extra_onetime_expense, name, url):
        self.price = float(price)
        self.rent = float(rent)
        self.area = float(area)
        self.extra_onetime_expense = float(extra_onetime_expense)
        self.name = name
        self.url = url

        self.annual_rent = self.rent * MONTHS

        self.permutation_stats = defaultdict(lambda: defaultdict(PermutationStats))

        self.local_max_roi = MINIMUM_DECIMAL
        self.local_max_stats = None

        self.local_max_roi_x_years = MINIMUM_DECIMAL
        self.local_max_stats_x_years = None

    def process(self, min_num_years, max_num_year, min_downpayment_percent, max_downpayment_percent):

        factory = PermutationStatsFactory()

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
                    self.permutation_stats[num_years][downpayment_percent] = ExceededMaxDownpaymentPermutationStats(
                        parent_prop=self,
                        num_years=num_years,
                        downpayment_percent=downpayment_percent,
                    )
                    break

        print(self.local_max_stats)


    def __repr__(self):
        return '\n=======Property {}/{} {}=======\n{}\n'.format(self.price, self.rent, self.name, self.url)


class PermutationStats(object):
    
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

        self.cal_expenses()
        self.calc_loan_fees()
        self.calc_monthly_installment()

        self.total_monthly_outcome = self.expense_total_monthly + self.monthly_installments
        self.total_monthly_income = self.parent_prop.rent - self.total_monthly_outcome

        self.total_annual_income = self.total_monthly_income * MONTHS

        self.equity = self.downpayment + self.total_fees
        self.monthly_ROI = self.total_monthly_income / self.equity
        self.annual_ROI = self.total_annual_income / self.equity

        self.calculate_afterloan_roi()
        self.cal_x_years_roi(period_years=10)

        return self.annual_ROI

    def calculate_afterloan_roi(self):
        self.afterloan_monthly_income = self.parent_prop.rent - self.expense_total_monthly
        self.afterloan_annual_income = self.afterloan_monthly_income * MONTHS

        self.equity_with_loan = self.equity
        if self.total_annual_income <= 0.00:
            # minus because total_annual_income is already negative
            self.equity_with_loan = self.equity - (self.total_annual_income * self.num_years)

        self.afterloan_annual_roi = self.afterloan_annual_income / self.equity_with_loan
        return self.afterloan_annual_roi

    def cal_expenses(self):
        self.expense_annual_govt_tax = (self.parent_prop.annual_rent * 0.80) * 0.15
        self.expense_monthly_govt_tax = self.expense_annual_govt_tax / MONTHS
        self.expense_annual_reletting = self.parent_prop.rent * 3.0
        self.expense_monthly_reletting = self.expense_annual_reletting / MONTHS

        self.expense_total_monthly = (MONTHLY_OPERATING_EXPENSES + self.expense_monthly_govt_tax
            + self.expense_monthly_reletting + MONTHLY_MAINTAINENCE)

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

    def cal_x_years_roi(self, period_years=10):

        self.equity_x_years = self.equity
        loan_period_income = 0.0
        if self.total_annual_income <= 0.00:
            # minus because total_annual_income is already negative
            self.equity_x_years = self.equity - (self.total_annual_income * self.num_years)
        else:
            loan_period_income = self.total_annual_income * self.num_years

        if self.num_years < period_years:
            self.total_income_x_years = (self.afterloan_annual_income * (period_years - self.num_years)) + loan_period_income
        else:
            self.total_income_x_years = loan_period_income

        self.x_years_avg_annual_roi = (self.total_income_x_years / period_years) / self.equity_x_years


    def __repr__(self):
        return '''num_years: {num_years}      |      downpayment_percent: {downpayment_percent}%
monthly_installments: {monthly_installments:.2f}      |      downpayment: {downpayment:.2f}
loan_principal: {loan_principal:.2f}      |      num_payments: {num_payments:.2f}
expense_annual_govt_tax: {expense_annual_govt_tax:.2f}      |      expense_monthly_govt_tax: {expense_monthly_govt_tax:.2f}
expense_annual_reletting: {expense_annual_reletting:.2f}      |      expense_monthly_reletting: {expense_monthly_reletting:.2f}
expense_total_monthly: {expense_total_monthly:.2f}
total_loan_fees: {total_fees:.2f}      |      total_monthly_outcome: {total_monthly_outcome:.2f}
total_monthly_income: {total_monthly_income:.2f}      |      total_annual_income: {total_annual_income:.2f}
equity: {equity:.2f}      |      annual_ROI: {annual_ROI:.2f}
  ***
afterloan_monthly_income: {afterloan_monthly_income:.2f}      |      equity_with_loan: {equity_with_loan:.2f}
afterloan_annual_roi: {afterloan_annual_roi:.2f}
  ***
equity_x_years: {equity_x_years:.2f}      |      x_years_avg_annual_roi: {x_years_avg_annual_roi:.2f}
'''.format(**self.__dict__)


class ShortTermPermutationStats(PermutationStats):
    def calculate(self):
        super(ShortTermPermutationStats, self).calculate()
        return self.afterloan_annual_roi


class ExceededMaxDownpaymentPermutationStats(PermutationStats):
    def __repr__(self):
        return 'Exceeded Max Downpayment {}'.format(MAX_DOWNPAYMENT)


class PermutationStatsFactory(object):

    def create(self, parent_prop, num_years, downpayment_percent):
        if num_years <= 5:
            return ShortTermPermutationStats(parent_prop, num_years, downpayment_percent)
        else:
            return PermutationStats(parent_prop, num_years, downpayment_percent)

class LandlordPropertyCalculator(object):

    def __init__(self, input_path, file_format="csv"):
        self.input_path = input_path
        self.file_format = file_format

        self.global_max_roi = MINIMUM_DECIMAL
        self.global_max_stats = None


    def read_csv_input(self, input_path):
        # TODO - read from file
        # return SAMPLE_INPUT
        with open(input_path, 'rt') as input_file:
            reader = csv.DictReader(input_file)

            return [
                row for row in reader
            ]
            # for row in reader:
            #     print(row)

    def read_input(self):
        try:
            reader = self.readers[self.file_format]
        except KeyError:
            print('Input file format "{}" not supported. Supported formats are: {}'.format(
                self.file_format, self.readers.keys()))
            return False

        input_data = reader(self, self.input_path)
        available_properties = []
        for prop in input_data:
            available_properties.append(PropStats(**prop))

        return available_properties

    def process(self):

        available_properties = self.read_input()
        if not available_properties:
            print('No input to process!')
            return

        print('Processing {} properties'.format(len(available_properties)))

        for prop in available_properties:
            # TODO - read data from input
            prop.process(
                min_num_years=1,
                max_num_year=30,
                min_downpayment_percent=20,
                max_downpayment_percent=100,
            )
            if prop.local_max_roi > self.global_max_roi:
                self.global_max_roi = prop.local_max_roi
                self.global_max_stats = prop.local_max_stats

        print("\n\n\n==============\n************\nFINAL RESULT:\n************\n==============\n")
        pprint(self.global_max_stats)

        return {
            'properties': available_properties,
            'global_max_stats': self.global_max_stats,
            'global_max_roi': self.global_max_roi,
        }


    readers = {
        'csv': read_csv_input,
    }


class XLSXWriter(object):

    def __init__(self, data, output_file_path='output.xlsx'):
        self.data = data
        self.output_file_path = output_file_path

    def write(self):
        self.workbook = xlsxwriter.Workbook(self.output_file_path)

        self.formats = {
            'bold': self.workbook.add_format({'bold': True, 'align': 'center'}),
            'center': self.create_center_format(),
            'center_wrapped': self.create_center_wrapped_format(),
            'vcenter_wrapped': self.create_vcenter_wrapped_format(),
            'header_blue': self.create_header_blue_format(),
            'header_lime': self.create_header_lime_format(),
        }

        try:
            self.write_summary_sheet()

            for prop in self.data['properties']:
                self.write_prop_worksheet(prop)

        except Exception as e:
            raise
        else:
            print('Wrote results to {}'.format(self.output_file_path))
        finally:
            self.workbook.close()

    def create_center_wrapped_format(self):
        center_wrapped = self.workbook.add_format()
        self.centralize_format(center_wrapped)
        center_wrapped.set_align('vjustify')
        return center_wrapped

    def create_vcenter_wrapped_format(self):
        vcenter_wrapped = self.workbook.add_format()
        vcenter_wrapped.set_align('vcenter')
        vcenter_wrapped.set_align('vjustify')
        return vcenter_wrapped

    def create_header_blue_format(self):
        header_blue = self.workbook.add_format({
            'bold': True,
            'font_color': 'white',
            'bg_color': 'blue',
        })
        self.centralize_format(header_blue)
        header_blue.set_align('vjustify')
        header_blue.set_border_color('white')
        header_blue.set_border(1)
        return header_blue

    def create_header_lime_format(self):
        header_yellow = self.workbook.add_format({
            'bold': True,
            'bg_color': 'lime',
        })
        self.centralize_format(header_yellow)
        header_yellow.set_align('vjustify')
        header_yellow.set_border(1)
        return header_yellow

    def create_center_format(self):
        center = self.workbook.add_format()
        self.centralize_format(center)
        return center

    def create_rotated_format(self, angel=30):
        rotated = self.workbook.add_format()
        rotated.set_rotation(angel)
        return rotated

    def centralize_format(self, format_to_centralize):
        format_to_centralize.set_align('center')
        format_to_centralize.set_align('vcenter')

    def write_summary_sheet(self):
        summary_worksheet = self.workbook.add_worksheet('SUMMARY')
        summary_worksheet.merge_range('B2:E2', 'SUMMARY', self.formats['center'])

    def write_prop_worksheet(self, prop):
        worksheet_title = prop.name
        if len(worksheet_title) > 31:
            worksheet_title = worksheet_title[:31]

        worksheet = self.workbook.add_worksheet(worksheet_title)

        self.write_prop_basics_header(prop, worksheet)
        self.write_prop_basics_data(prop, worksheet)
        self.write_prop_winner_roi(prop, worksheet)
        self.write_prop_winner_roi_x_years(prop, worksheet)

        self.write_permutations(prop, worksheet)

    def write_permutations(self, prop, worksheet):

        header_row = row = 50
        header_col = col = 3

        rotated_format = self.create_rotated_format()
        worksheet.merge_range('C28:D29', 'Year vs.\nDownpayment', rotated_format)

        header_lime_format = self.formats['header_lime']
        for num_years_stats in prop.permutation_stats.items():
            row += 1
            col = header_col

            worksheet.write_number(row, header_col, num_years_stats[0], header_lime_format)

            for perm in num_years_stats[1].items():
                col += 1
                worksheet.write_number(header_row, col, perm[0], header_lime_format)
                self.write_permutation(perm[1], worksheet, row, col)

    def write_prop_basics_header(self, prop, worksheet):

        worksheet.merge_range('B2:G2', prop.name, self.formats['center_wrapped'])

        header_blue_format = self.formats['header_blue']

        worksheet.write('B3', 'Price', header_blue_format)
        worksheet.write('C3', 'Monthly Rent', header_blue_format)
        worksheet.write('D3', 'Area', header_blue_format)
        worksheet.write('E3', 'Extra Onetime Expense', header_blue_format)
        worksheet.write('F3', 'Url', header_blue_format)
        worksheet.write('G3', 'Annual Rent', header_blue_format)

    def write_prop_basics_data(self, prop, worksheet):
        center_wrapped = self.formats['center_wrapped']

        worksheet.write_number('B4', prop.price, center_wrapped)
        worksheet.write_number('C4', prop.rent, center_wrapped)
        worksheet.write_number('D4', prop.area, center_wrapped)
        worksheet.write_number('E4', prop.extra_onetime_expense, center_wrapped)
        worksheet.write_url('F4', prop.url, self.formats['center'])
        worksheet.write_number('G4', prop.annual_rent, center_wrapped)

    def write_prop_winner_roi(self, prop, worksheet):
        title_format = self.create_vcenter_wrapped_format()
        title_format.set_bg_color('silver')
        # title_format.set_font_color(colors['font'])
        worksheet.merge_range(
            'B7:E7',
            'WINNER Immediate ROI   -->   Annual ROI: {:.2f}%'.format(
                prop.local_max_roi * 100.00,
            ),
            title_format,
        )

        worksheet.merge_range('B8:E25', str(prop.local_max_stats), self.formats['vcenter_wrapped'])

    def write_prop_winner_roi_x_years(self, prop, worksheet):
        title_format = self.create_vcenter_wrapped_format()
        title_format.set_bg_color('silver')
        # title_format.set_font_color(colors['font'])
        worksheet.merge_range(
            'B28:E28',
            'WINNER X-Years ROI   -->   Annual ROI: {:.2f}%'.format(
                prop.local_max_stats_x_years.x_years_avg_annual_roi * 100.00,
            ),
            title_format,
        )

        worksheet.merge_range('B29:E46', str(prop.local_max_stats_x_years), self.formats['vcenter_wrapped'])


    def write_permutation(self, permutation, worksheet, cell_row, cell_col):

        worksheet.set_column(cell_row, cell_col, 55)
        cell_format = self.create_vcenter_wrapped_format()

        try:
            content = '----> Immediate Annual ROI: {:.2f}%  X-Years ROI: {:.2f}% <----\n{}'.format(
                permutation.annual_ROI * 100.00,
                permutation.x_years_avg_annual_roi * 100.00,
                str(permutation)
            )
        except AttributeError:
            content = str(permutation)
        else:
            color = self.get_roi_color(permutation)
            cell_format.set_bg_color(color)
            cell_format.set_font_color('white')
            # cell_format.set_border_color(colors['font'])
        
        cell_format.set_border(1)
        worksheet.write(cell_row, cell_col, content, cell_format)

    def get_roi_color(self, permutation):
        # 10 years bg_color
        # range -15% to 15% for red and green ??
        # immediate border color

        range_roi = 12
        range_rgb = 255.00

        immediate_ROI = min(range_roi, max(-range_roi, permutation.annual_ROI * 100.00))
        x_years_ROI = min(range_roi, max(-range_roi, permutation.x_years_avg_annual_roi * 100.00))

        # import ipdb; ipdb.set_trace()

        red = min(max(int(((range_roi - immediate_ROI)/range_roi) * range_rgb), 0), 255)
        green = min(max(int((x_years_ROI/range_roi) * range_rgb), 0), 255)

        return '#%02x%02x%02x' % (red, green, 0)

        # if roi >= 0.025:
        #     return {
        #         'bg': 'green',
        #         'font': 'white',
        #     }
        # elif roi > 0.00:
        #     return {
        #         'bg': 'yellow',
        #         'font': 'black',
        #     }
        # else:
        #     return {
        #         'bg': 'red',
        #         'font': 'white',
        #     }


if __name__ == '__main__':

    # TODO - read user input

    calculator = LandlordPropertyCalculator('/Users/malhyari/Downloads/PropertyOptions - Sheet1.csv', 'csv')
    results = calculator.process()

    writer = XLSXWriter(results)
    writer.write()


