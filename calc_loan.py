
import csv

import xlsxwriter

from constants import *
from property import Property


class LandlordPropertyCalculator(object):

    def __init__(self, input_path, file_format="csv"):
        self.input_path = input_path
        self.file_format = file_format

        self.global_max_roi = MINIMUM_DECIMAL
        self.global_max_stats = None


    def read_csv_input(self, input_path):
        with open(input_path, 'rt') as input_file:
            reader = csv.DictReader(input_file)

            return [
                row for row in reader
            ]

    def read_input(self):
        try:
            reader = self.readers[self.file_format]
        except KeyError:
            print('Input file format "{}" not supported. Supported formats are: {}'.format(
                self.file_format, self.readers.keys()))
            return False

        input_data = reader(self, self.input_path)
        available_properties = [
            Property(**prop) for prop in input_data
        ]

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
            if prop.max_roi > self.global_max_roi:
                self.global_max_roi = prop.max_roi
                self.global_max_stats = prop.max_stats

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

        worksheet.write('D3', 'Reletting Factor', header_blue_format)
        worksheet.write('E3', 'Gov Tax Discount', header_blue_format)
        worksheet.write('F3', 'Annual Gov Tax', header_blue_format)
        worksheet.write('G3', 'Monthly Total Expenses', header_blue_format)
        worksheet.write('H3', 'afterloan_monthly_income', header_blue_format)

        worksheet.write('I3', 'Area', header_blue_format)
        worksheet.write('J3', 'Extra Onetime Expense', header_blue_format)
        worksheet.write('K3', 'Url', header_blue_format)
        worksheet.write('L3', 'Annual Rent', header_blue_format)

    def write_prop_basics_data(self, prop, worksheet):
        center_wrapped = self.formats['center_wrapped']

        worksheet.write_number('B4', prop.price, center_wrapped)
        worksheet.write_number('C4', prop.rent, center_wrapped)

        worksheet.write_number('D4', prop.reletting_factor, center_wrapped)
        worksheet.write_number('E4', prop.gov_tax_discount, center_wrapped)
        worksheet.write_number('F4', prop.expense_annual_govt_tax, center_wrapped)
        worksheet.write_number('G4', prop.expense_total_monthly, center_wrapped)
        worksheet.write_number('H4', prop.afterloan_monthly_income, center_wrapped)

        worksheet.write_number('I4', prop.area, center_wrapped)
        worksheet.write_number('J4', prop.extra_onetime_expense, center_wrapped)
        worksheet.write_url('K4', prop.url, self.formats['center'])
        worksheet.write_number('L4', prop.annual_rent, center_wrapped)

    def write_prop_winner_roi(self, prop, worksheet):
        title_format = self.create_vcenter_wrapped_format()
        title_format.set_bg_color('silver')
        # title_format.set_font_color(colors['font'])
        worksheet.merge_range(
            'B7:E7',
            'WINNER Immediate ROI   -->   Annual ROI: {:.2f}%'.format(
                prop.max_roi * 100.00,
            ),
            title_format,
        )

        worksheet.merge_range('B8:E25', str(prop.max_stats), self.formats['vcenter_wrapped'])

    def write_prop_winner_roi_x_years(self, prop, worksheet):
        title_format = self.create_vcenter_wrapped_format()
        title_format.set_bg_color('silver')
        # title_format.set_font_color(colors['font'])
        worksheet.merge_range(
            'B28:E28',
            'WINNER X-Years ROI   -->   Annual ROI: {:.2f}%'.format(
                prop.max_stats_x_years.x_years_avg_annual_roi * 100.00,
            ),
            title_format,
        )

        worksheet.merge_range('B29:E46', str(prop.max_stats_x_years), self.formats['vcenter_wrapped'])


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
        # TODO - coloring needs revesiting

        # 10 years bg_color
        # range -15% to 15% for red and green ??
        # immediate border color

        range_roi = 12
        range_rgb = 255.00

        immediate_ROI = min(range_roi, max(-range_roi, permutation.annual_ROI * 100.00))
        x_years_ROI = min(range_roi, max(-range_roi, permutation.x_years_avg_annual_roi * 100.00))

        red = min(max(int(((range_roi - immediate_ROI)/range_roi) * range_rgb), 0), 255)
        green = min(max(int((x_years_ROI/range_roi) * range_rgb), 0), 255)

        return '#%02x%02x%02x' % (red, green, 0)


if __name__ == '__main__':

    # TODO - read user input

    calculator = LandlordPropertyCalculator('/Users/malhyari/Downloads/PropertyOptions - Sheet1.csv', 'csv')
    results = calculator.process()

    writer = XLSXWriter(results)
    writer.write()


