import pandas as pd
import numpy as np
import xlsxwriter
from pandas.tseries.offsets import MonthEnd
import datetime
from datetime import datetime, timedelta
import calendar


def monthdelta(d1, d2):
    """
    Returns the difference of two datetimes in months
    :param datetime d1: the earlier date
    :param datetime d2: the later date
    :return: the month difference
    :rtype: int
    """
    delta = 0
    while True:
        mdays = calendar.monthrange(d1.year, d1.month)[1]
        d1 += timedelta(days=mdays)
        if d1 <= d2:
            delta += 1
        else:
            break
    return delta


class Transaksi:
    """
    Transaksi holds the information of a transaction which conveys its amount and its transaction type
    """
    def __init__(self, jumlah, tipe):
        """
        Initializes the Transaksi object
        :param float jumlah: the amount of transaction
        :param str tipe: the type of transaction
        """
        self.jumlah = jumlah
        self.tipe = tipe

    def __repr__(self):
        """
        To String method
        :return: the string representation of Transaksi
        """
        return "({} sejumlah Rp{})".format(self.tipe, self.jumlah)


class EmployeeBalance:
    """
    Keeps track of the employee's balance throughout the date
    """
    INTEREST = 0.012  # the interest rate applied to daily average balance at the end of each month

    def __init__(self):
        """
        Initializes EmployeeBalance
        """
        self.balance = 0
        self.currDate = None

        self.__startDate = None
        self.__dailyBalance = 0
        self.__balanceAccumulation = 0

    def update(self, date, transaction, amount):
        """
        Updates the current balance given the transaction type, amount, and date
        :param datetime date: date of the transaction
        :param str transaction: the type of transaction
        :param float amount: the amount of transaction
        :return: the current employee balance
        :rtype: float
        """
        if self.currDate is None or self.__startDate is None:
            self.__startDate = pd.to_datetime(date, format="%Y%m%d")
            self.currDate = pd.to_datetime(date, format="%Y%m%d")
        if date.month != self.currDate.month:  # edge case if a month is skipped
            monthly_interest = self.update_monthly(date)
            print("monthly interest added = Rp{}".format(monthly_interest))

        if date != self.currDate:
            time_range = date - self.currDate
            # print(time_range.days)
            # print(date, end = '')
            self.__balanceAccumulation += self.balance * time_range.days
            self.currDate = pd.to_datetime(date, format="%Y%m%d")
        if transaction == 'Tabungan':
            self.balance += amount
        elif transaction == 'Tarikan ':
            self.balance -= amount
        # print('balance: {}'.format(self.balance))
        return self.balance

    def update_monthly(self, date):
        """
        Gets called when there is a change in month
        :param datetime date: the current transaction's date
        :return: the monthly added interest
        :rtype: float
        """
        end_of_month = calendar.monthrange(self.currDate.year, self.currDate.month)[1]
        end_of_month = datetime(self.currDate.year, self.currDate.month, end_of_month)
        time_range = end_of_month - self.currDate
        self.__balanceAccumulation += self.balance * time_range.days
        time_difference = end_of_month.day
        if time_difference is not 0:
            monthly_interest = self.__balanceAccumulation / time_difference * self.INTEREST
        else:
            monthly_interest = 0
        month_difference = monthdelta(self.__startDate, date) + 1
        self.balance += monthly_interest * month_difference
        self.__balanceAccumulation = 0
        self.__startDate = pd.to_datetime(date, format="%Y%m%d") + MonthEnd(1)
        self.currDate = end_of_month
        return monthly_interest


def load_file(data_frame):
    """
    Loads the excel file into cleaned up data structure that groups transactions by NIK and transaction date
    :param dataFrame data_frame: the data frame containing the excel document
    :return: the cleaned up data structure version of the excel file
    :rtype: Map<NIK, <Map<Tanggal, List<Transaksi(jumlah, tipe)>>>
    """

    nik_dict = dict()  # dictionary that contains all the transaction information
    for index, row, in data_frame.iterrows():
        if row['NIK'] not in nik_dict.keys():
            nik_dict[int(row['NIK'])] = {row['Tanggal']: []}
        elif row['Tanggal'] not in nik_dict[int(row['NIK'])].keys():
            nik_dict[int(row['NIK'])][row['Tanggal']] = []
        curr_transaksi = nik_dict[int(row['NIK'])][row['Tanggal']]
        jumlah = abs(row['Masuk'] + row['Keluar'])  # merge these columns to one
        curr_transaksi.append(Transaksi(jumlah, row['Transaksi']))
    print(nik_dict)

    return nik_dict


def generate_report(nik_dict):
    """
    Generates the report that also tracks employee balances for each employee
    :param Map<NIK, <Map<Tanggal, List<Transaksi(jumlah, tipe)>>> nik_dict: the data frame as a data structure
    """
    workbook = xlsxwriter.Workbook('nik_report.xlsx')
    date_format = workbook.add_format({'num_format': 'dd/mm/yy'})
    money_format = workbook.add_format({'num_format': 'Rp#,##0.00'})

    for nik in nik_dict.keys():
        print("\n\nNIK {}".format(nik))
        employee_balance = EmployeeBalance()
        curr_sheet = workbook.add_worksheet("NIK {}".format(nik))
        curr_sheet.set_column(1, 1, 12)
        curr_sheet.set_column(3, 3, 12)
        curr_sheet.write(0, 0, "Tanggal")
        curr_sheet.write(0, 1, "Jumlah Transaksi")
        curr_sheet.write(0, 2, "Tipe Transaksi")
        curr_sheet.write(0, 3, "Saldo")
        row = 0
        col = 0
        for date in nik_dict[nik].keys():
            item_list = nik_dict[nik][date]
            # print(item)
            for item in item_list:
                # Filters the types of transactions to be written
                if item.tipe == "Tabungan" or item.tipe == "Tarikan ":
                    row += 1
                    curr_sheet.write(row, col, date, date_format)
                    curr_sheet.write(row, col + 1, item.jumlah, money_format)
                    curr_sheet.write(row, col + 2, item.tipe)
                    balance = employee_balance.update(date, item.tipe, item.jumlah)
                    curr_sheet.write(row, col + 3, balance, money_format)

    workbook.close()


if __name__ == "__main__":
    df = pd.read_excel('toBeParsed.xlsx', sheet_name='Kas')
    df = df[np.isfinite(df['NIK'])]  # dropping nonexistent NIKs
    df['Masuk'].fillna(0, inplace=True)  # replace all NaN with 0
    df['Keluar'].fillna(0, inplace=True)  # replace all NaN with 0

    nik_dict = load_file(df)
    # print(nik_dict)

    generate_report(nik_dict)
