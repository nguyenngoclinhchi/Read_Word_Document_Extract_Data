import calendar
import pandas as pd


def get_last_date_month(year, month):
    print(calendar.monthrange(year, month))
    return calendar.monthrange(year, month)[1]


def is_last_month(index):
    return index.day == get_last_date_month(index.year, index.month)


def main():
    market_cap = pd.read_csv('linhchi.csv', parse_dates=['Date'], index_col='Date')
    last_month_date_list = list()
    for i in pd.date_range('2010-01', '2019-06', freq='M').strftime("%Y-%m").tolist():
        date = pd.to_datetime(market_cap.loc[i, :].index.tolist()[-1])
        converted_date = str(date.year) + "-" + str(date.month) + "-" + str(date.day)
        last_month_date_list.append(converted_date)
    print(last_month_date_list)
    market_cap['is_chosen'] = 'no'
    for i in last_month_date_list:
        market_cap.loc[i, 'is_chosen'] = 'yes'
    result = market_cap[market_cap['is_chosen'] == 'yes']
    result.to_excel

if __name__ == '__main__':
    main()
