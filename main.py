import sys
import mariadb
import xlsxwriter
import pandas as pd


def open_excel():
    global df
    tableName = input('Введите название таблицы xlsx:\n')
    df = pd.read_excel(tableName, dtype=str)


def save_excel():
    global df
    if not df.empty:
        tableName = input('Введите название таблицы xlsx:\n')
        writer = pd.ExcelWriter(tableName, engine='xlsxwriter')
        df.style.set_properties(**{'text-align': 'center'}).to_excel(writer, index=False, sheet_name='Таблица 1')
        worksheet = writer.sheets['Таблица 1']
        worksheet.set_column(0, len(df.columns), 10)
        writer.save()
    else:
        print('Таблица пуста')


def add_info():
    global df, cur

    try:
        y, m = map(int, input('Введите дату вида: yyyy-mm\n').split('-'))
        pp = str(pd.Period(str(y) + '-' + str(m), 'M'))
        if not df.empty:
            df.index = df.index * 2
            df.loc[3] = [pp, 0, 0] + [''] * (len(df.columns) - extraCol)
            df = df.sort_index()
            df.reset_index(inplace=True, drop=True)
            ops = [0] * (len(df.columns) - extraCol + 1)
        else:
            df['Date'] = [' ', ' ', pp]
            df[' '] = [' ', ' ', 0]
            df['  '] = [' ', ' ', 0]
            ops = [0]

        cur.execute(
            f"SELECT username, op_count, created FROM user WHERE MONTH(date) = {m} and YEAR(DATE) = {y}",
        )

        for (username, op_count, created) in cur:
            if (y - created.year) * 12 + m - created.month < len(ops):
                ops[(y - created.year) * 12 + m - created.month] += op_count

        # fulSum = 0
        #
        # for (username, op_count, created) in cur:
        #     fulSum += op_count
        #     dayCmp = days_between(youDate[1:11], created)
        #     if dayCmp < len(ops):
        #         ops[dayCmp] += op_count

        percent = round(ops[0] / sum(ops) * 100) if sum(ops) != 0 else 0
        df.insert(3, pp, [sum(ops), str(percent) + '%'] + ops)

        maxi = 0
        for i in range(2, len(df.columns) - 1):
            df[' '][i] += df.iloc[:, extraCol][i]
            maxi = max(df[' '][i], maxi)
        if maxi == 0:
            df['  '] = [' ', ' '] + [str(0) + '%'] * (len(df['  ']) - 2)
        else:
            for i in range(2, len(df.columns) - 1):
                df['  '][i] = str(round(df[' '][i] / maxi * 100, 2)) + '%'
    except ValueError:
        print('Неверно введены данные\n')


try:
    conn = mariadb.connect(
        user="blockchain",
        password="98kangaro\!ep",
        host="localhost",
        database="blockchain"
    )
except mariadb.Error as e:
    print(f"Error connecting to MariaDB Platform: {e}")
    sys.exit()


cur = conn.cursor()
df = pd.DataFrame()
extraCol = 3

while True:
    funcName = input('Введите название команды:\n')
    if funcName == 'add':
        add_info()
    elif funcName == 'save':
        save_excel()
    elif funcName == 'open':
        open_excel()
    elif funcName == 'exit':
        sys.exit()
    elif funcName == 'print':
        print(df)
    else:
        print('Wrong command\n')
