import openpyxl


def read_xlsx(path):
    accounts = {}

    workbook = openpyxl.load_workbook(path)
    sheet = workbook.active

    row_len = sheet.max_row
    col_len = sheet.max_column

    for row_i in range(2, row_len + 1):
        post = {}
        user = sheet.cell(row = row_i, column = 4).value

        if user not in accounts:
            accounts[user] = []

        for col_j in range(1, col_len + 1):
            key = sheet.cell(row = 1, column = col_j).value
            val = sheet.cell(row = row_i, column = col_j).value
            post[key] = val
        accounts[user].append(post)

    return accounts

def find_bots(accounts):

    scrambled_names = []
    for user in accounts:
        if is_name_alphanumeric_scramble(user[""]):


    bots = []
    return bots

def is_name_alphanumeric_scramble(name):

def frequent_hashtag(text):
    return []

def main():
    path = "/home/ybrenin/PycharmProjects/IGS_502_Data_Parser/data/defundthepolice-full.xlsx"
    account_list = read_xlsx(path)
    test = account_list['Natalia_Army_of_1@mstdn.party']

    print(len(test))

main()
