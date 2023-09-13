# version 1.1.0
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
from openpyxl.utils.cell import column_index_from_string
import pandas as pd
# from dateTime import yesterday, GIPdate, current_month, recons_yesterday
import os
import time
import calendar
from datetime import datetime
from datetime import date
from datetime import timedelta

prompt = input("Recons for yesterday? (Y/N) :")
if prompt.upper() == 'Y':
    today = datetime.today().strftime('%d_%b_%y')

    # Yesterday's date
    yesterday = (datetime.now() - timedelta(1)).strftime('_%d %b_%y')
    print(yesterday)

    # Current month
    current_month = calendar.month_abbr[date.today().month].upper()
    print(current_month)

    recons_yesterday = str((datetime.now() - timedelta(1)).strftime('%Y-%m-%d') + ' 00:00:00')
    print(recons_yesterday)

    if date.today().month < 10:
        GIPmonth = "0" + str(date.today().month)
    else:
        GIPmonth = str(date.today().month)

    if date.today().day < 10:
        GIPday = "0" + str(date.today().day + 1)
    else:
        GIPday = str(date.today().day + 1)

    GIPdate = str(str(date.today().year) + str(GIPmonth) + str((GIPday)))
    print(GIPdate)
else:
    print('Please enter the date...')
    day = input("Recons Day: ")
    GIPday = int(day) + 1
    if int(day) < 10:
        day = "0" + day
    month = input("Recons Month: ")
    month_name = calendar.month_abbr[int(month)]
    if int(month) < 10:
        month = "0" + month
    year = input("Recons Year: ")

    yesterday = "_" + day + " " + month_name + "_" + year
    print(yesterday)

    recons_yesterday = "20" + year + "-" + month + "-" + day + ' 00:00:00'
    print(recons_yesterday)

    current_month = calendar.month_abbr[int(month)].upper()
    print(current_month)

    if int(month) < 10:
        GIPmonth = month

    if int(GIPday) < 10:
        GIPday = "0" + str(GIPday)
    GIPdate = str("20" + str(year) + str(GIPmonth) + str((GIPday)))
    print(GIPdate)


def get_column(keyword):
    for i in range(first_row, last_row + 1):
        for j in range(first_col, last_col + 1):
            if sheet[str(get_column_letter(j)) + str(i)].value == keyword:
                return (get_column_letter(j))


dir_list = os.listdir()
# Check if metabase file is available
metaFileFound = False
for file in dir_list:
    if file == 'Metabase' + yesterday + '.xlsx':
        metaFileFound = True
        break
    else:
        metaFileFound = False
# ############################################# SLYDEPAY01 OVA ########################################################
#
#
fileFound = False
matchOVAFound = False
for file in dir_list:
    if file == 'MIGS 01' + yesterday + '.xlsx':
        fileFound = True
        matchOVAFound = True
        wb = load_workbook('MIGS 01' + yesterday + '.xlsx', read_only=False)
        sheet = wb[wb.sheetnames[0]]
        wb.active = wb[wb.sheetnames[0]]

        # DEFINE MAX AND MIN COLUMNS AND ROWS
        first_col = wb.active.min_column
        last_col = wb.active.max_column
        first_row = wb.active.min_row
        last_row = wb.active.max_row

        MIGS_01_OVA_Volume = (last_row + 1) - (first_row + 4)
        OVA_VOLUME1 = MIGS_01_OVA_Volume
        print('MIGS_01_OVA_Volume: ' + str(MIGS_01_OVA_Volume))

        MIGS_01_OVA_Sum = 0.00

        amountColumn = get_column('Amount')
        for i in range(first_row + 4, last_row + 1):
            MIGS_01_OVA_Sum = MIGS_01_OVA_Sum + abs(float(sheet[amountColumn + str(i)].value))

        print('MIGS_01_OVA_Sum: ' + str(MIGS_01_OVA_Sum))
        OVA_VALUE1 = MIGS_01_OVA_Sum
        ova_id_col = get_column('Merchant Transaction Reference')
        tmp_id_col = ova_id_col

        alternate_id_col = get_column("Transaction ID")
        wb.close()
        wb.save('MIGS 01' + yesterday + '- Recons.xlsx')
        match_OVA = 'MIGS 01' + yesterday + '- Recons.xlsx'

        wb = load_workbook(match_OVA)
        ws1 = wb.create_sheet("Duplicates")
        ws1.title = 'Duplicates'

        ws1 = wb.create_sheet("Missing Integrator Transactions")
        ws1.title = 'Missing Integrator Transactions'

        ws1 = wb.create_sheet("Missing Integrator Transactions")
        ws1.title = 'Missing Integrator Transactions'

        wb.close()
        wb.save(match_OVA)
        break
    else:
        fileFound = False

if fileFound is False:
    OVA_VOLUME1 = 0
    OVA_VALUE1 = 0

#
#
#
fileFound = False
matchINTFound = False
#  ##################################################### SLYDEPAY01 INT #################################################
for file in dir_list:
    if file == 'MIGS 01 Metabase' + yesterday + '.xlsx':
        fileFound = True
        matchINTFound = True
        wb = load_workbook('MIGS 01 Metabase' + yesterday + '.xlsx')
        sheet = wb[wb.sheetnames[0]]
        wb.active = wb[wb.sheetnames[0]]
        file = pd.read_excel('MIGS 01 Metabase' + yesterday + '.xlsx')
        match_INT = 'MIGS 01 Metabase' + yesterday + '.xlsx'

        # DEFINE MAX AND MIN COLUMNS AND ROWS
        first_col = wb.active.min_column
        last_col = wb.active.max_column
        first_row = wb.active.min_row
        last_row = wb.active.max_row

        MIGS_01_INT_Volume = (last_row + 1) - (first_row + 1)
        INT_VOLUME1 = MIGS_01_INT_Volume
        print('MIGS_01_INT_Volume: ' + str(MIGS_01_INT_Volume))

        MIGS_01_INT_Sum = 0.00

        int_id_col = get_column('External Payment Request â†’ Institution Trans ID')
        id_Col = 'External Payment Request â†’ Institution Trans ID'
        if int_id_col is None:
            int_id_col = get_column("Institution Trans ID")
            id_Col = "Institution Trans ID"
        amountColumn = get_column('Amount')
        if amountColumn is None:
            amountColumn = get_column("Amount ($)")
        if amountColumn is None:
            amountColumn = get_column("Actual Amount ($)")
        for i in range(first_row + 1, last_row + 1):
            MIGS_01_INT_Sum = MIGS_01_INT_Sum + abs(float(sheet[amountColumn + str(i)].value))

        print('MIGS_01_INT_Sum: ' + str(MIGS_01_INT_Sum))
        INT_VALUE1 = MIGS_01_INT_Sum
        wb.close()
        wb.save('MIGS 01 Metabase' + yesterday + '.xlsx')
        # ---------------------------------------- Get duplicates -----------------------------------------------------
        wb = load_workbook('MIGS 01 Metabase' + yesterday + '.xlsx')
        Owb = load_workbook(match_OVA)
        sheet = wb[wb.sheetnames[0]]
        wb.active = wb[wb.sheetnames[0]]
        list = []
        duplicates = []
        count = 1
        counter = 0
        headerChecker = True
        duplicates_value_sum = 0.00
        for id in range(first_row + 1, last_row + 1):
            if sheet[int_id_col + str(id)].value not in list:
                list.append(sheet[int_id_col + str(id)].value)
            elif sheet[int_id_col + str(id)].value not in duplicates:
                    data = file[file[id_Col] == sheet[int_id_col + str(id)].value]
                    try:
                        data_set_sum = data.iloc[1:, data.columns.get_loc("Amount ($)")].sum()
                    except:
                        data_set_sum = data.iloc[1:, data.columns.get_loc("Actual Amount ($)")].sum()
                    duplicates_value_sum = duplicates_value_sum + data_set_sum
                    duplicates.append(sheet[int_id_col + str(id)].value)
                    if counter > 0:
                        headerChecker = False
                    with pd.ExcelWriter(match_OVA, mode='a', engine="openpyxl", if_sheet_exists='overlay') as writer:
                        data.to_excel(writer, sheet_name='Duplicates', startrow=count, header=headerChecker)
                        if headerChecker is False:
                            count += 2
                        else:
                            count += (data.shape[0]+1)
                        counter += 1
            else:
                if sheet[int_id_col + str(id)].value in duplicates:
                    counter += 1
                continue
        wb.close()
        Owb.close()
        Owb.save(match_OVA)
        wb.save('MIGS 01 Metabase' + yesterday + '.xlsx')

        wb = load_workbook(match_OVA)
        wb.active = wb['Duplicates']
        dup_sheet = wb['Duplicates']
        sheet =dup_sheet

        first_col = wb.active.min_column
        last_col = wb.active.max_column
        first_row = wb.active.min_row
        last_row = wb.active.max_row

        if counter > 0:
            amountColumn = get_column('Amount')
            if amountColumn is None:
                amountColumn = get_column("Amount ($)")
            if amountColumn is None:
                amountColumn = get_column("Actual Amount ($)")

            print(f"Number of duplicates: {counter}")
            dup_sheet[amountColumn + str(last_row + 4)].value = str(round(duplicates_value_sum, 2))
            dup_sheet[amountColumn + str(last_row + 5)].value = str(counter)

            print(f"Duplicates: {duplicates}")

            DUPLICATES_VOLUME1 = counter
            DUPLICATES_VALUE1 = duplicates_value_sum
        else:
            DUPLICATES_VOLUME1 = 0
            DUPLICATES_VALUE1 = 0

        wb.close()
        wb.save(match_OVA)

        # ---------------------------------------- MISSING INT TRANSACTIONS -----------------------------------------
        if matchOVAFound is True and matchINTFound is True:
            Owb = load_workbook(match_OVA)
            Osheet = Owb[Owb.sheetnames[0]]
            Owb.active = Owb[Owb.sheetnames[0]]
            file = pd.read_excel(match_OVA, header=[3])

            Iwb = load_workbook(match_INT)
            Isheet = Iwb[Iwb.sheetnames[0]]
            Iwb.active = Iwb[Iwb.sheetnames[0]]

            Ofirst_col = Owb.active.min_column
            Olast_col = Owb.active.max_column
            Ofirst_row = Owb.active.min_row
            Olast_row = Owb.active.max_row

            Ifirst_col = Iwb.active.min_column
            Ilast_col = Iwb.active.max_column
            Ifirst_row = Iwb.active.min_row
            Ilast_row = Iwb.active.max_row

            missing_INT_list = []
            missing_OVA_list = []
            headerChecker = True
            matchFound = False
            counter = 0

            for i in range(Ofirst_row + 4, Olast_row + 1):
                for j in range(Ifirst_row + 1, Ilast_row + 1):
                    if Osheet[ova_id_col + str(i)].value is None or Osheet[ova_id_col + str(i)].value == "#VALUE!":
                        ova_id_col = alternate_id_col
                        matchFound = False
                        break
                    if str(Osheet[ova_id_col + str(i)].value) == str(Isheet[int_id_col + str(j)].value):
                        matchFound = True
                        break
                    else:
                        matchFound = False
                if matchFound is False:
                    missing_INT_list.append(Osheet[ova_id_col + str(i)].value)
                    ova_id_col = tmp_id_col
            print(f"Missing integrator transactions: {missing_INT_list}")

            count = 0
            for transaction in missing_INT_list:
                data = file[file['Merchant Transaction Reference'] == transaction]
                if data.empty:
                    data = file[file['Transaction ID'] == transaction]
                if counter > 0:
                    headerChecker = False
                with pd.ExcelWriter(match_OVA, mode='a', engine="openpyxl", if_sheet_exists='overlay') as writer:
                    data.to_excel(writer, sheet_name='Missing Integrator Transactions', startrow=count,
                                  header=headerChecker)
                    if headerChecker is False:
                        count += 1
                    else:
                        count += 2
                    counter += 1

            Iwb.close()
            Owb.close()

            # ----------------------------------------- MISSING OVA TRANSACTIONS -------------------------------------------
            Owb = load_workbook(match_OVA)
            Osheet = Owb[Owb.sheetnames[0]]
            Owb.active = Owb[Owb.sheetnames[0]]

            Iwb = load_workbook(match_INT)
            Isheet = Iwb[Iwb.sheetnames[0]]
            Iwb.active = Iwb[Iwb.sheetnames[0]]
            file = pd.read_excel(match_INT)

            Ofirst_col = Owb.active.min_column
            Olast_col = Owb.active.max_column
            Ofirst_row = Owb.active.min_row
            Olast_row = Owb.active.max_row

            Ifirst_col = Iwb.active.min_column
            Ilast_col = Iwb.active.max_column
            Ifirst_row = Iwb.active.min_row
            Ilast_row = Iwb.active.max_row

            missing_INT_list = []
            missing_OVA_list = []
            matchFound = False
            headerChecker = True
            counter = 0

            for i in range(Ifirst_row + 1, Ilast_row + 1):
                for j in range(Ofirst_row + 1, Olast_row + 1):
                    if str(Isheet[int_id_col + str(i)].value) == str(Osheet[ova_id_col + str(j)].value):
                        matchFound = True
                        break
                    else:
                        matchFound = False
                if matchFound is False:
                    missing_OVA_list.append(Isheet[int_id_col + str(i)].value)
            print(f"Missing OVA transactions: {missing_OVA_list}")
            count = 0
            for transaction in missing_OVA_list:
                data = file[file[id_Col] == transaction]
                if counter > 0:
                    headerChecker = False
                with pd.ExcelWriter(match_OVA, mode='a', engine="openpyxl", if_sheet_exists='overlay') as writer:
                    data.to_excel(writer, sheet_name='Missing OVA Transactions', startrow=count, header=headerChecker)
                    if headerChecker is False:
                        count += 1
                    else:
                        count += 2
                    counter += 1
            Iwb.close()
            Owb.close()
        break
    else:
        fileFound = False

if fileFound is False:
    INT_VOLUME1 = 0
    INT_VALUE1 = 0
    DUPLICATES_VOLUME1 = 0
    DUPLICATES_VALUE1 = 0

#
#
#
#
matchOVAFound = False
matchINTFound = False
fileFound = False
#  ######################################## SLYDEPULL PROMPTS OVA #####################################################
for file in dir_list:
    if file == 'MTN Prompt' + yesterday + '.xlsx':
        fileFound = True
        matchOVAFound = True
        Prompt_wb = load_workbook('MTN Prompt' + yesterday + '.xlsx')
        sheet = Prompt_wb[Prompt_wb.sheetnames[0]]

        # DEFINE MAX AND MIN COLUMNS AND ROWS
        first_col = Prompt_wb.active.min_column
        last_col = Prompt_wb.active.max_column
        first_row = Prompt_wb.active.min_row
        last_row = Prompt_wb.active.max_row

        MTN_PROMPT_OVA_Volume = (last_row + 1) - (first_row + 1)
        OVA_VOLUME2 = MTN_PROMPT_OVA_Volume
        print('MTN_PROMPT_OVA_Volume: ' + str(MTN_PROMPT_OVA_Volume))

        MTN_PROMPT_OVA_Sum = 0.00

        amountColumn = get_column('Amount')
        ova_id_col = get_column('External Transaction Id')
        tmp_id_col = ova_id_col
        alternate_id_col = get_column("Id")
        for i in range(first_row + 1, last_row + 1):
            check = sheet[amountColumn + str(i)].value
            MTN_PROMPT_OVA_Sum = MTN_PROMPT_OVA_Sum + abs(float(sheet[amountColumn + str(i)].value))

        print('MTN_PROMPT_OVA_Sum: ' + str(MTN_PROMPT_OVA_Sum))
        OVA_VALUE2 = MTN_PROMPT_OVA_Sum
        Prompt_wb.close()
        Prompt_wb.save('MTN Prompt' + yesterday + '- Recons.xlsx')
        match_OVA = 'MTN Prompt' + yesterday + '- Recons.xlsx'

        wb = load_workbook(match_OVA)
        ws1 = wb.create_sheet("Duplicates")
        ws1.title = 'Duplicates'

        ws1 = wb.create_sheet("Missing Integrator Transactions")
        ws1.title = 'Missing Integrator Transactions'

        ws1 = wb.create_sheet("Missing Integrator Transactions")
        ws1.title = 'Missing Integrator Transactions'

        wb.close()
        wb.save(match_OVA)
        break
    else:
        fileFound = False

if fileFound is False:
    OVA_VOLUME2 = 0
    OVA_VALUE2 = 0
#
#
#

# ################################### MTN(SLYDEPULL) PROMPT INT ###############################################
if metaFileFound is True:
    matchINTFound = True
    wb = load_workbook('Metabase' + yesterday + '.xlsx')
    file = pd.read_excel('Metabase' + yesterday + '.xlsx', sheet_name='Query result')

    sheet = wb['Query result']
    wb.active = wb['Query result']
    first_col = wb.active.min_column
    last_col = wb.active.max_column
    first_row = wb.active.min_row
    last_row = wb.active.max_row

    int_id_col = get_column('integratorTransId')
    transIdHeader = "integratorTransId"
    if int_id_col is None:
        int_id_col = get_column("IntegratorTransId")
        transIdHeader = "IntegratorTransId"
    debitCreditFlagColumn = get_column('creditDebitFlag')
    creditDebitHeader = "creditDebitFlag"
    if debitCreditFlagColumn is None:
        debitCreditFlagColumn = get_column("CreditDebitFlag")
        creditDebitHeader = "CreditDebitFlag"
    serviceName_col = get_column('serviceName')
    serviceNameHeader = "serviceName"
    if serviceName_col is None:
        serviceName_col = get_column("ServiceName")
        serviceNameHeader = "ServiceName"
    amountColumn = get_column("amount")
    if amountColumn is None:
        amountColumn = get_column("Amount")

    INT_VOLUME2 = 0.00
    INT_VALUE2 = 0.00
    row_count = 1
    count = 0
    sum = 0.00
    counter = 0
    headerChecker = True
    # ------------------------------------------------- Duplicates ---------------------------------------------------------
    list = []
    duplicates = []
    duplicates_value_sum= 0.00
    for i in range(first_row + 1, last_row + 1):
        if sheet[serviceName_col + str(i)].value == 'MTN OVA' and sheet[debitCreditFlagColumn + str(i)].value == 'C':
            count += 1
            sum = sum + abs(float(sheet[amountColumn + str(i)].value))
            if sheet[int_id_col + str(i)].value not in list:
                list.append(sheet[int_id_col + str(i)].value)
            elif sheet[int_id_col + str(i)].value not in duplicates:
                    data = file[(file[transIdHeader] == sheet[int_id_col + str(i)].value) & (file[serviceNameHeader] == 'MTN OVA') & (file[creditDebitHeader] == 'C')]
                    try:
                        data_set_sum = data.iloc[1:, data.columns.get_loc("Amount")].sum()
                    except:
                        data_set_sum = data.iloc[1:, data.columns.get_loc("amount")].sum()
                    duplicates_value_sum = duplicates_value_sum + data_set_sum
                    duplicates.append(sheet[int_id_col + str(i)].value)
                    if counter > 0:
                        headerChecker = False
                    with pd.ExcelWriter(match_OVA, mode='a', engine="openpyxl", if_sheet_exists='overlay') as writer:
                        data.to_excel(writer, sheet_name='Duplicates', startrow=row_count, header=headerChecker)
                        if headerChecker is False:
                            row_count += 2
                        else:
                            row_count += (data.shape[0]+1)
                        counter += 1
            else:
                if sheet[int_id_col + str(i)].value in duplicates:
                    counter += 1
                continue
        else:
            continue
    INT_VOLUME2 = count
    INT_VALUE2 = sum
    print(f"MTN Prompt INT Volume: {count}")
    print(f"MTN Prompt INT Value: {sum}")

    wb.close()
    Mwb = load_workbook(match_OVA)

    wb = load_workbook(match_OVA)
    wb.active = wb['Duplicates']
    dup_sheet = wb['Duplicates']
    sheet = dup_sheet

    first_col = wb.active.min_column
    last_col = wb.active.max_column
    first_row = wb.active.min_row
    last_row = wb.active.max_row

    if counter > 0:
        amountColumn = get_column('Amount')
        if amountColumn is None:
            amountColumn = get_column("amount")

        print(f"Number of duplicates: {counter}")
        dup_sheet[amountColumn + str(last_row + 4)].value = str(round(duplicates_value_sum, 2))
        dup_sheet[amountColumn + str(last_row + 5)].value = str(counter)

        print(f"Duplicates: {duplicates}")

        DUPLICATES_VOLUME2 = counter
        DUPLICATES_VALUE2 = duplicates_value_sum
    else:
        DUPLICATES_VOLUME2 = 0
        DUPLICATES_VALUE2 = 0

    wb.close()
    wb.save(match_OVA)

    if matchOVAFound is True and matchINTFound is True:
        # --------------------------------------- MISSING INT TRANSACTIONS ------------------------------------------------
        Owb = load_workbook(match_OVA)
        Osheet = Owb[Owb.sheetnames[0]]
        Owb.active = Owb[Owb.sheetnames[0]]
        file = pd.read_excel(match_OVA)
        sheet = Osheet

        Iwb = load_workbook('Metabase' + yesterday + '.xlsx')
        Isheet = Iwb['Query result']
        Iwb.active = Iwb['Query result']

        Ofirst_col = Owb.active.min_column
        Olast_col = Owb.active.max_column
        Ofirst_row = Owb.active.min_row
        Olast_row = Owb.active.max_row

        Ifirst_col = Iwb.active.min_column
        Ilast_col = Iwb.active.max_column
        Ifirst_row = Iwb.active.min_row
        Ilast_row = Iwb.active.max_row

        missing_INT_list = []
        missing_OVA_list = []
        matchFound = False
        headerChecker = True
        counter = 0

        for i in range(Ofirst_row + 1, Olast_row + 1):
            for j in range(Ifirst_row + 1, Ilast_row + 1):
                if Osheet[ova_id_col + str(i)].value is None or Osheet[ova_id_col + str(i)].value == "#VALUE!":
                    ova_id_col = alternate_id_col
                    matchFound = False
                    break
                if (Isheet[serviceName_col + str(j)].value == 'MTN OVA') and (
                        Isheet[debitCreditFlagColumn + str(j)].value == 'C'):
                    if str(Osheet[ova_id_col + str(i)].value).lstrip('=') in str(Isheet[int_id_col + str(j)].value):
                        matchFound = True
                        break
                    else:
                        matchFound = False
            if matchFound is False:
                missing_INT_list.append(Osheet[ova_id_col + str(i)].value)
                ova_id_col = tmp_id_col

        print(f"Missing integrator transactions: {missing_INT_list}")

        count = 0
        for transaction in missing_INT_list:
            data = (file[file['External Transaction Id'] == transaction])
            if data.empty:
                data = (file[file['Id'] == transaction])
            if counter > 0:
                headerChecker = False
            with pd.ExcelWriter(match_OVA, mode='a', engine="openpyxl", if_sheet_exists='overlay') as writer:
                data.to_excel(writer, sheet_name='Missing Integrator Transactions', startrow=count,
                              header=headerChecker)
                if headerChecker is False:
                    count += 1
                else:
                    count += 2
                counter += 1

        Iwb.close()
        Owb.close()
        Mwb.close()

        # -------------------------------------- MISSING OVA TRANSACTIONS --------------------------------------------------
        Owb = load_workbook(match_OVA)
        Osheet = Owb[Owb.sheetnames[0]]
        Owb.active = Owb[Owb.sheetnames[0]]
        sheet = Osheet
        Mwb = load_workbook(match_OVA)

        Iwb = load_workbook('Metabase' + yesterday + '.xlsx')
        file = pd.read_excel('Metabase' + yesterday + '.xlsx', sheet_name='Query result')
        Isheet = Iwb['Query result']
        Iwb.active = Iwb['Query result']

        Ofirst_col = Owb.active.min_column
        Olast_col = Owb.active.max_column
        Ofirst_row = Owb.active.min_row
        Olast_row = Owb.active.max_row

        Ifirst_col = Iwb.active.min_column
        Ilast_col = Iwb.active.max_column
        Ifirst_row = Iwb.active.min_row
        Ilast_row = Iwb.active.max_row

        missing_INT_list = []
        missing_OVA_list = []
        matchFound = False
        headerChecker = True
        counter = 0

        for i in range(Ifirst_row + 1, Ilast_row + 1):
            if (Isheet[serviceName_col + str(i)].value == 'MTN OVA') and (
                    Isheet[debitCreditFlagColumn + str(i)].value == 'C'):
                for j in range(Ofirst_row + 1, Olast_row + 1):
                    if str(Osheet[ova_id_col + str(j)].value).lstrip('=') in str(Isheet[int_id_col + str(i)].value):
                        matchFound = True
                        break
                    else:
                        matchFound = False
                if matchFound is False:
                    missing_OVA_list.append(Isheet[int_id_col + str(i)].value)
            else:
                continue

        print(f"Missing OVA transactions: {missing_OVA_list}")

        count = 0
        for transaction in missing_OVA_list:
            try:
                data = file[(file['integratorTransId'] == str(transaction))]
            except KeyError:
                data = file[(file['IntegratorTransId'] == str(transaction))]
            if counter > 0:
                headerChecker = False
            with pd.ExcelWriter(match_OVA, mode='a', engine="openpyxl", if_sheet_exists='overlay') as writer:
                data.to_excel(writer, sheet_name='Missing OVA Transactions', startrow=count, header=headerChecker)
                if headerChecker is False:
                    count += 1
                else:
                    count += 2
                counter += 1

        Iwb.close()
        Owb.close()
        Mwb.close()
else:
    INT_VOLUME2 = 0
    INT_VALUE2 = 0
    DUPLICATES_VOLUME2 = 0
    DUPLICATES_VALUE2 = 0

#
#
OVA_VOLUME3 = 0
OVA_VALUE3 = 0
INT_VOLUME3 = 0
INT_VALUE3 = 0
DUPLICATES_VALUE3 = 0
DUPLICATES_VOLUME3 = 0
#
#
#
matchOVAFound = False
matchINTFound = False
fileFound = False
#  ################################################ MTN PORTAL OVA #####################################################
for file in dir_list:
    if file == 'MTN Cashout' + yesterday + '.xlsx':
        fileFound = True
        matchOVAFound = True
        wb = load_workbook('MTN Cashout' + yesterday + '.xlsx')
        sheet = wb[wb.sheetnames[0]]
        wb.active = wb[wb.sheetnames[0]]

        # DEFINE MAX AND MIN COLUMNS AND ROWS
        first_col = wb.active.min_column
        last_col = wb.active.max_column
        first_row = wb.active.min_row
        last_row = wb.active.max_row

        MTN_APPROVALS_OVA_Volume = (last_row + 1) - (first_row + 1)
        OVA_VOLUME4 = MTN_APPROVALS_OVA_Volume
        print('MTN_PORTAL_OVA_Volume: ' + str(MTN_APPROVALS_OVA_Volume))

        MTN_APPROVALS_OVA_Sum = 0.00

        amountColumn = get_column('Amount')
        ova_id_col = get_column('External Transaction Id')
        tmp_id_col = ova_id_col
        alternate_id_col = get_column("Id")
        for i in range(first_row + 1, last_row + 1):
            MTN_APPROVALS_OVA_Sum = MTN_APPROVALS_OVA_Sum + abs(float(sheet[amountColumn + str(i)].value))

        print('MTN_PORTAL_OVA_Sum: ' + str(MTN_APPROVALS_OVA_Sum))
        OVA_VALUE4 = MTN_APPROVALS_OVA_Sum
        wb.close()
        wb.save('MTN Cashout' + yesterday + '- Recons.xlsx')
        match_OVA = 'MTN Cashout' + yesterday + '- Recons.xlsx'

        wb = load_workbook(match_OVA)
        ws1 = wb.create_sheet("Duplicates")
        ws1.title = 'Duplicates'

        ws1 = wb.create_sheet("Missing Integrator Transactions")
        ws1.title = 'Missing Integrator Transactions'

        ws1 = wb.create_sheet("Missing Integrator Transactions")
        ws1.title = 'Missing Integrator Transactions'

        wb.close()
        wb.save(match_OVA)
        break
    else:
        fileFound = False

if fileFound is False:
    OVA_VOLUME4 = 0
    OVA_VALUE4 = 0
#
#
#  ################################################ MTN PORTAL INT #####################################################
if metaFileFound is True:
    matchINTFound = True
    wb = load_workbook('Metabase' + yesterday + '.xlsx')
    file = pd.read_excel('Metabase' + yesterday + '.xlsx', sheet_name='Query result')

    sheet = wb['Query result']
    wb.active = wb['Query result']
    first_col = wb.active.min_column
    last_col = wb.active.max_column
    first_row = wb.active.min_row
    last_row = wb.active.max_row

    int_id_col = get_column('integratorTransId')
    transIdHeader = "integratorTransId"
    if int_id_col is None:
        int_id_col = get_column("IntegratorTransId")
        transIdHeader = "IntegratorTransId"
    debitCreditFlagColumn = get_column('creditDebitFlag')
    creditDebitHeader = "creditDebitFlag"
    if debitCreditFlagColumn is None:
        debitCreditFlagColumn = get_column("CreditDebitFlag")
        creditDebitHeader = "CreditDebitFlag"
    serviceName_col = get_column('serviceName')
    serviceNameHeader = "serviceName"
    if serviceName_col is None:
        serviceName_col = get_column("ServiceName")
        serviceNameHeader = "ServiceName"

    amountColumn = get_column("amount")
    if amountColumn is None:
        amountColumn = get_column("Amount")

    INT_VOLUME4 = 0.00
    INT_VALUE4 = 0.00
    row_count = 1
    count = 0
    sum = 0.00
    counter = 0
    headerChecker = True
    # ------------------------------------------------- Duplicates ---------------------------------------------------------
    list = []
    duplicates = []
    for i in range(first_row + 1, last_row + 1):
        if sheet[serviceName_col + str(i)].value == 'MTN OVA' and sheet[debitCreditFlagColumn + str(i)].value == 'D':
            count += 1
            sum = sum + abs(float(sheet[amountColumn + str(i)].value))
            if sheet[int_id_col + str(i)].value not in list:
                list.append(sheet[int_id_col + str(i)].value)
            elif sheet[int_id_col + str(i)].value not in duplicates:
                    data = file[(file[transIdHeader] == sheet[int_id_col + str(i)].value) & (file[serviceNameHeader] == 'MTN OVA') & (file[creditDebitHeader] == 'D')]
                    try:
                        data_set_sum = data.iloc[1:, data.columns.get_loc("Amount")].sum()
                    except:
                        data_set_sum = data.iloc[1:, data.columns.get_loc("amount")].sum()
                    duplicates_value_sum = duplicates_value_sum + data_set_sum
                    duplicates.append(sheet[int_id_col + str(i)].value)
                    if counter > 0:
                        headerChecker = False
                    with pd.ExcelWriter(match_OVA, mode='a', engine="openpyxl", if_sheet_exists='overlay') as writer:
                        data.to_excel(writer, sheet_name='Duplicates', startrow=row_count, header=headerChecker)
                        if headerChecker is False:
                            row_count += 2
                        else:
                            row_count += (data.shape[0]+1)
                        counter += 1
            else:
                if sheet[int_id_col + str(i)].value in duplicates:
                    counter += 1
                continue
        else:
            continue
    INT_VOLUME4 = count
    INT_VALUE4 = sum
    print(f"MTN Portal INT Volume: {count}")
    print(f"MTN Portal INT Value: {sum}")
    print(f"Duplicates:{duplicates}")
    wb.close()

    wb = load_workbook(match_OVA)
    wb.active = wb['Duplicates']
    dup_sheet = wb['Duplicates']
    sheet = dup_sheet

    first_col = wb.active.min_column
    last_col = wb.active.max_column
    first_row = wb.active.min_row
    last_row = wb.active.max_row

    if counter > 0:
        amountColumn = get_column('Amount')
        if amountColumn is None:
            amountColumn = get_column("amount")


        print(f"Number of duplicates: {counter}")
        dup_sheet[amountColumn + str(last_row + 4)].value = str(round(duplicates_value_sum, 2))
        dup_sheet[amountColumn + str(last_row + 5)].value = str(counter)


        DUPLICATES_VOLUME4 = counter
        DUPLICATES_VALUE4 = duplicates_value_sum
    else:
        DUPLICATES_VOLUME4 = 0
        DUPLICATES_VALUE4 = 0

    wb.close()
    wb.save(match_OVA)

    Mwb = load_workbook(match_OVA)

    if matchOVAFound is True and matchINTFound is True:
        # --------------------------------------- MISSING INT TRANSACTIONS ------------------------------------------------
        Owb = load_workbook(match_OVA)
        Osheet = Owb[Owb.sheetnames[0]]
        Owb.active = Owb[Owb.sheetnames[0]]
        file = pd.read_excel(match_OVA)
        sheet = Osheet

        Iwb = load_workbook('Metabase' + yesterday + '.xlsx')
        Isheet = Iwb['Query result']
        sheet = Isheet
        Iwb.active = Iwb['Query result']

        Ofirst_col = Owb.active.min_column
        Olast_col = Owb.active.max_column
        Ofirst_row = Owb.active.min_row
        Olast_row = Owb.active.max_row

        Ifirst_col = Iwb.active.min_column
        Ilast_col = Iwb.active.max_column
        Ifirst_row = Iwb.active.min_row
        Ilast_row = Iwb.active.max_row

        missing_INT_list = []
        missing_OVA_list = []
        matchFound = False
        headerChecker = True
        counter = 0
        for i in range(Ofirst_row + 1, Olast_row + 1):
            for j in range(Ifirst_row + 1, Ilast_row + 1):
                if Osheet[ova_id_col + str(i)].value is None or Osheet[ova_id_col + str(i)].value == "#VALUE!":
                    ova_id_col = alternate_id_col
                    matchFound = False
                    break
                if (Isheet[serviceName_col + str(j)].value == 'MTN OVA') and (
                        Isheet[debitCreditFlagColumn + str(j)].value == 'D'):
                    if str(Osheet[ova_id_col + str(i)].value) in Isheet[int_id_col + str(j)].value:
                        matchFound = True
                        break
                    else:
                        matchFound = False
            if matchFound is False:
                missing_INT_list.append(Osheet[ova_id_col + str(i)].value)
                ova_id_col = tmp_id_col

        print(f"Missing integrator transactions: {missing_INT_list}")

        count = 0
        for transaction in missing_INT_list:
            data = (file[file['External Transaction Id'] == str(transaction)])
            if data.empty:
                data = (file[file['Id'] == transaction])

            if counter > 0:
                headerChecker = False
            with pd.ExcelWriter(match_OVA, mode='a', engine="openpyxl", if_sheet_exists='overlay') as writer:
                data.to_excel(writer, sheet_name='Missing Integrator Transactions', startrow=count,
                              header=headerChecker)
                if headerChecker is False:
                    count += 1
                else:
                    count += 2
                counter += 1

        Iwb.close()
        Owb.close()
        Mwb.close()

        # -------------------------------------- MISSING OVA TRANSACTIONS --------------------------------------------------
        Owb = load_workbook(match_OVA)
        Osheet = Owb[Owb.sheetnames[0]]
        Owb.active = Owb[Owb.sheetnames[0]]
        sheet = Osheet
        Mwb = load_workbook(match_OVA)

        Iwb = load_workbook('Metabase' + yesterday + '.xlsx')
        file = pd.read_excel('Metabase' + yesterday + '.xlsx', sheet_name='Query result')

        Isheet = Iwb['Query result']
        Iwb.active = Iwb['Query result']

        Ofirst_col = Owb.active.min_column
        Olast_col = Owb.active.max_column
        Ofirst_row = Owb.active.min_row
        Olast_row = Owb.active.max_row

        Ifirst_col = Iwb.active.min_column
        Ilast_col = Iwb.active.max_column
        Ifirst_row = Iwb.active.min_row
        Ilast_row = Iwb.active.max_row

        missing_INT_list = []
        missing_OVA_list = []
        matchFound = False
        headerChecker = True
        counter = 0

        for i in range(Ifirst_row + 1, Ilast_row + 1):
            if (Isheet[serviceName_col + str(i)].value == 'MTN OVA') and (
                    Isheet[debitCreditFlagColumn + str(i)].value == 'D'):
                for j in range(Ofirst_row + 1, Olast_row + 1):
                    if str(Osheet[ova_id_col + str(j)].value) in Isheet[int_id_col + str(i)].value:
                        matchFound = True
                        break
                    else:
                        matchFound = False
            else:
                break
            if matchFound is False:
                missing_OVA_list.append(Isheet[int_id_col + str(i)].value)

        print(f"Missing OVA transactions: {missing_OVA_list}")

        count = 0
        for transaction in missing_OVA_list:
            data = file[(file[transIdHeader] == transaction)]
            if counter > 0:
                headerChecker = False
            with pd.ExcelWriter(match_OVA, mode='a', engine="openpyxl", if_sheet_exists='overlay') as writer:
                data.to_excel(writer, sheet_name='Missing OVA Transactions', startrow=count, header=headerChecker)
                if headerChecker is False:
                    count += 1
                else:
                    count += 2
                counter += 1
        Iwb.close()
        Owb.close()
        Mwb.close()

else:
    INT_VOLUME4 = 0.00
    INT_VALUE4 = 0.00
    DUPLICATES_VOLUME4 = 0
    DUPLICATES_VALUE4 = 0
#
#
#
OVA_VOLUME5 = 0
OVA_VALUE5 = 0
INT_VALUE5 = 0
INT_VOLUME5 = 0
DUPLICATES_VALUE5 = 0
DUPLICATES_VOLUME5 = 0
#
#
fileFound = False
matchINTFound = False
matchOVAFound = False
# ########################################### SP AIRTEL CASHOUT OVA #############################################
for file in dir_list:
    if file == 'AirtelTigo Cashout' + yesterday + '.xlsx':
        fileFound = True
        matchOVAFound = True
        wb = load_workbook('AirtelTigo Cashout' + yesterday + '.xlsx')
        sheet = wb[wb.sheetnames[0]]
        wb.active = wb[wb.sheetnames[0]]

        # DEFINE MAX AND MIN COLUMNS AND ROWS
        first_col = wb.active.min_column
        last_col = wb.active.max_column
        first_row = wb.active.min_row
        last_row = wb.active.max_row

        AIRTEL_CASHOUT_OVA_Volume = (last_row + 1) - (first_row + 5)
        OVA_VOLUME6 = AIRTEL_CASHOUT_OVA_Volume
        print('AIRTEL_CASHOUT_OVA_Volume: ' + str(AIRTEL_CASHOUT_OVA_Volume))

        AIRTEL_Cashout_OVA_Sum = 0.00
        amountColumn = get_column('Transaction Amount (GHC.)')
        ova_id_col = get_column('Transaction Id')
        for i in range(first_row + 5, last_row + 1):
            AIRTEL_Cashout_OVA_Sum = AIRTEL_Cashout_OVA_Sum + abs(float(sheet[amountColumn + str(i)].value))

        print('AIRTEL_Cashout_OVA_Sum: ' + str(AIRTEL_Cashout_OVA_Sum))
        OVA_VALUE6 = AIRTEL_Cashout_OVA_Sum
        wb.close()
        wb.save('AirtelTigo Cashout' + yesterday + '- Recons.xlsx')
        match_OVA = 'AirtelTigo Cashout' + yesterday + '- Recons.xlsx'

        wb = load_workbook(match_OVA)
        ws1 = wb.create_sheet("Duplicates")
        ws1.title = 'Duplicates'

        ws1 = wb.create_sheet("Missing Integrator Transactions")
        ws1.title = 'Missing Integrator Transactions'

        ws1 = wb.create_sheet("Missing Integrator Transactions")
        ws1.title = 'Missing Integrator Transactions'

        wb.close()
        wb.save(match_OVA)
        break
    else:
        fileFound = False
if fileFound is False:
    OVA_VOLUME6 = 0
    OVA_VALUE6 = 0
#
#
#
fileFound = False
# ################################################ SP AIRTEL CASHOUT INT ################################################
if metaFileFound is True:

    wb = load_workbook('Metabase' + yesterday + '.xlsx')
    file = pd.read_excel('Metabase' + yesterday + '.xlsx', sheet_name='Query result')

    sheet = wb['Query result']
    wb.active = wb['Query result']

    first_col = wb.active.min_column
    last_col = wb.active.max_column
    first_row = wb.active.min_row
    last_row = wb.active.max_row

    int_id_col = get_column('integratorTransId')
    transIdHeader = 'integratorTransId'
    if int_id_col is None:
        int_id_col = get_column('BillerTransId')
        transIdHeader = "BillerTransId"
    if int_id_col is None:
        int_id_col = get_column("IntegratorTransId")
        transIdHeader = "IntegratorTransId"
    if int_id_col is None:
        int_id_col = get_column('billerTransId')
        transIdHeader = "billerTransId"

    debitCreditFlagColumn = get_column('creditDebitFlag')
    creditDebitHeader = "creditDebitFlag"
    if debitCreditFlagColumn is None:
        debitCreditFlagColumn = get_column("CreditDebitFlag")
        creditDebitHeader = "CreditDebitFlag"
    serviceName_col = get_column('serviceName')
    serviceNameHeader = "serviceName"
    if serviceName_col is None:
        serviceName_col = get_column("ServiceName")
        serviceNameHeader = "ServiceName"
    amountColumn = get_column("amount")
    if amountColumn is None:
        amountColumn = get_column("Amount")

    INT_VALUE6 = 0
    INT_VOLUME6 = 0
    row_count = 1
    count = 0
    sum = 0.00
    counter = 0
    headerChecker = True

    # ------------------------------------------------- Duplicates ---------------------------------------------------------
    list = []
    duplicates = []
    for i in range(first_row + 1, last_row + 1):
        if ((sheet[serviceName_col + str(i)].value == 'Airtel Money Agent') or (
                sheet[serviceName_col + str(i)].value == 'AirtelMoney_Slydepay')) and sheet[debitCreditFlagColumn + str(i)].value == 'D':
            count += 1
            sum = sum + abs(float(sheet[amountColumn + str(i)].value))
            if sheet[int_id_col + str(i)].value not in list:
                list.append(sheet[int_id_col + str(i)].value)
            elif sheet[int_id_col + str(i)].value not in duplicates:
                    data = file[(file[transIdHeader] == str(sheet[int_id_col + str(i)].value)) & (file[serviceNameHeader] == 'Airtel Money Agent') & (file[creditDebitHeader] == 'D')]
                    try:
                        data_set_sum = data.iloc[1:, data.columns.get_loc("Amount")].sum()
                    except:
                        data_set_sum = data.iloc[1:, data.columns.get_loc("amount")].sum()
                    duplicates_value_sum = duplicates_value_sum + data_set_sum
                    duplicates.append(sheet[int_id_col + str(i)].value)
                    if counter > 0:
                        headerChecker = False
                    with pd.ExcelWriter(match_OVA, mode='a', engine="openpyxl", if_sheet_exists='overlay') as writer:
                        data.to_excel(writer, sheet_name='Duplicates', startrow=row_count, header=headerChecker)
                        if headerChecker is False:
                            row_count += 2
                        else:
                            row_count += (data.shape[0]+1)
                        counter += 1
            else:
                if sheet[int_id_col + str(i)].value in duplicates:
                    counter += 1
                continue
        else:
            continue
    INT_VOLUME6 = count
    INT_VALUE6 = sum
    print(f"AIRTEL_CASHOUT INT Volume: {count}")
    print(f"AIRTEL Cashout INT Value: {sum}")
    print(f"Duplicates:{duplicates}")
    wb.close()

    if matchOVAFound is True:
        wb = load_workbook(match_OVA)
        wb.active = wb['Duplicates']
        dup_sheet = wb['Duplicates']
        sheet = dup_sheet

        first_col = wb.active.min_column
        last_col = wb.active.max_column
        first_row = wb.active.min_row
        last_row = wb.active.max_row

        if counter > 0:
            amountColumn = get_column('Amount')
            if amountColumn is None:
                amountColumn = get_column("amount")

            print(f"Number of duplicates: {counter}")
            dup_sheet[amountColumn + str(last_row + 4)].value = str(round(duplicates_value_sum, 2))
            dup_sheet[amountColumn + str(last_row + 5)].value = str(counter)


            DUPLICATES_VOLUME6 = counter
            DUPLICATES_VALUE6 = duplicates_value_sum
        else:
            DUPLICATES_VOLUME6 = 0
            DUPLICATES_VALUE6 = 0

        wb.close()
        wb.save(match_OVA)

        # ------------------------------------- MISSING INT TRANSACTIONS ------------------------------------------------
        if matchOVAFound is True and metaFileFound is True:
            Owb = load_workbook(match_OVA)
            Osheet = Owb[Owb.sheetnames[0]]
            Owb.active = Owb[Owb.sheetnames[0]]
            Iwb = load_workbook('Metabase' + yesterday + '.xlsx')
            Isheet = Iwb['Query result']
            Iwb.active = Iwb['Query result']

            Ofirst_col = Owb.active.min_column
            Olast_col = Owb.active.max_column
            Ofirst_row = Owb.active.min_row
            Olast_row = Owb.active.max_row

            Ifirst_col = Iwb.active.min_column
            Ilast_col = Iwb.active.max_column
            Ifirst_row = Iwb.active.min_row
            Ilast_row = Iwb.active.max_row

            missing_INT_list = []
            missing_OVA_list = []
            matchFound = False
            headerChecker = True
            counter = 0

            for i in range(Ofirst_row + 6, Olast_row + 1):
                for j in range(Ifirst_row + 1, Ilast_row + 1):
                    if ((Isheet[serviceName_col + str(j)].value == 'Airtel Money Agent') or (
                            Isheet[serviceName_col + str(j)].value == 'AirtelMoney_Slydepay')) and Isheet[
                        debitCreditFlagColumn + str(j)].value == 'D':
                        if str(Osheet[ova_id_col + str(i)].value) in str(Isheet[int_id_col + str(j)].value):
                            matchFound = True
                            break
                        else:
                            matchFound = False
                if matchFound is False:
                    missing_INT_list.append(Osheet[ova_id_col + str(i)].value)
            print(f"Missing integrator transactions: {missing_INT_list}")
            file = pd.read_excel(match_OVA, header=[5])
            count = 0
            for transaction in missing_INT_list:
                data = file[file['Transaction Id'] == str(transaction)]
                if counter > 0:
                    headerChecker = False
                with pd.ExcelWriter(match_OVA, mode='a', engine="openpyxl", if_sheet_exists='overlay') as writer:
                    data.to_excel(writer, sheet_name='Missing Integrator Transactions', startrow=count,
                                  header=headerChecker)
                    if headerChecker is False:
                        count += 1
                    else:
                        count += 2
                    counter += 1
            Iwb.close()
            Owb.close()
            # Mwb.close()

            #        # -------------------------------------- MISSING OVA TRANSACTIONS --------------------------------------------------
            Owb = load_workbook(match_OVA)
            Osheet = Owb[Owb.sheetnames[0]]
            Owb.active = Owb[Owb.sheetnames[0]]
            sheet = Osheet
            Mwb = load_workbook(match_OVA)

            Iwb = load_workbook('Metabase' + yesterday + '.xlsx')
            file = pd.read_excel('Metabase' + yesterday + '.xlsx', sheet_name='Query result')
            Isheet = Iwb['Query result']
            Iwb.active = Iwb['Query result']

            Ofirst_col = Owb.active.min_column
            Olast_col = Owb.active.max_column
            Ofirst_row = Owb.active.min_row
            Olast_row = Owb.active.max_row

            Ifirst_col = Iwb.active.min_column
            Ilast_col = Iwb.active.max_column
            Ifirst_row = Iwb.active.min_row
            Ilast_row = Iwb.active.max_row

            missing_INT_list = []
            missing_OVA_list = []
            matchFound = False
            headerChecker = True
            counter = 0

            for i in range(Ifirst_row + 1, Ilast_row + 1):
                if ((Isheet[serviceName_col + str(i)].value == 'Airtel Money Agent') or (
                        Isheet[serviceName_col + str(i)].value == 'AirtelMoney_Slydepay')) and Isheet[
                    debitCreditFlagColumn + str(i)].value == 'D':
                    for j in range(Ofirst_row + 1, Olast_row + 1):
                        if str(Osheet[ova_id_col + str(j)].value) in str(Isheet[int_id_col + str(i)].value):
                            matchFound = True
                            break
                        else:
                            matchFound = False
                    if matchFound is False:
                        missing_OVA_list.append(Isheet[int_id_col + str(i)].value)
                else:
                    continue

            print(f"Missing OVA transactions: {missing_OVA_list}")

            count = 0
            for transaction in missing_OVA_list:
                data = file[(file[transIdHeader] == str(transaction))]
                if counter > 0:
                    headerChecker = False
                with pd.ExcelWriter(match_OVA, mode='a', engine="openpyxl", if_sheet_exists='overlay') as writer:
                    data.to_excel(writer, sheet_name='Missing OVA Transactions', startrow=count, header=headerChecker)
                    if headerChecker is False:
                        count += 1
                    else:
                        count += 2
                    counter += 1

            Iwb.close()
            Owb.close()
    else:
        DUPLICATES_VOLUME6 = 0
        DUPLICATES_VALUE6 = 0
else:
    INT_VOLUME6 = 0.00
    INT_VALUE6 = 0.00
    DUPLICATES_VOLUME6 = 0
    DUPLICATES_VALUE6 = 0

#
#
matchOVAFound = False
matchINTFound = False
fileFound = False
# ###################################### SP VODAFONE CASHIN OVA #######################################################

for file in dir_list:
    if file == 'Vodafone Cashin' + yesterday + '.xlsx':
        fileFound = True
        matchOVAFound = True
        wb = load_workbook('Vodafone Cashin' + yesterday + '.xlsx')
        sheet = wb[wb.sheetnames[0]]
        wb.active = wb[wb.sheetnames[0]]
        ids = []

        # DEFINE MAX AND MIN COLUMNS AND ROWS
        first_col = wb.active.min_column
        last_col = wb.active.max_column
        first_row = wb.active.min_row
        last_row = wb.active.max_row

        VODA_CASHIN_OVA_Volume = (last_row + 1) - (first_row + 6)
        OVA_VOLUME7 = VODA_CASHIN_OVA_Volume
        print('VODA_CASHIN_OVA_Volume: ' + str(VODA_CASHIN_OVA_Volume))

        VODA_Cashin_OVA_Sum = 0.00
        amountColumn = get_column('Paid In')
        for i in range(first_row + 6, last_row + 1):
            if ((sheet[amountColumn + str(i)].value) == '') or ((sheet[amountColumn + str(i)].value) is None):
                continue
            VODA_Cashin_OVA_Sum = VODA_Cashin_OVA_Sum + abs(float(sheet[amountColumn + str(i)].value))

        partyIdCol = get_column('Details')
        ova_id_col = get_column('TransId')
        tmp_id_col = ova_id_col
        alternate_id_col = get_column("Receipt No.")
        for i in range(first_row + 6, last_row + 1):
            ids.append(sheet[ova_id_col + str(i)].value)

        print('VODA_Cashin_OVA_Sum: ' + str(VODA_Cashin_OVA_Sum))
        OVA_VALUE7 = VODA_Cashin_OVA_Sum
        wb.close()
        wb.save('Vodafone Cashin' + yesterday + '- Recons.xlsx')
        match_OVA = 'Vodafone Cashin' + yesterday + '- Recons.xlsx'

        wb = load_workbook(match_OVA)
        ws1 = wb.create_sheet("Duplicates")
        ws1.title = 'Duplicates'

        ws1 = wb.create_sheet("Missing Integrator Transactions")
        ws1.title = 'Missing Integrator Transactions'

        ws1 = wb.create_sheet("Missing Integrator Transactions")
        ws1.title = 'Missing Integrator Transactions'

        wb.close()
        wb.save(match_OVA)
        break
    else:
        fileFound = False

if fileFound is False:
    OVA_VOLUME7 = 0
    OVA_VALUE7 = 0
#
#
# ############################################# SP VODA CASHIN INT ####################################################
if metaFileFound is True:
    wb = load_workbook('Metabase' + yesterday + '.xlsx')
    file = pd.read_excel('Metabase' + yesterday + '.xlsx', sheet_name='Query result')

    sheet = wb['Query result']
    wb.active = wb['Query result']
    first_col = wb.active.min_column
    last_col = wb.active.max_column
    first_row = wb.active.min_row
    last_row = wb.active.max_row

    int_id_col = get_column('integratorTransId')
    transIdHeader = "integratorTransId"
    if int_id_col is None:
        int_id_col = get_column("IntegratorTransId")
        transIdHeader = "IntegratorTransId"
    debitCreditFlagColumn = get_column('creditDebitFlag')
    creditDebitHeader = "creditDebitFlag"
    if debitCreditFlagColumn is None:
        debitCreditFlagColumn = get_column("CreditDebitFlag")
        creditDebitHeader = "CreditDebitFlag"
    serviceName_col = get_column('serviceName')
    serviceNameHeader = "serviceName"
    if serviceName_col is None:
        serviceName_col = get_column("ServiceName")
        serviceNameHeader = "ServiceName"
    amountColumn = get_column("amount")
    if amountColumn is None:
        amountColumn = get_column("Amount")

    INT_VOLUME7 = 0.00
    INT_VALUE7 = 0.00
    row_count = 1
    count = 0
    sum = 0.00
    counter = 0
    headerChecker = True
    # ------------------------------------------------- Duplicates ---------------------------------------------------------
    list = []
    duplicates = []
    duplicates_value_sum = 0.00
    for i in range(first_row + 1, last_row + 1):
        if sheet[serviceName_col + str(i)].value == 'Vodafone Cash' and sheet[debitCreditFlagColumn + str(i)].value == 'C':
            count += 1
            sum = sum + abs(float(sheet[amountColumn + str(i)].value))
            if sheet[int_id_col + str(i)].value not in list:
                list.append(sheet[int_id_col + str(i)].value)
            elif sheet[int_id_col + str(i)].value not in duplicates:
                    data = file[(file[transIdHeader] == sheet[int_id_col + str(i)].value) & (file[serviceNameHeader] == 'Vodafone Cash') & (file[creditDebitHeader] == 'C')]
                    try:
                        data_set_sum = data.iloc[1:, data.columns.get_loc("Amount")].sum()
                    except:
                        data_set_sum = data.iloc[1:, data.columns.get_loc("amount")].sum()
                    duplicates_value_sum = duplicates_value_sum + data_set_sum
                    duplicates.append(sheet[int_id_col + str(i)].value)
                    if counter > 0:
                        headerChecker = False
                    with pd.ExcelWriter(match_OVA, mode='a', engine="openpyxl", if_sheet_exists='overlay') as writer:
                        data.to_excel(writer, sheet_name='Duplicates', startrow=row_count, header=headerChecker)
                        if headerChecker is False:
                            row_count += 2
                        else:
                            row_count += (data.shape[0]+1)
                        counter = 1
            else:
                if sheet[int_id_col + str(i)].value in duplicates:
                    counter += 1
                continue
        else:
            continue
    INT_VOLUME7 = count
    INT_VALUE7 = sum
    print(f"Vodafone Cashin INT Volume: {count}")
    print(f"Vodafone Cashin INT Value: {sum}")
    print(f"Duplicates:{duplicates}")
    wb.close()

    wb = load_workbook(match_OVA)
    wb.active = wb['Duplicates']
    dup_sheet = wb['Duplicates']
    sheet = dup_sheet

    first_col = wb.active.min_column
    last_col = wb.active.max_column
    first_row = wb.active.min_row
    last_row = wb.active.max_row

    if counter > 0:
        amountColumn = get_column('Amount')
        if amountColumn is None:
            amountColumn = get_column("amount")


        print(f"Number of duplicates: {counter}")
        dup_sheet[amountColumn + str(last_row + 4)].value = str(round(duplicates_value_sum, 2))
        dup_sheet[amountColumn + str(last_row + 5)].value = str(counter)

        DUPLICATES_VOLUME7 = counter
        DUPLICATES_VALUE7 = duplicates_value_sum
    else:
        DUPLICATES_VOLUME7 = 0
        DUPLICATES_VALUE7 = 0

    wb.close()
    wb.save(match_OVA)

    # ------------------------------------- MISSING INT TRANSACTIONS ------------------------------------------------
    if matchOVAFound is True and metaFileFound is True:
        Owb = load_workbook(match_OVA)
        Osheet = Owb[Owb.sheetnames[0]]
        Owb.active = Owb[Owb.sheetnames[0]]

        Ofirst_col = Owb.active.min_column
        Olast_col = Owb.active.max_column
        Ofirst_row = Owb.active.min_row
        Olast_row = Owb.active.max_row

        Iwb = load_workbook('Metabase' + yesterday + '.xlsx')
        Isheet = Iwb['Query result']
        Iwb.active = Iwb['Query result']

        Mwb = load_workbook(match_OVA)
        sheet = Mwb[Mwb.sheetnames[0]]
        Mwb.active = Mwb[Mwb.sheetnames[0]]

        Ifirst_col = Iwb.active.min_column
        Ilast_col = Iwb.active.max_column
        Ifirst_row = Iwb.active.min_row
        Ilast_row = Iwb.active.max_row

        missing_INT_list = []
        missing_OVA_list = []
        matchFound = False
        headerChecker = True
        counter = 0

        for i in range(Ofirst_row + 6, Olast_row + 1):
            for j in range(Ifirst_row + 1, Ilast_row + 1):
                if Osheet[ova_id_col + str(i)].value is None or Osheet[ova_id_col + str(i)].value == "#VALUE!":
                    ova_id_col = alternate_id_col
                    matchFound = False
                    break
                if Isheet[serviceName_col + str(j)].value == 'Vodafone Cash' and Isheet[
                    debitCreditFlagColumn + str(j)].value == 'C':
                    if str(Osheet[ova_id_col + str(i)].value) in str(Isheet[int_id_col + str(j)].value):
                        matchFound = True
                        break
                    else:
                        matchFound = False
            if matchFound is False:
                missing_INT_list.append(str(Osheet[ova_id_col + str(i)].value))
                ova_id_col = tmp_id_col
        print(f"Missing integrator transactions: {missing_INT_list}")
        file = pd.read_excel(match_OVA, header=[5])
        count = 0
        for transaction in missing_INT_list:
            data = file[file['TransId'] == transaction]
            if counter > 0:
                headerChecker = False
            with pd.ExcelWriter(match_OVA, mode='a', engine="openpyxl", if_sheet_exists='overlay') as writer:
                data.to_excel(writer, sheet_name='Missing Integrator Transactions', startrow=count,
                              header=headerChecker)
                if headerChecker is False:
                    count += 1
                else:
                    count += 2
                counter += 1
        Iwb.close()
        Mwb.close()

        # ------------------------------------- MISSING OVA TRANSACTIONS --------------------------------------------
        Iwb = load_workbook('Metabase' + yesterday + '.xlsx')
        Isheet = Iwb['Query result']
        Iwb.active = Iwb['Query result']

        Ifirst_col = Iwb.active.min_column
        Ilast_col = Iwb.active.max_column
        Ifirst_row = Iwb.active.min_row
        Ilast_row = Iwb.active.max_row

        missing_INT_list = []
        missing_OVA_list = []
        matchFound = False
        headerChecker = False
        counter = 0

        for i in range(Ifirst_row + 1, Ilast_row + 1):
            if Isheet[serviceName_col + str(i)].value == 'Vodafone Cash' and Isheet[debitCreditFlagColumn + str(i)].value == 'C':
                for j in ids:
                    if str(j) in str(Isheet[int_id_col + str(i)].value):
                        matchFound = True
                        break
                    else:
                        matchFound = False
                if matchFound is False:
                    missing_OVA_list.append(Isheet[int_id_col + str(i)].value)
        print(f"Missing OVA transactions: {missing_OVA_list}")
        file = pd.read_excel('Metabase' + yesterday + '.xlsx', sheet_name='Query result')
        count = 0
        for transaction in missing_OVA_list:
            data = file[(file[transIdHeader] == transaction) & (file[serviceNameHeader] == 'Vodafone Cash') & (
                    file[creditDebitHeader] == 'C')]
            if counter > 0:
                headerChecker = False
            with pd.ExcelWriter(match_OVA, mode='a', engine="openpyxl", if_sheet_exists='overlay') as writer:
                data.to_excel(writer, sheet_name='Missing OVA Transactions', startrow=count, header=headerChecker)
                if headerChecker is False:
                    count += 1
                else:
                    count += 2
                counter += 1
        Iwb.close()


else:
    INT_VOLUME7 = 0
    INT_VALUE7 = 0
    DUPLICATES_VOLUME7 = 0
    DUPLICATES_VALUE7 = 0


#
#
#
matchOVAFound = False
matchINTFound = False
fileFound = False

#  ########################################## VODA SP CASHOUT OVA #####################################################
for file in dir_list:
    if file == 'Vodafone Cashout' + yesterday + '.xlsx':
        fileFound = True
        matchOVAFound = True
        wb = load_workbook('Vodafone Cashout' + yesterday + '.xlsx')
        sheet = wb[wb.sheetnames[0]]
        wb.active = wb[wb.sheetnames[0]]
        ids = []

        # DEFINE MAX AND MIN COLUMNS AND ROWS
        first_col = wb.active.min_column
        last_col = wb.active.max_column
        first_row = wb.active.min_row
        last_row = wb.active.max_row

        VODA_CASHOUT_OVA_Volume = (last_row + 1) - (first_row + 6)
        OVA_VOLUME8 = VODA_CASHOUT_OVA_Volume
        print('VODA_CASHOUT_OVA_Volume: ' + str(VODA_CASHOUT_OVA_Volume))

        VODA_Cashout_OVA_Sum = 0.00
        amountColumn = get_column('Withdrawn')
        for i in range(first_row + 6, last_row + 1):
            if ((sheet[amountColumn + str(i)].value) == '') or ((sheet[amountColumn + str(i)].value) is None):
                continue
            VODA_Cashout_OVA_Sum = VODA_Cashout_OVA_Sum + abs(float(sheet[amountColumn + str(i)].value))

        partyIdCol = get_column('Details')
        ova_id_col = get_column('TransId')
        tmp_id_col = ova_id_col
        alternate_id_col = get_column("Receipt No.")
        for i in range(first_row + 6, last_row + 1):
            ids.append(str(sheet[ova_id_col + str(i)].value))
        print('VODA_Cashout_OVA_Sum: ' + str(VODA_Cashout_OVA_Sum))
        OVA_VALUE8 = VODA_Cashout_OVA_Sum
        wb.close()
        wb.save('Vodafone Cashout' + yesterday + '- Recons.xlsx')
        match_OVA = 'Vodafone Cashout' + yesterday + '- Recons.xlsx'

        wb = load_workbook(match_OVA)
        ws1 = wb.create_sheet("Duplicates")
        ws1.title = 'Duplicates'

        ws1 = wb.create_sheet("Missing Integrator Transactions")
        ws1.title = 'Missing Integrator Transactions'

        ws1 = wb.create_sheet("Missing Integrator Transactions")
        ws1.title = 'Missing Integrator Transactions'

        wb.close()
        wb.save(match_OVA)
        break
    else:
        fileFound = False

if fileFound is False:
    OVA_VOLUME8 = 0
    OVA_VALUE8 = 0
#
# ########################################## SP VODA CASHOUT INT #################################################
if metaFileFound is True:
    wb = load_workbook('Metabase' + yesterday + '.xlsx')
    file = pd.read_excel('Metabase' + yesterday + '.xlsx')

    sheet = wb['Query result']
    wb.active = wb['Query result']
    first_col = wb.active.min_column
    last_col = wb.active.max_column
    first_row = wb.active.min_row
    last_row = wb.active.max_row

    int_id_col = get_column('integratorTransId')
    transIdHeader = "integratorTransId"
    if int_id_col is None:
        int_id_col = get_column("IntegratorTransId")
        transIdHeader = "IntegratorTransId"
    debitCreditFlagColumn = get_column('creditDebitFlag')
    creditDebitHeader = "creditDebitFlag"
    if debitCreditFlagColumn is None:
        debitCreditFlagColumn = get_column("CreditDebitFlag")
        creditDebitHeader = "CreditDebitFlag"
    serviceName_col = get_column('serviceName')
    serviceNameHeader = "serviceName"
    if serviceName_col is None:
        serviceName_col = get_column("ServiceName")
        serviceNameHeader = "ServiceName"
    amountColumn = get_column("amount")
    if amountColumn is None:
        amountColumn = get_column("Amount")

    INT_VOLUME8 = 0.00
    INT_VALUE8 = 0.00
    row_count = 1
    count = 0
    sum = 0.00
    counter = 0
    headerChecker = True

    # ------------------------------------------------- Duplicates ---------------------------------------------------------
    list = []
    duplicates = []
    duplicates_value_sum = 0.00
    for i in range(first_row + 1, last_row + 1):
        if sheet[serviceName_col + str(i)].value == 'Vodafone Cash' and sheet[
            debitCreditFlagColumn + str(i)].value == 'D':
            count += 1
            sum = sum + abs(float(sheet[amountColumn + str(i)].value))
            if sheet[int_id_col + str(i)].value not in list:
                list.append(sheet[int_id_col + str(i)].value)
            elif sheet[int_id_col + str(i)].value not in duplicates:
                    data = file[(file[transIdHeader] == str(sheet[int_id_col + str(i)].value)) & (file[serviceNameHeader] == 'Vodafone Cash') & (file[creditDebitHeader] == 'D')]
                    try:
                        data_set_sum = data.iloc[1:, data.columns.get_loc("Amount")].sum()
                    except:
                        data_set_sum = data.iloc[1:, data.columns.get_loc("amount")].sum()
                    duplicates_value_sum = duplicates_value_sum + data_set_sum
                    duplicates.append(sheet[int_id_col + str(i)].value)
                    if counter > 0:
                        headerChecker = False
                    with pd.ExcelWriter(match_OVA, mode='a', engine="openpyxl", if_sheet_exists='overlay') as writer:
                        data.to_excel(writer, sheet_name='Duplicates', startrow=row_count, header=headerChecker)
                        if headerChecker is False:
                            row_count += 2
                        else:
                            row_count += (data.shape[0] +1)
                        counter += 1
            else:
                if sheet[int_id_col + str(i)].value in duplicates:
                    counter += 1
                continue
        else:
            continue
    INT_VOLUME8 = count
    INT_VALUE8 = sum
    print(f"Vodafone Cashout INT Volume: {count}")
    print(f"Vodafone Cashout INT Value: {sum}")
    print(f"Duplicates:{duplicates}")
    wb.close()

    wb = load_workbook(match_OVA)
    wb.active = wb['Duplicates']
    dup_sheet = wb['Duplicates']
    sheet = dup_sheet

    first_col = wb.active.min_column
    last_col = wb.active.max_column
    first_row = wb.active.min_row
    last_row = wb.active.max_row

    if counter > 0:
        amountColumn = get_column('Amount')
        if amountColumn is None:
            amountColumn = get_column("amount")


        print(f"Number of duplicates: {counter}")
        dup_sheet[amountColumn + str(last_row + 4)].value = str(round(duplicates_value_sum, 2))
        dup_sheet[amountColumn + str(last_row + 5)].value = str(counter)


        DUPLICATES_VOLUME8 = counter
        DUPLICATES_VALUE8 = duplicates_value_sum
    else:
        DUPLICATES_VOLUME8 = 0
        DUPLICATES_VALUE8 = 0

    wb.close()
    wb.save(match_OVA)

    # ------------------------------------- MISSING INT TRANSACTIONS ------------------------------------------------
    if matchOVAFound is True and metaFileFound is True:
        Owb = load_workbook(match_OVA)
        Osheet = Owb[Owb.sheetnames[0]]
        Owb.active = Owb[Owb.sheetnames[0]]

        Ofirst_col = Owb.active.min_column
        Olast_col = Owb.active.max_column
        Ofirst_row = Owb.active.min_row
        Olast_row = Owb.active.max_row

        Iwb = load_workbook('Metabase' + yesterday + '.xlsx')
        Isheet = Iwb['Query result']
        Iwb.active = Iwb['Query result']

        Ifirst_col = Iwb.active.min_column
        Ilast_col = Iwb.active.max_column
        Ifirst_row = Iwb.active.min_row
        Ilast_row = Iwb.active.max_row

        Mwb = load_workbook(match_OVA)
        sheet = Mwb[Mwb.sheetnames[0]]
        Mwb.active = Mwb[Mwb.sheetnames[0]]

        missing_INT_list = []
        missing_OVA_list = []
        matchFound = False
        counter = 0
        headerChecker = True

        for i in range(Ofirst_row + 6, Olast_row + 1):
            for j in range(Ifirst_row + 1, Ilast_row + 1):
                if Osheet[ova_id_col + str(i)].value is None or Osheet[ova_id_col + str(i)].value == "#VALUE!":
                    ova_id_col = alternate_id_col
                    matchFound = False
                    break
                if Isheet[serviceName_col + str(j)].value == 'Vodafone Cash' and Isheet[
                    debitCreditFlagColumn + str(j)].value == 'D':
                    if str(Isheet[int_id_col + str(j)].value) in str(Osheet[ova_id_col + str(i)].value):
                        matchFound = True
                        break
                    else:
                        matchFound = False
            if matchFound is False:
                missing_INT_list.append(Osheet[ova_id_col + str(i)].value)
                ova_id_col = tmp_id_col
        print(f"Missing integrator transactions: {missing_INT_list}")
        file = pd.read_excel(match_OVA, header=[5])
        count = 0
        for transaction in missing_INT_list:
            data = file[file['TransId'] == int(transaction)]
            if data.empty:
                data = file[file['Receipt No.'] == int(transaction)]
            if counter > 0:
                headerChecker = False
            with pd.ExcelWriter(match_OVA, mode='a', engine="openpyxl", if_sheet_exists='overlay') as writer:
                data.to_excel(writer, sheet_name='Missing Integrator Transactions', startrow=count,
                              header=headerChecker)
                if headerChecker is False:
                    count += 1
                else:
                    count += 2
                counter += 1
        Iwb.close()
        Mwb.close()

        # ------------------------------------- MISSING OVA TRANSACTIONS --------------------------------------------
        Iwb = load_workbook('Metabase' + yesterday + '.xlsx')
        Isheet = Iwb['Query result']
        Iwb.active = Iwb['Query result']

        Mwb = load_workbook(match_OVA)

        Ifirst_col = Iwb.active.min_column
        Ilast_col = Iwb.active.max_column
        Ifirst_row = Iwb.active.min_row
        Ilast_row = Iwb.active.max_row

        missing_INT_list = []
        missing_OVA_list = []
        matchFound = False
        headerChecker = True
        counter = 0

        for i in range(Ifirst_row + 1, Ilast_row + 1):
            if Isheet[serviceName_col + str(i)].value == 'Vodafone Cash' and Isheet[debitCreditFlagColumn + str(i)].value == 'D':
                for j in ids:
                    if str(j) in str(Isheet[int_id_col + str(i)].value):
                        matchFound = True
                        break
                    else:
                        matchFound = False
                if matchFound is False:
                    missing_OVA_list.append(Isheet[int_id_col + str(i)].value)
        print(f"Missing OVA transactions: {missing_OVA_list}")
        file = pd.read_excel('Metabase' + yesterday + '.xlsx', sheet_name='Query result')
        count = 0
        for transaction in missing_OVA_list:
            data = file[(file[transIdHeader] == str(transaction)) & (file[serviceNameHeader] == 'Vodafone Cash') & (
                        file[creditDebitHeader] == 'D')]
            if counter > 0:
                headerChecker = False
            with pd.ExcelWriter(match_OVA, mode='a', engine="openpyxl", if_sheet_exists='overlay') as writer:
                data.to_excel(writer, sheet_name='Missing OVA Transactions', startrow=count, header=headerChecker)
                if headerChecker is False:
                    count += 1
                else:
                    count += 2
                counter += 1
        Iwb.close()
        Mwb.close()

else:
    INT_VOLUME8 = 0
    INT_VALUE8 = 0
    DUPLICATES_VOLUME8 = 0
    DUPLICATES_VALUE8 = 0


#
#
matchOVAFound = False
matchINTFound = False
fileFound = False
#  ######################################## STANBIC FI CR OVA #####################################################
for file in dir_list:
    if file == 'Stanbic FI Credit' + yesterday + '.xlsx':
        fileFound = True
        matchOVAFound = True
        wb = load_workbook('Stanbic FI Credit' + yesterday + '.xlsx')
        sheet = wb[wb.sheetnames[0]]
        # DEFINE MAX AND MIN COLUMNS AND ROWS
        first_col = wb.active.min_column
        last_col = wb.active.max_column
        first_row = wb.active.min_row
        last_row = wb.active.max_row

        Stanbic_FI_Credit_OVA_Volume = 0

        Stanbic_FI_Credit_OVA_Sum = 0.00

        amountColumn = get_column('AMOUNT')
        debitCreditFlagColumn = get_column('DEBITCREDIT')
        ova_id_col = get_column('REMARKS2')
        for i in range(first_row + 1, last_row + 1):
            if sheet[debitCreditFlagColumn + str(i)].value == 'C':
                Stanbic_FI_Credit_OVA_Volume = Stanbic_FI_Credit_OVA_Volume + 1
                Stanbic_FI_Credit_OVA_Sum = Stanbic_FI_Credit_OVA_Sum + abs(float(sheet[amountColumn + str(i)].value))

        print('Stanbic_FI_CREDIT_OVA_Volume: ' + str(Stanbic_FI_Credit_OVA_Volume))
        print('Stanbic_FI_CREDIT_OVA_Sum: ' + str(Stanbic_FI_Credit_OVA_Sum))
        OVA_VOLUME9 = Stanbic_FI_Credit_OVA_Volume
        OVA_VALUE9 = Stanbic_FI_Credit_OVA_Sum
        wb.close()
        wb.save('Stanbic FI Credit' + yesterday + '- Recons.xlsx')
        match_OVA = 'Stanbic FI Credit' + yesterday + '- Recons.xlsx'

        wb = load_workbook(match_OVA)
        ws1 = wb.create_sheet("Duplicates")
        ws1.title = 'Duplicates'

        ws1 = wb.create_sheet("Missing Integrator Transactions")
        ws1.title = 'Missing Integrator Transactions'

        ws1 = wb.create_sheet("Missing OVA Transactions")
        ws1.title = 'Missing OVA Transactions'

        wb.close()
        wb.save(match_OVA)
        break
    else:
        fileFound = False

if fileFound is False:
    OVA_VOLUME9 = 0
    OVA_VALUE9 = 0
#
#
#

#  ############################################# STANBIC FI CREDIT INT ################################################
for file in dir_list:
    if file == 'Stanbic FI Credit Metabase' + yesterday + '.xlsx':
        fileFound = True
        matchINTFound = True
        wb = load_workbook('Stanbic FI Credit Metabase' + yesterday + '.xlsx')
        sheet = wb[wb.sheetnames[0]]
        wb.active = wb[wb.sheetnames[0]]
        file = pd.read_excel('Stanbic FI Credit Metabase' + yesterday + '.xlsx')
        match_INT = 'Stanbic FI Credit Metabase' + yesterday + '.xlsx'


        first_col = wb.active.min_column
        last_col = wb.active.max_column
        first_row = wb.active.min_row
        last_row = wb.active.max_row

        Stanbic_FI_Credit_INT_Volume = (last_row + 1) - (first_row + 1)
        Stanbic_FI_Credit_INT_Sum = 0.00
        INT_VOLUME9 = Stanbic_FI_Credit_INT_Volume

        amountColumn = get_column('Amount ($)')
        if amountColumn is None:
            amountColumn = get_column("Actual Amount ($)")
        int_id_col = get_column('External Payment Request → Institution Trans ID')
        for i in range(first_row + 1, last_row + 1):
            Stanbic_FI_Credit_INT_Sum = Stanbic_FI_Credit_INT_Sum + abs(float(sheet[amountColumn + str(i)].value))

        print('Stanbic_FI_Credit_INT_Volume: ' + str(Stanbic_FI_Credit_INT_Volume))
        print('Stanbic_FI_Credit_INT_Sum: ' + str(Stanbic_FI_Credit_INT_Sum))
        INT_VALUE9 = Stanbic_FI_Credit_INT_Sum

        wb.close()
        wb.save('Stanbic FI Credit Metabase' + yesterday + '.xlsx')
        # ---------------------------------------- Duplicates -----------------------------------------------------
        wb = load_workbook('Stanbic FI Credit Metabase' + yesterday + '.xlsx')
        sheet = wb[wb.sheetnames[0]]
        list = []
        duplicates = []
        count = 1
        counter = 0
        headerChecker = True
        duplicates_value_sum = 0.00
        for id in range(first_row + 1, last_row + 1):
            if sheet[int_id_col + str(id)].value not in list:
                list.append(sheet[int_id_col + str(id)].value)
            elif sheet[int_id_col + str(id)].value not in duplicates:
                    data = file[file['External Payment Request → Institution Trans ID'] == (sheet[int_id_col + str(id)].value)]
                    try:
                        data_set_sum = data.iloc[1:, data.columns.get_loc("Amount ($)")].sum()
                    except:
                        data_set_sum = data.iloc[1:, data.columns.get_loc("Actual Amount ($)")].sum()

                    duplicates_value_sum = duplicates_value_sum + data_set_sum
                    duplicates.append(sheet[int_id_col + str(id)].value)
                    if counter > 0:
                        headerChecker = False
                    with pd.ExcelWriter(match_OVA, mode='a', engine="openpyxl", if_sheet_exists='overlay') as writer:
                        data.to_excel(writer, sheet_name='Duplicates', startrow=count, header=headerChecker)
                        if headerChecker is False:
                            count += 2
                        else:
                            count += (data.shape[0] +1)
                        counter += 1
            else:
                if sheet[int_id_col + str(id)].value in duplicates:
                    counter += 1
                continue
        wb.close()

        wb = load_workbook(match_OVA)
        wb.active = wb['Duplicates']
        dup_sheet = wb['Duplicates']
        sheet = dup_sheet

        first_col = wb.active.min_column
        last_col = wb.active.max_column
        first_row = wb.active.min_row
        last_row = wb.active.max_row

        dup_col = column_index_from_string(amountColumn) + 1
        print(f"Number of duplicates: {counter}")
        dup_sheet[get_column_letter(dup_col) + str(last_row + 4)].value = str(round(duplicates_value_sum, 2))
        dup_sheet[get_column_letter(dup_col) + str(last_row + 5)].value = str(counter)

        print(f"Duplicates: {duplicates}")
        DUPLICATES_VOLUME9 = counter
        DUPLICATES_VALUE9 = duplicates_value_sum

        wb.close()
        wb.save(match_OVA)
        if matchOVAFound is True and matchINTFound is True:
            # ################################## MISSING INT TRANSACTIONS #########################################

            Owb = load_workbook(match_OVA)
            Osheet = Owb[Owb.sheetnames[0]]
            Owb.active = Owb[Owb.sheetnames[0]]

            Iwb = load_workbook(match_INT)
            Isheet = Iwb[Iwb.sheetnames[0]]
            Iwb.active = Iwb[Iwb.sheetnames[0]]

            Ofirst_col = Owb.active.min_column
            Olast_col = Owb.active.max_column
            Ofirst_row = Owb.active.min_row
            Olast_row = Owb.active.max_row

            Ifirst_col = Iwb.active.min_column
            Ilast_col = Iwb.active.max_column
            Ifirst_row = Iwb.active.min_row
            Ilast_row = Iwb.active.max_row

            missing_INT_list = []
            missing_OVA_list = []
            matchFound = False
            headerChecker = True
            counter = 0

            for i in range(Ofirst_row + 1, Olast_row + 1):
                if Osheet[debitCreditFlagColumn + str(i)].value == 'C':
                    for j in range(Ifirst_row + 1, Ilast_row + 1):
                        if (str(Osheet[ova_id_col + str(i)].value)).upper() == (
                        str(Isheet[int_id_col + str(j)].value)).upper():
                            matchFound = True
                            break
                        else:
                            matchFound = False
                    if matchFound is False:
                        missing_INT_list.append(Osheet[ova_id_col + str(i)].value)
                else:
                    continue
            print(f"Missing integrator transactions: {missing_INT_list}")
            file = pd.read_excel(match_OVA)
            count = 0
            for transaction in missing_INT_list:
                data = file[file['REMARKS2'] == transaction]
                if counter > 0:
                    headerChecker = False
                with pd.ExcelWriter(match_OVA, mode='a', engine="openpyxl", if_sheet_exists='overlay') as writer:
                    data.to_excel(writer, sheet_name='Missing Integrator Transactions', startrow=count,
                                  header=headerChecker)
                    if headerChecker is False:
                        count += 1
                    else:
                        count += 2
                    counter += 1
            Iwb.close()
            Owb.close()

            # ######################################### MISSING OVA TRANSACTIONS #########################################
            Owb = load_workbook(match_OVA)
            Osheet = Owb[Owb.sheetnames[0]]
            Owb.active = Owb[Owb.sheetnames[0]]

            Iwb = load_workbook(match_INT)
            Isheet = Iwb[Iwb.sheetnames[0]]
            Iwb.active = Iwb[Iwb.sheetnames[0]]

            Ofirst_col = Owb.active.min_column
            Olast_col = Owb.active.max_column
            Ofirst_row = Owb.active.min_row
            Olast_row = Owb.active.max_row

            Ifirst_col = Iwb.active.min_column
            Ilast_col = Iwb.active.max_column
            Ifirst_row = Iwb.active.min_row
            Ilast_row = Iwb.active.max_row

            missing_INT_list = []
            missing_OVA_list = []
            matchFound = False
            headerChecker = True
            counter = 0

            for i in range(Ifirst_row + 1, Ilast_row + 1):
                if Osheet[debitCreditFlagColumn + str(i)].value == 'C':
                    for j in range(Ofirst_row + 1, Olast_row + 1):
                        if (str(Isheet[int_id_col + str(i)].value)).upper() == (str(Osheet[ova_id_col + str(j)].value)).upper():
                            matchFound = True
                            break
                        else:
                            matchFound = False
                    if matchFound is False:
                        missing_OVA_list.append(Isheet[int_id_col + str(i)].value)
                else:
                    continue
            print(f"Missing OVA transactions: {missing_OVA_list}")
            file = pd.read_excel(match_INT)
            count = 0
            for transaction in missing_OVA_list:
                data = file[file['External Payment Request → Institution Trans ID'] == transaction]
                if counter > 0:
                    headerChecker = False
                with pd.ExcelWriter(match_OVA, mode='a', engine="openpyxl", if_sheet_exists='overlay') as writer:
                    data.to_excel(writer, sheet_name='Missing OVA Transactions', startrow=count, header=headerChecker)
                    if headerChecker is False:
                        count += 1
                    else:
                        count += 2
                    counter += 1
            Iwb.close()
            Owb.close()
            break
    else:
        fileFound = False

if fileFound is False:
    INT_VOLUME9 = 0
    INT_VALUE9 = 0
    DUPLICATES_VOLUME9 = 0
    DUPLICATES_VALUE9 = 0
#
#
OVA_VOLUME10 = 0
OVA_VALUE10 = 0
INT_VALUE10 = 0
INT_VOLUME10 = 0
DUPLICATES_VOLUME10 = 0
DUPLICATES_VALUE10 = 0
#
matchOVAFound = False
matchINTFound = False
fileFound = False
#
# ################################################## GIP OVA #################################################
OVA_VALUE11 = 0
OVA_VOLUME11 = 0

try:
    if 'slydepay_sending_' + GIPdate + '.xlsx' in dir_list and 'slydepay_sendingGhlink_' + GIPdate + '.xlsx' in dir_list:
        fileFound = True
        wb8 = load_workbook('slydepay_sending_' + GIPdate + '.xlsx')
        wb9 = load_workbook('slydepay_sendingGhlink_' + GIPdate + '.xlsx')

        sheet8 = wb8[wb8.sheetnames[0]]
        sheet9 = wb9[wb9.sheetnames[0]]

        # DEFINE MAX AND MIN COLUMNS AND ROWS
        first_col8 = wb8.active.min_column
        last_col8 = wb8.active.max_column
        first_row8 = wb8.active.min_row
        last_row8 = wb8.active.max_row

        first_col9 = wb9.active.min_column
        last_col9 = wb9.active.max_column
        first_row9 = wb9.active.min_row
        last_row9 = wb9.active.max_row

        SLYDEPAYSENDING_Volume = (last_row8 + 1) - (first_row8 + 1)

        SLYDEPAYSENDING_Sum = 0.00
        sheet = sheet8
        first_row = first_row8
        last_row = last_row8
        first_col = first_col8
        last_col = last_col8
        amountColumn = get_column("TRANSACTION_AMOUNT")
        for i in range(first_row8 + 1, last_row8 + 1):
            SLYDEPAYSENDING_Sum = SLYDEPAYSENDING_Sum + abs(float(sheet8[amountColumn + str(i)].value))

        sheet = sheet9
        first_row = first_row9
        last_row = last_row9
        first_col = first_col9
        last_col = last_col9
        amountColumn = get_column("TRANSACTION_AMOUNT")
        SLYDEPAYGHLINK_Volume = (last_row9 + 1) - (first_row9 + 1)
        SLYDEPAYGHLINK_Sum = 0.00

        for i in range(first_row9 + 1, last_row9 + 1):
            SLYDEPAYGHLINK_Sum = SLYDEPAYGHLINK_Sum + abs(float(sheet9[amountColumn + str(i)].value))

        GIP_OVA_VOLUME = SLYDEPAYGHLINK_Volume + SLYDEPAYSENDING_Volume
        OVA_VOLUME11 = GIP_OVA_VOLUME
        GIP_OVA_SUM = SLYDEPAYGHLINK_Sum + SLYDEPAYSENDING_Sum
        OVA_VALUE11 = GIP_OVA_SUM

        print('GIP_OVA_VOLUME: ' + str(GIP_OVA_VOLUME))
        print('GIP_OVA_Sum: ' + str(GIP_OVA_SUM))
        wb8.close()
        wb9.close()
        wb8.save('slydepay_sending_' + GIPdate + '- Recons.xlsx')
        wb9.save('slydepay_sendingGhlink_' + GIPdate + '.xlsx')
        match_OVA = 'slydepay_sending_' + GIPdate + '- Recons.xlsx'

        wb = load_workbook(match_OVA)
        ws1 = wb.create_sheet('Duplicates')
        ws1.title = 'Duplicates'

        wb.close()
        wb.save(match_OVA)
    else:
        fileFound = False
except FileNotFoundError:
    fileFound = False
    OVA_VOLUME11 = 0
    OVA_VALUE11 = 0
#

#
fileFound = False
# ##################################################### GIP INT #####################################################
for file in dir_list:
    if file == 'GIP Metabase' + yesterday + '.xlsx':
        fileFound = True
        wb = load_workbook('GIP Metabase' + yesterday + '.xlsx')
        sheet = wb[wb.sheetnames[0]]
        wb.active = wb[wb.sheetnames[0]]
        file = pd.read_excel('GIP Metabase' + yesterday + '.xlsx')

        # DEFINE MAX AND MIN COLUMNS AND ROWS
        first_col = wb.active.min_column
        last_col = wb.active.max_column
        first_row = wb.active.min_row
        last_row = wb.active.max_row

        GIP_INT_Volume = (last_row + 1) - (first_row + 1)
        INT_VOLUME11 = GIP_INT_Volume
        print('GIP_INT_Volume: ' + str(GIP_INT_Volume))

        GIP_INT_Sum = 0.00

        amountColumn = get_column('amount')
        if amountColumn is None:
            amountColumn = get_column('Amount')

        int_id_col = get_column('integratorTransId')
        if int_id_col == None:
            int_id_col = get_column('IntegratorTransId')

        for i in range(first_row + 1, last_row + 1):
            GIP_INT_Sum = GIP_INT_Sum + abs(float(sheet[amountColumn + str(i)].value))

        print('GIP_INT_Sum: ' + str(GIP_INT_Sum))
        INT_VALUE11 = GIP_INT_Sum
        wb.close()
        wb.save('GIP Metabase' + yesterday + '.xlsx')

        wb = load_workbook('GIP Metabase' + yesterday + '.xlsx')
        sheet = wb[wb.sheetnames[0]]
        wb.active = wb[wb.sheetnames[0]]
        file = pd.read_excel('GIP Metabase' + yesterday + '.xlsx')
        # ------------------------------------------------ Duplicates -------------------------------------------------
        list = []
        duplicate = []
        count = 1
        counter = 0
        headerChecker = True
        duplicates_value_sum = 0.00
        for id in range(first_row + 1, last_row + 1):
            if sheet[int_id_col + str(id)].value not in list:
                list.append(sheet[int_id_col + str(id)].value)
            else:
                try:
                    data = file[file['integratorTransId'] == (sheet[int_id_col + str(id)].value)]
                except:
                    data = file[file['IntegratorTransId'] == (sheet[int_id_col + str(id)].value)]
                if sheet[int_id_col + str(id)].value not in duplicate:
                    duplicate.append(sheet[int_id_col + str(id)].value)
                    if counter > 0:
                        headerChecker = False
                    with pd.ExcelWriter(match_OVA, mode='a', engine="openpyxl", if_sheet_exists='overlay') as writer:
                        data.to_excel(writer, sheet_name='Duplicates', startrow=count, header=headerChecker)
                        if headerChecker is False:
                            count += 2
                        else:
                            count += 3
                        counter += 1
                else:
                    continue
        wb.close()

        wb = load_workbook(match_OVA)
        wb.active = wb['Duplicates']
        dup_sheet = wb['Duplicates']
        sheet = dup_sheet

        first_col = wb.active.min_column
        last_col = wb.active.max_column
        first_row = wb.active.min_row
        last_row = wb.active.max_row

        if counter > 0:
            dup_col = column_index_from_string(amountColumn) + 1
            for number in range(first_row + 2, last_row + 1, 2):
                duplicates_value_sum = duplicates_value_sum + abs(
                    float(dup_sheet[get_column_letter(dup_col) + str(number)].value))
                count += 3

            print(f"Number of duplicates: {counter}")
            dup_sheet[get_column_letter(dup_col) + str(last_row + 4)].value = str(round(duplicates_value_sum, 2))
            dup_sheet[get_column_letter(dup_col) + str(last_row + 5)].value = str(counter)

            print(f"Duplicates: {duplicate}")
            DUPLICATES_VOLUME11 = counter
            DUPLICATES_VALUE11 = duplicates_value_sum

            wb.close()
            wb.save(match_OVA)
            break
        else:
            DUPLICATES_VOLUME11 = 0
            DUPLICATES_VALUE11 = 0
    else:
        fileFound = False

if fileFound is False:
    INT_VOLUME11 = 0
    INT_VALUE11 = 0
    DUPLICATES_VOLUME11 = 0
    DUPLICATES_VALUE11 = 0

#
#
#
matchOVAFound = False
matchINTFound = False
fileFound = False
# ############################################# BB MIG_OVA ###########################################################
OVA_VALUE12 = 0
OVA_VOLUME12 = 0
try:
    if ('MIGS 08' + yesterday + '.xlsx') and ('MIGS 09' + yesterday + '.xlsx') in dir_list:
        fileFound = True
        wb8 = load_workbook('MIGS 08' + yesterday + '.xlsx')
        wb9 = load_workbook('MIGS 09' + yesterday + '.xlsx')

        sheet8 = wb8[wb8.sheetnames[0]]
        sheet9 = wb9[wb9.sheetnames[0]]

        # DEFINE MAX AND MIN COLUMNS AND ROWS
        first_col8 = wb8.active.min_column
        last_col8 = wb8.active.max_column
        first_row8 = wb8.active.min_row
        last_row8 = wb8.active.max_row

        first_col9 = wb9.active.min_column
        last_col9 = wb9.active.max_column
        first_row9 = wb9.active.min_row
        last_row9 = wb9.active.max_row

        BB_MIG8_Volume = (last_row8 + 1) - (first_row8 + 4)

        BB_MIG8_Sum = 0.00

        sheet = sheet8
        first_row = first_row8
        last_row = last_row8
        first_col = first_col8
        last_col = last_col8
        amountColumn = get_column("Amount")
        for i in range(first_row8 + 4, last_row8 + 1):
            BB_MIG8_Sum = BB_MIG8_Sum + abs(float(sheet8[amountColumn + str(i)].value))

        sheet = sheet9
        first_row = first_row9
        last_row = last_row9
        first_col = first_col9
        last_col = last_col9
        amountColumn = get_column("Amount")
        BB_MIG9_Volume = (last_row9 + 1) - (first_row9 + 4)
        BB_MIG9_Sum = 0.00

        for i in range(first_row9 + 4, last_row9 + 1):
            BB_MIG9_Sum = BB_MIG9_Sum + abs(float(sheet9[amountColumn + str(i)].value))

        BB_MIG_OVA_VOLUME = BB_MIG9_Volume + BB_MIG8_Volume
        OVA_VOLUME12 = BB_MIG_OVA_VOLUME
        BB_MIG_SUM = BB_MIG9_Sum + BB_MIG8_Sum
        OVA_VALUE12 = BB_MIG_SUM

        print('BB_MIG_OVA_VOLUME: ' + str(BB_MIG_OVA_VOLUME))
        print('BB_MIG_OVA_Sum: ' + str(BB_MIG_SUM))
        wb8.close()
        wb9.close()
        wb8.save('MIGS 08' + yesterday + '- Recons.xlsx')
        wb9.save('MIGS 09' + yesterday + '.xlsx')
        match_OVA = 'MIGS 08' + yesterday + '- Recons.xlsx'

        wb = load_workbook(match_OVA)
        ws1 = wb.create_sheet("Duplicates")
        ws1.title = 'Duplicates'
        wb.close()
        wb.save(match_OVA)

    elif ('MIGS 08' + yesterday + '.xlsx') in dir_list and ('MIGS 09' + yesterday + '.xlsx') not in dir_list:
        fileFound = True
        wb = load_workbook('MIGS 08' + yesterday + '.xlsx')
        sheet = wb[wb.sheetnames[0]]

        # DEFINE MAX AND MIN COLUMNS AND ROWS
        first_col = wb.active.min_column
        last_col = wb.active.max_column
        first_row = wb.active.min_row
        last_row = wb.active.max_row

        BB_MIG_OVA_Volume = (last_row + 1) - (first_row + 4)
        BB_MIG_Sum = 0.00

        amountColumn = get_column("Amount")
        for i in range(first_row + 4, last_row + 1):
            BB_MIG_Sum = BB_MIG_Sum + abs(float(sheet[amountColumn + str(i)].value))
        BB_MIG_Volume = (last_row + 1) - (first_row + 4)

        OVA_VOLUME12 = BB_MIG_OVA_Volume
        OVA_VALUE12 = BB_MIG_Sum

        print('BB_MIG_OVA_VOLUME: ' + str(BB_MIG_OVA_Volume))
        print('BB_MIG_OVA_Sum: ' + str(BB_MIG_Sum))

        wb.close()
        wb.save('MIGS 08' + yesterday + '- Recons.xlsx')
        match_OVA = 'MIGS 08' + yesterday + '- Recons.xlsx'

        wb = load_workbook(match_OVA)
        ws1 = wb.create_sheet("Duplicates")
        ws1.title = 'Duplicates'
        wb.close()
        wb.save(match_OVA)


    elif ('MIGS 08' + yesterday + '.xlsx') not in dir_list and ('MIGS 09' + yesterday + '.xlsx') in dir_list:
        fileFound = True
        wb = load_workbook('MIGS 09' + yesterday + '.xlsx')
        sheet = wb[wb.sheetnames[0]]

        # DEFINE MAX AND MIN COLUMNS AND ROWS
        first_col = wb.active.min_column
        last_col = wb.active.max_column
        first_row = wb.active.min_row
        last_row = wb.active.max_row

        BB_MIG_OVA_Volume = (last_row + 1) - (first_row + 4)
        BB_MIG_Sum = 0.00

        amountColumn = get_column("Amount")
        for i in range(first_row + 4, last_row + 1):
            BB_MIG_Sum = BB_MIG_Sum + abs(float(sheet[amountColumn + str(i)].value))

        OVA_VOLUME12 = BB_MIG_OVA_Volume
        OVA_VALUE12 = BB_MIG_Sum

        print('BB_MIG_OVA_VOLUME: ' + str(BB_MIG_OVA_Volume))
        print('BB_MIG_OVA_Sum: ' + str(BB_MIG_Sum))

        wb.close()
        wb.save('MIGS 09' + yesterday + '- Recons.xlsx')
        match_OVA = 'MIGS 09' + yesterday + '- Recons.xlsx'

        wb = load_workbook(match_OVA)
        ws1 = wb.create_sheet("Duplicates")
        ws1.title = 'Duplicates'
        wb.close()
        wb.save(match_OVA)


except FileNotFoundError:
    fileFound = False
    OVA_VOLUME12 = 0
    OVA_VALUE12 = 0

#

#
fileFound = False
# ############################################### BB MIG_INT #########################################################
for file in dir_list:
    if file == 'MiGS_trn' + yesterday + '.xlsx':
        fileFound = True
        wb = load_workbook('MiGS_trn' + yesterday + '.xlsx')
        wb.active = wb[wb.sheetnames[0]]
        sheet = wb[wb.sheetnames[0]]
        file = pd.read_excel('MiGS_trn' + yesterday + '.xlsx')
        match_INT = 'MiGS_trn' + yesterday + '.xlsx'

        # DEFINE MAX AND MIN COLUMNS AND ROWS
        first_col = wb.active.min_column
        last_col = wb.active.max_column
        first_row = wb.active.min_row
        last_row = wb.active.max_row

        MIGS_INT_Sum = 0.00
        MIGS_INT_VOLUME = 0

        amountColumn = get_column('Amount')
        statusFlagColumn = get_column('Status')
        int_id_col = get_column('Transaction Id')
        for i in range(first_row + 1, last_row + 1):
            if sheet[statusFlagColumn + str(i)].value == 'CONFIRMED':
                MIGS_INT_VOLUME = MIGS_INT_VOLUME + 1
                MIGS_INT_Sum = MIGS_INT_Sum + abs(float(sheet[amountColumn + str(i)].value))

        print('MIGS_INT_VOLUME: ' + str(MIGS_INT_VOLUME))
        INT_VOLUME12 = MIGS_INT_VOLUME
        print('MIGS_INT_Sum: ' + str(MIGS_INT_Sum))
        INT_VALUE12 = MIGS_INT_Sum
        wb.close()
        wb.save('MiGS_trn' + yesterday + '.xlsx')
        # ############################################ Duplicates #######################################################
        wb = load_workbook('MiGS_trn' + yesterday + '.xlsx')
        sheet = wb[wb.sheetnames[0]]
        wb.active = wb[wb.sheetnames[0]]
        list = []
        duplicate = []
        count = 1
        counter = 0
        headerChecker = True
        duplicates_value_sum = 0.00
        for id in range(first_row + 1, last_row + 1):
            if sheet[int_id_col + str(id)].value not in list:
                list.append(sheet[int_id_col + str(id)].value)
            else:
                if sheet[int_id_col + str(id)].value not in duplicate:
                    data = file[file['Transaction Id'] == (sheet[int_id_col + str(id)].value)]
                    duplicate.append(sheet[int_id_col + str(id)].value)
                    if counter > 0:
                        headerChecker = False
                    with pd.ExcelWriter(match_OVA, mode='a', engine="openpyxl", if_sheet_exists='overlay') as writer:
                        data.to_excel(writer, sheet_name='Duplicates', startrow=count, header=headerChecker)
                        if headerChecker is False:
                            count += 2
                        else:
                            count += 3
                        counter += 1
                else:
                    continue
        wb.close()
        wb.save('MiGS_trn' + yesterday + '.xlsx')

        wb = load_workbook(match_OVA)
        wb.active = wb['Duplicates']
        dup_sheet = wb['Duplicates']
        sheet = dup_sheet

        first_col = wb.active.min_column
        last_col = wb.active.max_column
        first_row = wb.active.min_row
        last_row = wb.active.max_row

        dup_col = column_index_from_string(amountColumn) + 1
        for number in range(first_row + 2, last_row + 1, 2):
            duplicates_value_sum = duplicates_value_sum + abs(
                float(dup_sheet[get_column_letter(dup_col) + str(number)].value))
            count += 3

        print(f"Number of duplicates: {counter}")
        dup_sheet[get_column_letter(dup_col) + str(last_row + 4)].value = str(round(duplicates_value_sum, 2))
        dup_sheet[get_column_letter(dup_col) + str(last_row + 5)].value = str(counter)

        print(f"Duplicates: {duplicate}")

        DUPLICATES_VOLUME12 = counter
        DUPLICATES_VALUE12 = duplicates_value_sum

        wb.close()
        wb.save(match_OVA)
        break

if fileFound is False:
    INT_VOLUME12 = 0
    INT_VALUE12 = 0
    DUPLICATES_VOLUME12 = 0
    DUPLICATES_VALUE12 = 0
#
#
matchOVAFound = False
matchINTFound = False
fileFound = False
# ################################################# MPGS OVA ##########################################################
try:
    if "MPGS" +yesterday +".xlsx" not in dir_list and "Quipu" + yesterday +".xlsx" in dir_list:
        fileFound = True
        matchOVAFound = True
        wb = load_workbook('Quipu' + yesterday + '.xlsx')
        sheet = wb[wb.sheetnames[0]]
        wb.active = wb[wb.sheetnames[0]]

        # DEFINE MAX AND MIN COLUMNS AND ROWS
        first_col = wb.active.min_column
        last_col = wb.active.max_column
        first_row = wb.active.min_row
        last_row = wb.active.max_row

        MPGS_OVA_Volume = 0
        ova_ids = []
        status_id_col = get_column("Status")

        MPGS_OVA_Sum = 0.00

        amountColumn = get_column('Amount')
        if amountColumn is None:
            amountColumn = get_column("amount")

        ova_id_col = get_column('Order Code')
        for i in range(first_row + 1, last_row + 1):
            if sheet[status_id_col + str(i)].value == "SUCCESS":
                MPGS_OVA_Volume = MPGS_OVA_Volume + 1
                MPGS_OVA_Sum = MPGS_OVA_Sum + abs(float(sheet[amountColumn + str(i)].value))
                ova_ids.append(sheet[ova_id_col + str(i)].value)

        OVA_VOLUME13 = MPGS_OVA_Volume

        print('MPGS_OVA_Sum: ' + str(MPGS_OVA_Sum))
        print('MPGS_OVA_Volume: ' + str(MPGS_OVA_Volume))
        OVA_VALUE13 = MPGS_OVA_Sum
        wb.close()
        wb.save('MPGS' + yesterday + '- Recons.xlsx')
        match_OVA = 'MPGS' + yesterday + '- Recons.xlsx'

        wb = load_workbook(match_OVA)
        ws1 = wb.create_sheet('Duplicates')
        ws1.title = 'Duplicates'
        ws1 = wb.create_sheet('Missing Integrator Transactions')
        ws1.title = 'Missing Integrator Transactions'
        ws1 = wb.create_sheet('Missing OVA Transactions')
        ws1.title = 'Missing OVA Transactions'

        wb.close()
        wb.save(match_OVA)

    elif ("MPGS" +yesterday +".xlsx" in dir_list) and ("Quipu" + yesterday +".xlsx" not in dir_list):
        fileFound = True
        matchOVAFound = True
        wb = load_workbook('MPGS' + yesterday + '.xlsx')
        sheet = wb[wb.sheetnames[0]]
        wb.active = wb[wb.sheetnames[0]]

        # DEFINE MAX AND MIN COLUMNS AND ROWS
        first_col = wb.active.min_column
        last_col = wb.active.max_column
        first_row = wb.active.min_row
        last_row = wb.active.max_row

        MPGS_OVA_Volume = (last_row + 1) - (first_row + 1)
        OVA_VOLUME13 = MPGS_OVA_Volume
        ova_ids = []
        print('MPGS_OVA_Volume: ' + str(MPGS_OVA_Volume))

        MPGS_OVA_Sum = 0.00

        amountColumn = get_column('Transaction Amount (amount only)')
        if amountColumn is None:
            amountColumn = get_column("Order Amount (amount only)")

        ova_id_col = get_column('Order ID')
        for i in range(first_row + 1, last_row + 1):
            MPGS_OVA_Sum = MPGS_OVA_Sum + abs(float(sheet[amountColumn + str(i)].value))
            ova_ids.append(sheet[ova_id_col + str(i)].value)

        print('MPGS_OVA_Sum: ' + str(MPGS_OVA_Sum))
        OVA_VALUE13 = MPGS_OVA_Sum
        wb.close()
        wb.save('MPGS' + yesterday + '- Recons.xlsx')
        match_OVA = 'MPGS' + yesterday + '- Recons.xlsx'

        wb = load_workbook(match_OVA)
        ws1 = wb.create_sheet('Duplicates')
        ws1.title = 'Duplicates'
        ws1 = wb.create_sheet('Missing Integrator Transactions')
        ws1.title = 'Missing Integrator Transactions'
        ws1 = wb.create_sheet('Missing OVA Transactions')
        ws1.title = 'Missing OVA Transactions'

        wb.close()
        wb.save(match_OVA)
    elif "MPGS" + yesterday +".xlsx" in dir_list and "Quipu" + yesterday + ".xlsx" in dir_list:
        fileFound = True
        matchOVAFound = True
        wb = load_workbook('Quipu' + yesterday + '.xlsx')
        sheet = wb[wb.sheetnames[0]]
        wb.active = wb[wb.sheetnames[0]]

        # DEFINE MAX AND MIN COLUMNS AND ROWS
        first_col = wb.active.min_column
        last_col = wb.active.max_column
        first_row = wb.active.min_row
        last_row = wb.active.max_row

        MPGS_OVA_Volume = 0
        ova_ids = []
        status_id_col = get_column("Status")

        MPGS_OVA_Sum = 0.00

        amountColumn = get_column('Amount')
        if amountColumn is None:
            amountColumn = get_column("amount")

        ova_id_col = get_column('Order Code')
        for i in range(first_row + 1, last_row + 1):
            if sheet[status_id_col + str(i)].value == "SUCCESS":
                MPGS_OVA_Volume = MPGS_OVA_Volume + 1
                MPGS_OVA_Sum = MPGS_OVA_Sum + abs(float(sheet[amountColumn + str(i)].value))
                ova_ids.append(sheet[ova_id_col + str(i)].value)

        wb.close()
        wb.save('MPGS' + yesterday + '- Recons.xlsx')
        match_OVA = 'MPGS' + yesterday + '- Recons.xlsx'
        wb.close()
        wb.save(match_OVA)

        wb = load_workbook('MPGS' + yesterday + '.xlsx')
        sheet = wb[wb.sheetnames[0]]
        wb.active = wb[wb.sheetnames[0]]

        # DEFINE MAX AND MIN COLUMNS AND ROWS
        first_col = wb.active.min_column
        last_col = wb.active.max_column
        first_row = wb.active.min_row
        last_row = wb.active.max_row

        MPGS_OVA_Volume = MPGS_OVA_Volume + ((last_row + 1) - (first_row + 1))
        OVA_VOLUME13 = MPGS_OVA_Volume
        print('MPGS_OVA_Volume: ' + str(MPGS_OVA_Volume))

        MPGS_OVA_Sum = 0.00

        amountColumn = get_column('Transaction Amount (amount only)')
        if amountColumn is None:
            amountColumn = get_column("Order Amount (amount only)")

        ova_id_col = get_column('Order ID')
        for i in range(first_row + 1, last_row + 1):
            MPGS_OVA_Sum = MPGS_OVA_Sum + abs(float(sheet[amountColumn + str(i)].value))
            ova_ids.append(sheet[ova_id_col + str(i)].value)

        print('MPGS_OVA_Sum: ' + str(MPGS_OVA_Sum))
        print('MPGS_OVA_Volume: ' + str(MPGS_OVA_Volume))
        OVA_VALUE13 = MPGS_OVA_Sum
        wb.close()
        wb.save('MPGS' + yesterday + '- Recons.xlsx')
        match_OVA = 'MPGS' + yesterday + '- Recons.xlsx'

        wb = load_workbook(match_OVA)
        ws1 = wb.create_sheet('Duplicates')
        ws1.title = 'Duplicates'
        ws1 = wb.create_sheet('Missing Integrator Transactions')
        ws1.title = 'Missing Integrator Transactions'
        ws1 = wb.create_sheet('Missing OVA Transactions')
        ws1.title = 'Missing OVA Transactions'

        wb.close()
        wb.save(match_OVA)

except FileNotFoundError:
    print("file not found")
    OVA_VALUE13 = 0.00
    OVA_VOLUME13 = 0

#
#
#
fileFound = False
matchINTFound = False
#
INT_VALUE13 = 0.00
INT_VOLUME13 = 0
# ############################################################ MPGS INT ######################################################################################
DUPLICATES_VOLUME13 = 0
DUPLICATES_VALUE13 = 0
try:
    if "MPGS KC" + yesterday + ".xlsx" in dir_list and 'MPGS_trn' + yesterday + '.xlsx' in dir_list:
        MPGS_INT_Volume = 0
        MPGS_INT_Sum = 0.00
        fileFound = True
        matchINTFound = True
        wb = load_workbook('MPGS KC' + yesterday + '.xlsx')
        sheet = wb[wb.sheetnames[0]]
        wb.active = wb[wb.sheetnames[0]]
        file = pd.read_excel('MPGS KC' + yesterday + '.xlsx')
        match_INT = 'MPGS KC' + yesterday + '.xlsx'

        # DEFINE MAX AND MIN COLUMNS AND ROWS
        first_col = wb.active.min_column
        last_col = wb.active.max_column
        first_row = wb.active.min_row
        last_row = wb.active.max_row

        MPGS_INT_Volume = (last_row + 1) - (first_row + 1)

        MPGS_INT_Sum = 0.00

        amountColumn = get_column('Real Total')
        int_id_col = get_column("External Transaction Reference")
        ids = []
        for i in range(first_row + 1, last_row + 1):
            MPGS_INT_Sum = MPGS_INT_Sum + abs(float(sheet[amountColumn + str(i)].value))
            ids.append(str(sheet[int_id_col + str(i)].value))

        wb.close()
        wb.save('MPGS KC' + yesterday + '.xlsx')

        wb = load_workbook('MPGS_trn' + yesterday + '.xlsx')
        sheet = wb[wb.sheetnames[0]]
        wb.active = wb[wb.sheetnames[0]]
        amountColumn = get_column("Amount")
        statusFlagColumn = get_column("Status")
        int_id_col = get_column("Transaction Id")

        first_col = wb.active.min_column
        last_col = wb.active.max_column
        first_row = wb.active.min_row
        last_row = wb.active.max_row

        for i in range(first_row + 1, last_row + 1):
            if sheet[statusFlagColumn + str(i)].value == "CONFIRMED":
                MPGS_INT_Sum = MPGS_INT_Sum + abs(float(sheet[amountColumn + str(i)].value))
                MPGS_INT_Volume += 1
                ids.append(str(sheet[int_id_col + str(i)].value))

        print('MPGS_INT_Volume: ' + str(MPGS_INT_Volume))
        print('MPGS_INT_Sum: ' + str(MPGS_INT_Sum))
        INT_VALUE13 = MPGS_INT_Sum
        INT_VOLUME13 = MPGS_INT_Volume
        wb.close()
        wb.save('MPGS_trn' + yesterday + '.xlsx')
        # -------------------------------------------- DUPLICATES --------------------------------------------------------------------
        wb = load_workbook('MPGS KC' + yesterday + '.xlsx')
        Owb = load_workbook(match_OVA)
        sheet = wb[wb.sheetnames[0]]
        wb.active = wb[wb.sheetnames[0]]

        first_col = wb.active.min_column
        last_col = wb.active.max_column
        first_row = wb.active.min_row
        last_row = wb.active.max_row

        amountColumn = get_column('Real Total')
        int_id_col = get_column("External Transaction Reference")

        list = []
        duplicate = []
        count = 1
        headerChecker = True
        counter = 0
        duplicates_value_sum = 0.00

        for i in range(first_row + 1, last_row + 1):
            if sheet[int_id_col + str(i)].value not in list:
                list.append(sheet[int_id_col + str(i)].value)
            else:
                if sheet[int_id_col + str(i)].value not in duplicate:
                    data = file[file['External Transaction Reference'] == sheet[int_id_col + str(i)].value]
                    duplicate.append(sheet[int_id_col + str(i)].value)
                    if counter > 0:
                        headerChecker = False
                    with pd.ExcelWriter(match_OVA, mode='a', engine="openpyxl", if_sheet_exists='overlay') as writer:
                        data.to_excel(writer, sheet_name='Duplicates', startrow=count, header=headerChecker)
                        if headerChecker is False:
                            count += 2
                        else:
                            count += 3
                        counter += 1
                else:
                    continue

        wb.close()
        wb.save('MPGS KC' + yesterday + '.xlsx')
        Owb.close()
        Owb.save(match_OVA)

        wb = load_workbook('MPGS_trn' + yesterday + '.xlsx')
        Owb = load_workbook(match_OVA)
        sheet = wb[wb.sheetnames[0]]
        wb.active = wb[wb.sheetnames[0]]

        first_col = wb.active.min_column
        last_col = wb.active.max_column
        first_row = wb.active.min_row
        last_row = wb.active.max_row
        int_id_col = get_column("Receipt No.")
        list = []
        duplicate = []
        count = 1
        duplicates_value_sum = 0.00

        for i in range(first_row + 1, last_row + 1):
            if sheet[statusFlagColumn + str(i)].value == "CONFIRMED":
                if sheet[int_id_col + str(i)].value not in list:
                    list.append(sheet[int_id_col + str(i)].value)
                else:
                    data = file[file['Transaction Id'] == (sheet[int_id_col + str(i)].value)]
                    duplicate.append(sheet[int_id_col + str(i)].value)
                    if counter > 0:
                        headerChecker = False
                    with pd.ExcelWriter(match_OVA, mode='a', engine="openpyxl", if_sheet_exists='overlay') as writer:
                        data.to_excel(writer, sheet_name='Duplicates', startrow=count, header=headerChecker)
                        if headerChecker is False:
                            count += 2
                        else:
                            count += 3
                        counter += 1

        wb.close()
        wb.save('MPGS_trn' + yesterday + '.xlsx')
        Owb.close()
        Owb.save(match_OVA)

        wb = load_workbook(match_OVA)
        wb.active = wb['Duplicates']
        dup_sheet = wb['Duplicates']

        first_col = wb.active.min_column
        last_col = wb.active.max_column
        first_row = wb.active.min_row
        last_row = wb.active.max_row

        dup_col = column_index_from_string(amountColumn) + 1
        for number in range(first_row + 2, last_row + 1, 2):
            duplicates_value_sum = duplicates_value_sum + dup_sheet[get_column_letter(dup_col) + str(number)].value
            count += 3

        print(f"Number of duplicates: {counter}")
        dup_sheet[get_column_letter(dup_col) + str(last_row + 4)].value = str(round(duplicates_value_sum, 2))
        dup_sheet[get_column_letter(dup_col) + str(last_row + 5)].value = str(counter)

        print(f"Duplicates: {duplicate}")

        DUPLICATES_VOLUME13 = counter
        DUPLICATES_VALUE13 = duplicates_value_sum
        wb.close()
        wb.save(match_OVA)

        if matchOVAFound is True and matchINTFound is True:
            #  ###################################################### MISSING INT TRANSACTIONS ##############################################################

            Owb = load_workbook(match_OVA)
            Osheet = Owb[Owb.sheetnames[0]]
            Owb.active = Owb[Owb.sheetnames[0]]
            file = pd.read_excel(match_OVA)

            Iwb = load_workbook(match_INT)
            Isheet = Iwb[Iwb.sheetnames[0]]
            Iwb.active = Iwb[Iwb.sheetnames[0]]

            Ofirst_col = Owb.active.min_column
            Olast_col = Owb.active.max_column
            Ofirst_row = Owb.active.min_row
            Olast_row = Owb.active.max_row

            Ifirst_col = Iwb.active.min_column
            Ilast_col = Iwb.active.max_column
            Ifirst_row = Iwb.active.min_row
            Ilast_row = Iwb.active.max_row

            missing_INT_list = []
            missing_OVA_list = []
            matchFound = False
            headerChecker = True
            counter = 0
            for i in range(Ofirst_row + 1, Olast_row + 1):
                for j in ids:
                    if str(Osheet[ova_id_col + str(i)].value) in str(j):
                        matchFound = True
                        break
                    else:
                        matchFound = False
                if matchFound is False:
                    missing_INT_list.append(Osheet[ova_id_col + str(i)].value)
            print(f"Missing integrator transactions: {missing_INT_list}")
            ws2 = Iwb.create_sheet('Missing Integrator Transactions')
            ws2.title = 'Missing Integrator Transactions'
            count = 0

            for transaction in missing_INT_list:
                data = file[file['Order ID'] == transaction]
                if counter > 0:
                    headerChecker = False
                with pd.ExcelWriter(match_OVA, mode='a', engine="openpyxl", if_sheet_exists='overlay') as writer:
                    data.to_excel(writer, sheet_name='Missing Integrator Transactions', startrow=count,
                                  header=headerChecker)
                    if headerChecker is False:
                        count += 1
                    else:
                        count += 2
                    counter += 1
            Iwb.close()
            Owb.close()

            # second check
            match_INT = 'MPGS_trn' + yesterday + '.xlsx'
            Owb = load_workbook(match_OVA)
            Osheet = Owb[Owb.sheetnames[0]]
            Owb.active = Owb[Owb.sheetnames[0]]
            file = pd.read_excel(match_OVA)

            Iwb = load_workbook(match_INT)
            Isheet = Iwb[Iwb.sheetnames[0]]
            Iwb.active = Iwb[Iwb.sheetnames[0]]

            Ofirst_col = Owb.active.min_column
            Olast_col = Owb.active.max_column
            Ofirst_row = Owb.active.min_row
            Olast_row = Owb.active.max_row

            Ifirst_col = Iwb.active.min_column
            Ilast_col = Iwb.active.max_column
            Ifirst_row = Iwb.active.min_row
            Ilast_row = Iwb.active.max_row

            missing_INT_list = []
            missing_OVA_list = []
            matchFound = False
            headerChecker = True
            counter = 0
            for i in range(Ofirst_row + 1, Olast_row + 1):
                for j in ids:
                    if str(Osheet[ova_id_col + str(i)].value) in str(j):
                        matchFound = True
                        break
                    else:
                        matchFound = False
                if matchFound is False:
                    missing_INT_list.append(Osheet[ova_id_col + str(i)].value)
            print(f"Missing integrator transactions: {missing_INT_list}")
            ws2 = Iwb.create_sheet('Missing Integrator Transactions')
            ws2.title = 'Missing Integrator Transactions'
            count = 0

            for transaction in missing_INT_list:
                data = file[file['Order ID'] == transaction]
                if counter > 0:
                    headerChecker = False
                with pd.ExcelWriter(match_OVA, mode='a', engine="openpyxl", if_sheet_exists='overlay') as writer:
                    data.to_excel(writer, sheet_name='Missing Integrator Transactions', startrow=count,
                                  header=headerChecker)
                    if headerChecker is False:
                        count += 1
                    else:
                        count += 2
                    counter += 1
            Iwb.close()
            Owb.close()

            # ######################################### MISSING OVA TRANSACTIONS ##############################################
            match_INT = 'MPGS KC' + yesterday + '.xlsx'
            Owb = load_workbook(match_OVA)
            Osheet = Owb[Owb.sheetnames[0]]
            Owb.active = Owb[Owb.sheetnames[0]]

            Iwb = load_workbook(match_INT)
            file = pd.read_excel(match_INT)
            Isheet = Iwb[Iwb.sheetnames[0]]
            Iwb.active = Iwb[Iwb.sheetnames[0]]

            Ofirst_col = Owb.active.min_column
            Olast_col = Owb.active.max_column
            Ofirst_row = Owb.active.min_row
            Olast_row = Owb.active.max_row

            Ifirst_col = Iwb.active.min_column
            Ilast_col = Iwb.active.max_column
            Ifirst_row = Iwb.active.min_row
            Ilast_row = Iwb.active.max_row

            missing_INT_list = []
            missing_OVA_list = []
            matchFound = False
            headerChecker = True
            counter = 0

            for i in ids:
                for j in range(Ofirst_row + 1, Olast_row + 1):
                    if str(Osheet[ova_id_col + str(j)].value) in str(i):
                        matchFound = True
                        break
                    else:
                        matchFound = False
                if matchFound is False:
                    missing_OVA_list.append(str(i))
            print(f"Missing OVA transactions: {missing_OVA_list}")
            count = 0
            for transaction in missing_OVA_list:
                data = file[file['External Transaction Reference'] == transaction]
                if counter > 0:
                    headerChecker = False
                with pd.ExcelWriter(match_OVA, mode='a', engine="openpyxl", if_sheet_exists='overlay') as writer:
                    data.to_excel(writer, sheet_name='Missing OVA Transactions', startrow=count, header=headerChecker)
                    if headerChecker is False:
                        count += 1
                    else:
                        count += 2
                    counter += 1
            Iwb.close()
            Owb.close()

            # second check

            Owb = load_workbook(match_OVA)
            Osheet = Owb[Owb.sheetnames[0]]
            Owb.active = Owb[Owb.sheetnames[0]]

            match_INT = 'MPGS_trn' + yesterday + '.xlsx'
            Iwb = load_workbook(match_INT)
            file = pd.read_excel(match_INT)
            Isheet = Iwb[Iwb.sheetnames[0]]
            Iwb.active = Iwb[Iwb.sheetnames[0]]

            Ofirst_col = Owb.active.min_column
            Olast_col = Owb.active.max_column
            Ofirst_row = Owb.active.min_row
            Olast_row = Owb.active.max_row

            Ifirst_col = Iwb.active.min_column
            Ilast_col = Iwb.active.max_column
            Ifirst_row = Iwb.active.min_row
            Ilast_row = Iwb.active.max_row

            missing_INT_list = []
            missing_OVA_list = []
            matchFound = False
            headerChecker = True
            counter = 0

            for i in ids:
                for j in range(Ofirst_row + 1, Olast_row + 1):
                    if str(Osheet[ova_id_col + str(j)].value) in str(i):
                        matchFound = True
                        break
                    else:
                        matchFound = False
                if matchFound is False:
                    missing_OVA_list.append(str(i))
            print(f"Missing OVA transactions: {missing_OVA_list}")
            count = 0
            for transaction in missing_OVA_list:
                data = file[file['Transaction Id'] == transaction]
                if counter > 0:
                    headerChecker = False
                with pd.ExcelWriter(match_OVA, mode='a', engine="openpyxl", if_sheet_exists='overlay') as writer:
                    data.to_excel(writer, sheet_name='Missing OVA Transactions', startrow=count, header=headerChecker)
                    if headerChecker is False:
                        count += 1
                    else:
                        count += 2
                    counter += 1
            Iwb.close()
            Owb.close()

    elif ("MPGS KC" + yesterday + ".xlsx" in dir_list) and ('MPGS_trn' + yesterday + '.xlsx' not in dir_list):
        MPGS_INT_Volume = 0
        MPGS_INT_Sum = 0.00
        fileFound = True
        matchINTFound = True
        wb = load_workbook('MPGS KC' + yesterday + '.xlsx')
        sheet = wb[wb.sheetnames[0]]
        wb.active = wb[wb.sheetnames[0]]
        file = pd.read_excel('MPGS KC' + yesterday + '.xlsx')
        match_INT = 'MPGS KC' + yesterday + '.xlsx'

        # DEFINE MAX AND MIN COLUMNS AND ROWS
        first_col = wb.active.min_column
        last_col = wb.active.max_column
        first_row = wb.active.min_row
        last_row = wb.active.max_row

        MPGS_INT_Volume = (last_row + 1) - (first_row + 1)

        MPGS_INT_Sum = 0.00

        amountColumn = get_column('Real Total')
        int_id_col = get_column("External Transaction Reference")
        ids = []
        for i in range(first_row + 1, last_row + 1):
            MPGS_INT_Sum = MPGS_INT_Sum + float(sheet[amountColumn + str(i)].value)
            ids.append(str(sheet[int_id_col + str(i)].value))

        INT_VALUE13 = MPGS_INT_Sum
        INT_VOLUME13 = MPGS_INT_Volume

        wb.close()
        wb.save(match_INT)

        # -------------------------------------------- DUPLICATES --------------------------------------------------------------------
        wb = load_workbook('MPGS KC' + yesterday + '.xlsx')
        Owb = load_workbook(match_OVA)
        sheet = wb[wb.sheetnames[0]]
        wb.active = wb[wb.sheetnames[0]]

        first_col = wb.active.min_column
        last_col = wb.active.max_column
        first_row = wb.active.min_row
        last_row = wb.active.max_row

        amountColumn = get_column('Real Total')
        int_id_col = get_column("External Transaction Reference")

        list = []
        duplicate = []
        count = 1
        headerChecker = True
        counter = 0
        duplicates_value_sum = 0.00

        for i in range(first_row + 1, last_row + 1):
            if sheet[int_id_col + str(i)].value not in list:
                list.append(sheet[int_id_col + str(i)].value)
            else:
                if sheet[int_id_col + str(i)].value not in duplicate:
                    data = file[file['External Transaction Reference'] == sheet[int_id_col + str(i)].value]
                    duplicate.append(sheet[int_id_col + str(i)].value)
                    if counter > 0:
                        headerChecker = False
                    with pd.ExcelWriter(match_OVA, mode='a', engine="openpyxl", if_sheet_exists='overlay') as writer:
                        data.to_excel(writer, sheet_name='Duplicates', startrow=count, header=headerChecker)
                        if headerChecker is False:
                            count += 2
                        else:
                            count += 3
                        counter += 1
                else:
                    continue

        wb.close()
        wb.save(match_INT)
        Owb.close()
        Owb.save(match_OVA)

        wb = load_workbook(match_OVA)
        wb.active = wb['Duplicates']
        dup_sheet = wb['Duplicates']
        sheet = dup_sheet

        first_col = wb.active.min_column
        last_col = wb.active.max_column
        first_row = wb.active.min_row
        last_row = wb.active.max_row

        dup_col = column_index_from_string(amountColumn) + 1
        for number in range(first_row + 2, last_row + 1, 2):
            duplicates_value_sum = duplicates_value_sum + dup_sheet[get_column_letter(dup_col) + str(number)].value
            count += 3

        print(f"Number of duplicates: {counter}")
        dup_sheet[get_column_letter(dup_col) + str(last_row + 4)].value = str(round(duplicates_value_sum, 2))
        dup_sheet[get_column_letter(dup_col) + str(last_row + 5)].value = str(counter)

        print(f"Duplicates: {duplicate}")
        DUPLICATES_VOLUME13 = counter
        DUPLICATES_VALUE13 = duplicates_value_sum

        wb.close()
        wb.save(match_OVA)

        if matchOVAFound is True and matchINTFound is True:
            #  ###################################################### MISSING INT TRANSACTIONS ##############################################################

            Owb = load_workbook(match_OVA)
            Osheet = Owb[Owb.sheetnames[0]]
            Owb.active = Owb[Owb.sheetnames[0]]
            file = pd.read_excel(match_OVA)

            Iwb = load_workbook(match_INT)
            Isheet = Iwb[Iwb.sheetnames[0]]
            Iwb.active = Iwb[Iwb.sheetnames[0]]

            Ofirst_col = Owb.active.min_column
            Olast_col = Owb.active.max_column
            Ofirst_row = Owb.active.min_row
            Olast_row = Owb.active.max_row

            Ifirst_col = Iwb.active.min_column
            Ilast_col = Iwb.active.max_column
            Ifirst_row = Iwb.active.min_row
            Ilast_row = Iwb.active.max_row

            missing_INT_list = []
            missing_OVA_list = []
            matchFound = False
            headerChecker = True
            counter = 0
            for i in range(Ofirst_row + 1, Olast_row + 1):
                for j in ids:
                    if str(Osheet[ova_id_col + str(i)].value) in str(j):
                        matchFound = True
                        break
                    else:
                        matchFound = False
                if matchFound is False:
                    missing_INT_list.append(Osheet[ova_id_col + str(i)].value)
            print(f"Missing integrator transactions: {missing_INT_list}")
            ws2 = Iwb.create_sheet('Missing Integrator Transactions')
            ws2.title = 'Missing Integrator Transactions'
            count = 0

            for transaction in missing_INT_list:
                data = file[file['Order ID'] == transaction]
                if counter > 0:
                    headerChecker = False
                with pd.ExcelWriter(match_OVA, mode='a', engine="openpyxl", if_sheet_exists='overlay') as writer:
                    data.to_excel(writer, sheet_name='Missing Integrator Transactions', startrow=count,
                                  header=headerChecker)
                    if headerChecker is False:
                        count += 1
                    else:
                        count += 2
                    counter += 1
            Iwb.close()
            Owb.close()

        # ######################################### MISSING OVA TRANSACTIONS ##############################################
        Owb = load_workbook(match_OVA)
        Osheet = Owb[Owb.sheetnames[0]]
        Owb.active = Owb[Owb.sheetnames[0]]

        Iwb = load_workbook(match_INT)
        file = pd.read_excel(match_INT)
        Isheet = Iwb[Iwb.sheetnames[0]]
        Iwb.active = Iwb[Iwb.sheetnames[0]]

        Ofirst_col = Owb.active.min_column
        Olast_col = Owb.active.max_column
        Ofirst_row = Owb.active.min_row
        Olast_row = Owb.active.max_row

        Ifirst_col = Iwb.active.min_column
        Ilast_col = Iwb.active.max_column
        Ifirst_row = Iwb.active.min_row
        Ilast_row = Iwb.active.max_row

        missing_INT_list = []
        missing_OVA_list = []
        matchFound = False
        headerChecker = True
        counter = 0

        for i in ids:
            for j in range(Ofirst_row + 1, Olast_row + 1):
                if str(Osheet[ova_id_col + str(j)].value) in str(i):
                    matchFound = True
                    break
                else:
                    matchFound = False
            if matchFound is False:
                missing_OVA_list.append(str(i))
        print(f"Missing OVA transactions: {missing_OVA_list}")
        count = 0
        for transaction in missing_OVA_list:
            data = file[file['External Transaction Reference'] == transaction]
            if counter > 0:
                headerChecker = False
            with pd.ExcelWriter(match_OVA, mode='a', engine="openpyxl", if_sheet_exists='overlay') as writer:
                data.to_excel(writer, sheet_name='Missing OVA Transactions', startrow=count, header=headerChecker)
                if headerChecker is False:
                    count += 1
                else:
                    count += 2
                counter += 1
        Iwb.close()
        Owb.close()

    elif ("MPGS KC" + yesterday + ".xlsx" not in dir_list) and ('MPGS_trn' + yesterday + '.xlsx' in dir_list):
        print('found kb')
        MPGS_INT_Volume = 0
        MPGS_INT_Sum = 0.00

        wb = load_workbook('MPGS_trn' + yesterday + '.xlsx')
        sheet = wb[wb.sheetnames[0]]
        wb.active = wb[wb.sheetnames[0]]

        first_col = wb.active.min_column
        last_col = wb.active.max_column
        first_row = wb.active.min_row
        last_row = wb.active.max_row

        amountColumn = get_column("Amount")
        statusFlagColumn = get_column("Status")
        int_id_col = get_column("Transaction Id")
        match_INT = 'MPGS_trn' + yesterday + '.xlsx'

        ids = []

        first_col = wb.active.min_column
        last_col = wb.active.max_column
        first_row = wb.active.min_row
        last_row = wb.active.max_row

        for i in range(first_row + 1, last_row + 1):
            if sheet[statusFlagColumn + str(i)].value == "CONFIRMED":
                MPGS_INT_Sum = MPGS_INT_Sum + float(sheet[amountColumn + str(i)].value)
                MPGS_INT_Volume += 1
                ids.append(str(sheet[int_id_col + str(i)].value))

        print('MPGS_INT_Volume: ' + str(MPGS_INT_Volume))
        print('MPGS_INT_Sum: ' + str(MPGS_INT_Sum))
        INT_VALUE13 = MPGS_INT_Sum
        INT_VOLUME13 = MPGS_INT_Volume
        wb.close()
        wb.save(match_INT)

        #   ------------------------------------------- DUPLICATES -------------------------------------------------------
        wb = load_workbook(match_INT)
        Owb = load_workbook(match_OVA)
        sheet = wb[wb.sheetnames[0]]
        wb.active = wb[wb.sheetnames[0]]

        first_col = wb.active.min_column
        last_col = wb.active.max_column
        first_row = wb.active.min_row
        last_row = wb.active.max_row
        int_id_col = get_column("Receipt No.")
        list = []
        duplicate = []
        count = 1
        counter = 0
        duplicates_value_sum = 0.00

        for i in range(first_row + 1, last_row + 1):
            if sheet[statusFlagColumn + str(i)].value == "CONFIRMED":
                if sheet[int_id_col + str(i)].value not in list:
                    list.append(sheet[int_id_col + str(i)].value)
                else:
                    if sheet[int_id_col + str(i)].value not in duplicate:
                        data = file[file['Transaction Id'] == (sheet[int_id_col + str(i)].value)]
                        duplicate.append(sheet[int_id_col + str(i)].value)
                        if counter > 0:
                            headerChecker = False
                        with pd.ExcelWriter(match_OVA, mode='a', engine="openpyxl", if_sheet_exists='overlay') as writer:
                            data.to_excel(writer, sheet_name='Duplicates', startrow=count, header=headerChecker)
                            if headerChecker is False:
                                count += 2
                            else:
                                count += 3
                            counter += 1
                    else:
                        continue

        wb.close()
        wb.save(match_INT)
        Owb.close()
        Owb.save(match_OVA)

        wb = load_workbook(match_OVA)
        wb.active = wb['Duplicates']
        dup_sheet = wb['Duplicates']
        sheet = dup_sheet

        first_col = wb.active.min_column
        last_col = wb.active.max_column
        first_row = wb.active.min_row
        last_row = wb.active.max_row

        dup_col = column_index_from_string(amountColumn) + 1
        for number in range(first_row + 2, last_row + 1, 2):
            duplicates_value_sum = duplicates_value_sum + dup_sheet[get_column_letter(dup_col) + str(number)].value
            count += 3

        print(f"Number of duplicates: {counter}")
        dup_sheet[get_column_letter(dup_col) + str(last_row + 4)].value = str(round(duplicates_value_sum, 2))
        dup_sheet[get_column_letter(dup_col) + str(last_row + 5)].value = str(counter)

        print(f"Duplicates: {duplicate}")

        DUPLICATES_VOLUME13 = counter
        DUPLICATES_VALUE13 = duplicates_value_sum
        wb.close()
        wb.save(match_OVA)
        # ------------------------------------- MISSING INT TRANSACTIONS ----------------------------------------------------
        match_INT = 'MPGS_trn' + yesterday + '.xlsx'
        Owb = load_workbook(match_OVA)
        Osheet = Owb[Owb.sheetnames[0]]
        Owb.active = Owb[Owb.sheetnames[0]]
        file = pd.read_excel(match_OVA)

        Iwb = load_workbook(match_INT)
        Isheet = Iwb[Iwb.sheetnames[0]]
        Iwb.active = Iwb[Iwb.sheetnames[0]]

        Ofirst_col = Owb.active.min_column
        Olast_col = Owb.active.max_column
        Ofirst_row = Owb.active.min_row
        Olast_row = Owb.active.max_row

        Ifirst_col = Iwb.active.min_column
        Ilast_col = Iwb.active.max_column
        Ifirst_row = Iwb.active.min_row
        Ilast_row = Iwb.active.max_row

        missing_INT_list = []
        missing_OVA_list = []
        matchFound = False
        headerChecker = True
        counter = 0
        for i in range(Ofirst_row + 1, Olast_row + 1):
            for j in ids:
                if str(Osheet[ova_id_col + str(i)].value) in str(j):
                    matchFound = True
                    break
                else:
                    matchFound = False
            if matchFound is False:
                missing_INT_list.append(Osheet[ova_id_col + str(i)].value)
        print(f"Missing integrator transactions: {missing_INT_list}")
        ws2 = Iwb.create_sheet('Missing Integrator Transactions')
        ws2.title = 'Missing Integrator Transactions'
        count = 0

        for transaction in missing_INT_list:
            data = file[file['Order ID'] == transaction]
            if counter > 0:
                headerChecker = False
            with pd.ExcelWriter(match_OVA, mode='a', engine="openpyxl", if_sheet_exists='overlay') as writer:
                data.to_excel(writer, sheet_name='Missing Integrator Transactions', startrow=count,
                              header=headerChecker)
                if headerChecker is False:
                    count += 1
                else:
                    count += 2
                counter += 1
        Iwb.close()
        Owb.close()

        # -------------------------------------- MISSING OVA TRANSACTIONS ----------------------------------------------

        Owb = load_workbook(match_OVA)
        Osheet = Owb[Owb.sheetnames[0]]
        Owb.active = Owb[Owb.sheetnames[0]]

        match_INT = 'MPGS_trn' + yesterday + '.xlsx'
        Iwb = load_workbook(match_INT)
        file = pd.read_excel(match_INT)
        Isheet = Iwb[Iwb.sheetnames[0]]
        Iwb.active = Iwb[Iwb.sheetnames[0]]

        Ofirst_col = Owb.active.min_column
        Olast_col = Owb.active.max_column
        Ofirst_row = Owb.active.min_row
        Olast_row = Owb.active.max_row

        Ifirst_col = Iwb.active.min_column
        Ilast_col = Iwb.active.max_column
        Ifirst_row = Iwb.active.min_row
        Ilast_row = Iwb.active.max_row

        missing_INT_list = []
        missing_OVA_list = []
        matchFound = False
        headerChecker = True
        counter = 0

        for i in ids:
            for j in range(Ofirst_row + 1, Olast_row + 1):
                if str(Osheet[ova_id_col + str(j)].value) in str(i):
                    matchFound = True
                    break
                else:
                    matchFound = False
            if matchFound is False:
                missing_OVA_list.append(str(i))
        print(f"Missing OVA transactions: {missing_OVA_list}")
        count = 0
        for transaction in missing_OVA_list:
            data = file[file['Transaction Id'] == transaction]
            if counter > 0:
                headerChecker = False
            with pd.ExcelWriter(match_OVA, mode='a', engine="openpyxl", if_sheet_exists='overlay') as writer:
                data.to_excel(writer, sheet_name='Missing OVA Transactions', startrow=count, header=headerChecker)
                if headerChecker is False:
                    count += 1
                else:
                    count += 2
                counter += 1
        Iwb.close()
        Owb.close()

except FileNotFoundError:
    print("file not found")
    INT_VALUE13 = 0.00
    INT_VOLUME13 = 0
    DUPLICATES_VOLUME13 = 0
    DUPLICATES_VALUE13 = 0

#
fileFound = False
matchOVAFound = False
matchINTFound = False
#
# ################################################ MTN KR CREDIT OVA ################################################
for file in dir_list:
    if file == 'KR MTN Credit' + yesterday + '.xlsx':
        fileFound = True
        matchOVAFound = True
        wb = load_workbook('KR MTN Credit' + yesterday + '.xlsx')
        sheet = wb[wb.sheetnames[0]]
        wb.active = wb[wb.sheetnames[0]]
        # DEFINE MAX AND MIN COLUMNS AND ROWS
        first_col = wb.active.min_column
        last_col = wb.active.max_column
        first_row = wb.active.min_row
        last_row = wb.active.max_row

        MTN_KR_Credit_OVA_Volume = (last_row + 1) - (first_row + 1)
        OVA_VOLUME14 = MTN_KR_Credit_OVA_Volume
        print('MTN_KR_Credit_OVA_Volume: ' + str(MTN_KR_Credit_OVA_Volume))

        MTN_KR_Credit_OVA_Sum = 0.00

        amountColumn = get_column('Amount')
        ova_id_col = get_column('External Transaction Id')
        tmp_id_col = ova_id_col
        alternate_id_col = get_column("Id")
        for i in range(first_row + 1, last_row + 1):
            MTN_KR_Credit_OVA_Sum = MTN_KR_Credit_OVA_Sum + abs(float(sheet[amountColumn + str(i)].value))

        print('MTN_KR_Credit_OVA_Sum: ' + str(MTN_KR_Credit_OVA_Sum))
        OVA_VALUE14 = MTN_KR_Credit_OVA_Sum
        wb.close()
        wb.save('KR MTN Credit' + yesterday + '- Recons.xlsx')
        match_OVA = 'KR MTN Credit' + yesterday + '- Recons.xlsx'

        wb = load_workbook(match_OVA)
        ws1 = wb.create_sheet("Duplicates")
        ws1.title = 'Duplicates'

        ws1 = wb.create_sheet("Missing Integrator Transactions")
        ws1.title = 'Missing Integrator Transactions'

        ws1 = wb.create_sheet("Missing Integrator Transactions")
        ws1.title = 'Missing Integrator Transactions'

        wb.close()
        wb.save(match_OVA)
        break
    else:
        fileFound = False
if fileFound is False:
    OVA_VOLUME14 = 0
    OVA_VALUE14 = 0
#
#
#
fileFound = False
#
# ############################################### MTN KR CREDIT INT ############################################
for file in dir_list:
    if file == 'KR MTN Disb_mBase' + yesterday + '.xlsx':
        fileFound = True
        matchINTFound = True
        match_INT = 'KR MTN Disb_mBase' + yesterday + '.xlsx'
        wb = load_workbook(match_INT)
        sheet = wb[wb.sheetnames[0]]
        wb.active = wb[wb.sheetnames[0]]
        file = pd.read_excel(match_INT)

        # DEFINE MAX AND MIN COLUMNS AND ROWS
        first_col = wb.active.min_column
        last_col = wb.active.max_column
        first_row = wb.active.min_row
        last_row = wb.active.max_row

        MTN_KR_Credit_INT_Volume = (last_row + 1) - (first_row + 1)
        INT_VOLUME14 = MTN_KR_Credit_INT_Volume
        print('MTN_KR_CREDIT_INT_Volume: ' + str(MTN_KR_Credit_INT_Volume))

        MTN_KR_Credit_INT_Sum = 0.00

        amountColumn = get_column('Amount')
        int_id_col = get_column('IntegratorTransId')
        for i in range(first_row + 1, last_row + 1):
            MTN_KR_Credit_INT_Sum = MTN_KR_Credit_INT_Sum + abs(float(sheet[amountColumn + str(i)].value))

        print('MTN_KR_CREDIT_INT_Sum: ' + str(MTN_KR_Credit_INT_Sum))
        INT_VALUE14 = MTN_KR_Credit_INT_Sum
        wb.close()
        wb.save(match_INT)
        # -------------------------------------------- Duplicates -----------------------------------------------
        wb = load_workbook(match_INT)
        sheet = wb[wb.sheetnames[0]]
        wb.active = wb[wb.sheetnames[0]]
        Owb = load_workbook(match_OVA)
        file = pd.read_excel(match_INT)
        list = []
        duplicates = []
        count = 1
        counter = 0
        headerChecker = True
        duplicates_value_sum = 0.00
        for id in range(first_row + 1, last_row + 1):
            if sheet[int_id_col + str(id)].value not in list:
                list.append(sheet[int_id_col + str(id)].value)
            elif sheet[int_id_col + str(id)].value not in duplicates:
                    data = file[file['IntegratorTransId'] == (sheet[int_id_col + str(id)].value)]
                    data_set_sum = data.iloc[1:, data.columns.get_loc("Amount")].sum()
                    duplicates_value_sum = duplicates_value_sum + data_set_sum
                    duplicates.append(sheet[int_id_col + str(id)].value)
                    if counter > 0:
                        headerChecker = False
                    with pd.ExcelWriter(match_OVA, mode='a', engine="openpyxl", if_sheet_exists='overlay') as writer:
                        data.to_excel(writer, sheet_name='Duplicates', startrow=count, header=headerChecker)

                        if headerChecker is False:
                            count += 2
                        else:
                            count += (data.shape[0] +1)
                        counter += 1
            else:
                if sheet[int_id_col + str(id)].value in duplicates:
                    counter += 1
                continue
        wb.close()
        Owb.close()

        wb = load_workbook(match_OVA)
        wb.active = wb['Duplicates']
        dup_sheet = wb['Duplicates']
        sheet = dup_sheet

        first_col = wb.active.min_column
        last_col = wb.active.max_column
        first_row = wb.active.min_row
        last_row = wb.active.max_row

        dup_col = column_index_from_string(amountColumn) + 1

        print(f"Number of duplicates: {counter}")
        dup_sheet[get_column_letter(dup_col) + str(last_row + 4)].value = str(round(duplicates_value_sum, 2))
        dup_sheet[get_column_letter(dup_col) + str(last_row + 5)].value = str(counter)

        print(f"Duplicates: {duplicates}")
        DUPLICATES_VOLUME14 = counter
        DUPLICATES_VALUE14 = duplicates_value_sum

        wb.close()
        wb.save(match_OVA)
#
        if matchOVAFound is True and matchINTFound is True:
            # ###################################### MISSING INT TRANSACTIONS ################################################
            Owb = load_workbook(match_OVA)
            Osheet = Owb[Owb.sheetnames[0]]
            Owb.active = Owb[Owb.sheetnames[0]]

            Iwb = load_workbook(match_INT)
            Isheet = Iwb[Iwb.sheetnames[0]]
            Iwb.active = Iwb[Iwb.sheetnames[0]]

            Ofirst_col = Owb.active.min_column
            Olast_col = Owb.active.max_column
            Ofirst_row = Owb.active.min_row
            Olast_row = Owb.active.max_row

            Ifirst_col = Iwb.active.min_column
            Ilast_col = Iwb.active.max_column
            Ifirst_row = Iwb.active.min_row
            Ilast_row = Iwb.active.max_row

            missing_INT_list = []
            missing_OVA_list = []
            matchFound = False
            headerChecker = True
            counter = 0

            for i in range(Ofirst_row + 1, Olast_row + 1):
                for j in range(Ifirst_row + 1, Ilast_row + 1):
                    if Osheet[ova_id_col + str(i)].value is None or Osheet[ova_id_col + str(i)].value == "#VALUE!":
                        ova_id_col = alternate_id_col
                        matchFound = False
                        break
                    if str(Osheet[ova_id_col + str(i)].value) == str(Isheet[int_id_col + str(j)].value):
                        matchFound = True
                        break
                    else:
                        matchFound = False
                if matchFound is False:
                    missing_INT_list.append(Osheet[ova_id_col + str(i)].value)
                    ova_id_col = tmp_id_col
            print(f"Missing integrator transactions: {missing_INT_list}")
            file = pd.read_excel(match_OVA)
            count = 0
            for transaction in missing_INT_list:
                data = file[file['External Transaction Id'] == transaction]
                if data.empty:
                    data = file[file['Id'] == transaction]
                if counter > 0:
                    headerChecker = False
                with pd.ExcelWriter(match_OVA, mode='a', engine="openpyxl", if_sheet_exists='overlay') as writer:
                    data.to_excel(writer, sheet_name='Missing Integrator Transactions', startrow=count,
                                  header=headerChecker)
                    if headerChecker is False:
                        count += 1
                    else:
                        count += 2
                    counter += 1

            Iwb.close()
            Owb.close()

            # ######################################### MISSING OVA TRANSACTIONS ##############################################
            Owb = load_workbook(match_OVA)
            Osheet = Owb[Owb.sheetnames[0]]
            Owb.active = Owb[Owb.sheetnames[0]]

            Iwb = load_workbook(match_INT)
            file = pd.read_excel(match_INT)
            Isheet = Iwb[Iwb.sheetnames[0]]
            Iwb.active = Iwb[Iwb.sheetnames[0]]

            Ofirst_col = Owb.active.min_column
            Olast_col = Owb.active.max_column
            Ofirst_row = Owb.active.min_row
            Olast_row = Owb.active.max_row

            Ifirst_col = Iwb.active.min_column
            Ilast_col = Iwb.active.max_column
            Ifirst_row = Iwb.active.min_row
            Ilast_row = Iwb.active.max_row

            missing_INT_list = []
            missing_OVA_list = []
            matchFound = False
            headerChecker = True
            counter = 0

            for i in range(Ifirst_row + 1, Ilast_row + 1):
                for j in range(Ofirst_row + 1, Olast_row + 1):
                    if str(Isheet[int_id_col + str(i)].value) == str(Osheet[ova_id_col + str(j)].value):
                        matchFound = True
                        break
                    else:
                        matchFound = False
                if matchFound is False:
                    missing_OVA_list.append(Isheet[int_id_col + str(i)].value)
            print(f"Missing OVA transactions: {missing_OVA_list}")
            count = 0
            for transaction in missing_OVA_list:
                data = file[file['IntegratorTransId'] == transaction]
                if counter > 0:
                    headerChecker = False
                with pd.ExcelWriter(match_OVA, mode='a', engine="openpyxl", if_sheet_exists='overlay') as writer:
                    data.to_excel(writer, sheet_name='Missing OVA Transactions', startrow=count, header=headerChecker)
                    if headerChecker is False:
                        count += 1
                    else:
                        count += 2
                    counter += 1
            Iwb.close()
            Owb.close()
        break
    else:
        fileFound = False

if fileFound is False:
    INT_VOLUME14 = 0
    INT_VALUE14 = 0
    DUPLICATES_VOLUME14 = 0
    DUPLICATES_VALUE14 = 0

#
#
fileFound = False
matchOVAFound = False
matchINTFound = False
#
# ################################################# MTN KR DEBIT OVA #################################################
for file in dir_list:
    if file == 'KR MTN Debit' + yesterday + '.xlsx':
        fileFound = True
        matchOVAFound = True
        match_OVA = 'KR MTN Debit' + yesterday + '.xlsx'
        wb = load_workbook(match_OVA)
        sheet = wb[wb.sheetnames[0]]
        wb.active = wb[wb.sheetnames[0]]

        # DEFINE MAX AND MIN COLUMNS AND ROWS
        first_col = wb.active.min_column
        last_col = wb.active.max_column
        first_row = wb.active.min_row
        last_row = wb.active.max_row

        MTN_KR_Debit_OVA_Volume = (last_row + 1) - (first_row + 1)
        OVA_VOLUME15 = MTN_KR_Debit_OVA_Volume
        print('MTN_KR_Debit_OVA_Volume: ' + str(MTN_KR_Debit_OVA_Volume))

        MTN_KR_Debit_OVA_Sum = 0.00

        amountColumn = get_column('Amount')
        ova_id_col = get_column('External Transaction Id')
        tmp_id_col = ova_id_col
        alternate_id_col = get_column('Id')
        for i in range(first_row + 1, last_row + 1):
            MTN_KR_Debit_OVA_Sum = MTN_KR_Debit_OVA_Sum + abs(float(sheet[amountColumn + str(i)].value))

        print('MTN_KR_Debit_OVA_Sum: ' + str(MTN_KR_Debit_OVA_Sum))
        OVA_VALUE15 = MTN_KR_Debit_OVA_Sum
        wb.close()
        wb.save('KR MTN Debit' + yesterday + '- Recons.xlsx')
        match_OVA = 'KR MTN Debit' + yesterday + '- Recons.xlsx'

        wb = load_workbook(match_OVA)
        ws1 = wb.create_sheet("Duplicates")
        ws1.title = 'Duplicates'

        ws1 = wb.create_sheet("Missing Integrator Transactions")
        ws1.title = 'Missing Integrator Transactions'

        ws1 = wb.create_sheet("Missing Integrator Transactions")
        ws1.title = 'Missing Integrator Transactions'

        wb.close()
        wb.save(match_OVA)
        break
    else:
        fileFound = False

if fileFound is False:
    OVA_VOLUME15 = 0
    OVA_VALUE15 = 0
#
#
fileFound = False
#
# ############################################ MTN KR DEBIT INT ###################################################
for file in dir_list:
    if file == 'KR MTN Coll_mBase' + yesterday + '.xlsx':
        fileFound = True
        matchINTFound = True
        match_INT = 'KR MTN Coll_mBase' + yesterday + '.xlsx'
        wb = load_workbook(match_INT)
        sheet = wb[wb.sheetnames[0]]
        wb.active = wb[wb.sheetnames[0]]
        file = pd.read_excel(match_INT)

        # DEFINE MAX AND MIN COLUMNS AND ROWS
        first_col = wb.active.min_column
        last_col = wb.active.max_column
        first_row = wb.active.min_row
        last_row = wb.active.max_row

        MTN_KR_Debit_INT_Volume = (last_row + 1) - (first_row + 1)
        INT_VOLUME15 = MTN_KR_Debit_INT_Volume
        print('MTN_KR_Debit_INT_Volume: ' + str(MTN_KR_Debit_INT_Volume))

        MTN_KR_Debit_INT_Sum = 0.00

        amountColumn = get_column('Amount')
        int_id_col = get_column('IntegratorTransId')
        for i in range(first_row + 1, last_row + 1):
            MTN_KR_Debit_INT_Sum = MTN_KR_Debit_INT_Sum + float(sheet[amountColumn + str(i)].value)

        print('MTN_KR_Debit_INT_Sum: ' + str(MTN_KR_Debit_INT_Sum))
        INT_VALUE15 = MTN_KR_Debit_INT_Sum
        wb.close()
        wb.save(match_INT)
        # ---------------------------------------- DUPLICATES -------------------------------------------------
        wb = load_workbook(match_INT)
        sheet = wb[wb.sheetnames[0]]
        wb.active = wb[wb.sheetnames[0]]

        first_col = wb.active.min_column
        last_col = wb.active.max_column
        first_row = wb.active.min_row
        last_row = wb.active.max_row

        Owb = load_workbook(match_OVA)

        list = []
        duplicates = []
        count = 0
        counter = 0
        headerChecker = True
        duplicates_value_sum = 0.00
        for id in range(first_row + 1, last_row + 1):
            if sheet[int_id_col + str(id)].value not in list:
                list.append(sheet[int_id_col + str(id)].value)
            elif sheet[int_id_col+str(id)].value not in duplicates:
                    data = file[file['IntegratorTransId'] == (sheet[int_id_col + str(id)].value)]
                    data_set_sum = data.iloc[1:, data.columns.get_loc("Amount")].sum()
                    duplicates_value_sum = duplicates_value_sum + data_set_sum
                    duplicates.append(sheet[int_id_col + str(id)].value)
                    if counter > 0:
                        headerChecker = False
                    with pd.ExcelWriter(match_OVA, mode='a', engine="openpyxl", if_sheet_exists='overlay') as writer:
                        data.to_excel(writer, sheet_name='Duplicates', startrow=count, header=headerChecker)
                        if headerChecker is False:
                            count += 2
                        else:
                            count += (data.shape[0] +1)
                        counter += 1
            else:
                if sheet[int_id_col + str(id)].value in duplicates:
                    counter += 1
                continue
        wb.close()
        wb.save('KR MTN Coll_mBase' + yesterday + '.xlsx')
        Owb.close()

        Owb = load_workbook(match_OVA)
        Owb.active = Owb['Duplicates']
        dup_sheet = Owb['Duplicates']
        sheet = dup_sheet

        first_col = Owb.active.min_column
        last_col = Owb.active.max_column
        first_row = Owb.active.min_row
        last_row = Owb.active.max_row

        dup_col = column_index_from_string(amountColumn) + 1

        print(f"Number of duplicates: {counter}")
        dup_sheet[get_column_letter(dup_col) + str(last_row + 4)].value = str(round(duplicates_value_sum, 2))
        dup_sheet[get_column_letter(dup_col) + str(last_row + 5)].value = str(counter)

        print(f"Duplicates: {duplicates}")

        DUPLICATES_VOLUME15 = counter
        DUPLICATES_VALUE15 = duplicates_value_sum

        Owb.close()
        Owb.save(match_OVA)

        if matchOVAFound is True and matchINTFound is True:
            # ########################################### MISSING INT TRANSACTIONS ###############################################
            Owb = load_workbook(match_OVA)
            Osheet = Owb[Owb.sheetnames[0]]
            Owb.active = Owb[Owb.sheetnames[0]]
            file = pd.read_excel(match_OVA)

            Iwb = load_workbook(match_INT)
            Isheet = Iwb[Iwb.sheetnames[0]]
            Iwb.active = Iwb[Iwb.sheetnames[0]]

            Ofirst_col = Owb.active.min_column
            Olast_col = Owb.active.max_column
            Ofirst_row = Owb.active.min_row
            Olast_row = Owb.active.max_row

            Ifirst_col = Iwb.active.min_column
            Ilast_col = Iwb.active.max_column
            Ifirst_row = Iwb.active.min_row
            Ilast_row = Iwb.active.max_row

            missing_INT_list = []
            missing_OVA_list = []
            matchFound = False
            headerChecker = True
            counter = 0
            isInt = False
            for i in range(Ofirst_row + 1, Olast_row + 1):
                for j in range(Ifirst_row + 1, Ilast_row + 1):
                    if Osheet[ova_id_col + str(i)].value is None or Osheet[ova_id_col + str(i)].value == "#VALUE!":
                        ova_id_col = alternate_id_col
                        isInt = True
                        matchFound = False
                        break
                    if str(Osheet[ova_id_col + str(i)].value).strip('=') == str(
                            Isheet[int_id_col + str(j)].value).strip('='):
                        matchFound = True
                        break
                    else:
                        matchFound = False
                if matchFound is False:
                    if isInt is False:
                        missing_INT_list.append(str(Osheet[ova_id_col + str(i)].value).strip('='))
                    else:
                        missing_INT_list.append(str(Osheet[ova_id_col + str(i)].value))
                    ova_id_col = tmp_id_col

            print(f"Missing integrator transactions: {missing_INT_list}")
            count = 0
            for transaction in missing_INT_list:
                if type(transaction) is str:
                    data = file[file['External Transaction Id'] == transaction]
                else:
                    data = file[file['Id'] == int(transaction)]
                if counter > 0:
                    headerChecker = False
                with pd.ExcelWriter(match_OVA, mode='a', engine="openpyxl", if_sheet_exists='overlay') as writer:
                    data.to_excel(writer, sheet_name='Missing Integrator Transactions', startrow=count,
                                  header=headerChecker)
                    if headerChecker is False:
                        count += 1
                    else:
                        count += 2
                    counter += 1

            Iwb.close()
            Owb.close()

            # ######################################### MISSING OVA TRANSACTIONS ##############################################
            Owb = load_workbook(match_OVA)
            Osheet = Owb[Owb.sheetnames[0]]
            Owb.active = Owb[Owb.sheetnames[0]]

            Iwb = load_workbook(match_INT)
            Isheet = Iwb[Iwb.sheetnames[0]]
            Iwb.active = Iwb[Iwb.sheetnames[0]]
            file = pd.read_excel(match_INT)

            Ofirst_col = Owb.active.min_column
            Olast_col = Owb.active.max_column
            Ofirst_row = Owb.active.min_row
            Olast_row = Owb.active.max_row

            Ifirst_col = Iwb.active.min_column
            Ilast_col = Iwb.active.max_column
            Ifirst_row = Iwb.active.min_row
            Ilast_row = Iwb.active.max_row

            missing_INT_list = []
            missing_OVA_list = []
            matchFound = False
            headerChecker = True
            counter = 0

            for i in range(Ifirst_row + 1, Ilast_row + 1):
                for j in range(Ofirst_row + 1, Olast_row + 1):
                    if str(Isheet[int_id_col + str(i)].value).strip('=') == str(
                            Osheet[ova_id_col + str(j)].value).strip('='):
                        matchFound = True
                        break
                    else:
                        matchFound = False
                if matchFound is False:
                    missing_OVA_list.append(str(Isheet[int_id_col + str(i)].value).strip('='))
            print(f"Missing OVA transactions: {missing_OVA_list}")
            count = 0
            for transaction in missing_OVA_list:
                data = file[file['IntegratorTransId'] == transaction]
                if counter > 0:
                    headerChecker = False
                with pd.ExcelWriter(match_OVA, mode='a', engine="openpyxl", if_sheet_exists='overlay') as writer:
                    data.to_excel(writer, sheet_name='Missing OVA Transactions', startrow=count, header=headerChecker)
                    if headerChecker is False:
                        count += 1
                    else:
                        count += 2
                    counter += 1
            Iwb.close()
            Owb.close()
        break
    else:
        fileFound = False

if fileFound is False:
    INT_VOLUME15 = 0
    INT_VALUE15 = 0
    DUPLICATES_VOLUME15 = 0
    DUPLICATES_VALUE15 = 0
#
#
matchOVAFound = False
matchINTFound = False
fileFound = False
# ##################################### AIRTEL KR CASHIN OVA ############################################################
for file in dir_list:
    if file == 'KR AirtelTigo' + yesterday + '.xlsx':
        fileFound = True
        matchOVAFound = True
        match_OVA = 'KR AirtelTigo' + yesterday + '.xlsx'
        wb = load_workbook(match_OVA)
        sheet = wb[wb.sheetnames[0]]
        wb.active = wb[wb.sheetnames[0]]
        file = pd.read_excel(match_OVA)

        first_col = wb.active.min_column
        last_col = wb.active.max_column
        first_row = wb.active.min_row
        last_row = wb.active.max_row

        AIRTEL_KR_Cashin_OVA_Volume = 0
        AIRTEL_KR_Cashin_OVA_Sum = 0.00

        serviceTypeCol = get_column("Service Type")
        amountColumn = get_column('Transaction Amount')
        ova_id_col = get_column('External Transaction Id')

        for i in range(first_row+1, last_row+1):
            if sheet[serviceTypeCol+str(i)].value == "Merchant Payment":
                AIRTEL_KR_Cashin_OVA_Volume = AIRTEL_KR_Cashin_OVA_Volume + 1
                AIRTEL_KR_Cashin_OVA_Sum = AIRTEL_KR_Cashin_OVA_Sum + sheet[amountColumn+str(i)].value


        OVA_VOLUME16 = AIRTEL_KR_Cashin_OVA_Volume
        print('AIRTEL_KR_Cashin_OVA_Volume: ' + str(AIRTEL_KR_Cashin_OVA_Volume))
        print('AIRTEL_KR_Cashin_OVA_Sum: ' + str(AIRTEL_KR_Cashin_OVA_Sum))
        OVA_VALUE16 = AIRTEL_KR_Cashin_OVA_Sum
        wb.close()
        wb.save('KR AirtelTigo Cashin' + yesterday + '- Recons.xlsx')
        match_OVA = 'KR AirtelTigo Cashin' + yesterday + '- Recons.xlsx'

        wb = load_workbook(match_OVA)
        ws1 = wb.create_sheet("Duplicates")
        ws1.title = 'Duplicates'

        ws1 = wb.create_sheet("Missing Integrator Transactions")
        ws1.title = 'Missing Integrator Transactions'

        ws1 = wb.create_sheet("Missing Integrator Transactions")
        ws1.title = 'Missing Integrator Transactions'

        wb.close()
        wb.save(match_OVA)
        break
    else:
        fileFound = False

if fileFound is False:
    OVA_VOLUME16 = 0
    OVA_VALUE16 = 0

fileFound = False

# ####################################### AIRTEL KR CASHIN INT #####################################################
for file in dir_list:
    if file == 'KR AirtelTigo Coll_mBase' + yesterday + '.xlsx':
        fileFound = True
        matchINTFound = True
        match_INT = 'KR AirtelTigo Coll_mBase' + yesterday + '.xlsx'
        wb = load_workbook(match_INT)
        sheet = wb[wb.sheetnames[0]]
        wb.active = wb[wb.sheetnames[0]]
        file = pd.read_excel(match_INT)

        # DEFINE MAX AND MIN COLUMNS AND ROWS
        first_col = wb.active.min_column
        last_col = wb.active.max_column
        first_row = wb.active.min_row
        last_row = wb.active.max_row

        statusFlagColumn = get_column("Status")
        amountColumn = get_column('Amount')
        int_id_col = get_column('Transaction Id')
        init_list = []
        AIRTEL_KR_Cashin_INT_Volume = 0
        AIRTEL_KR_Cashin_INT_Sum = 0.00

        for i in range(first_row + 1, last_row + 1):
            if sheet[statusFlagColumn + str(i)].value == 'CONFIRMED':
                AIRTEL_KR_Cashin_INT_Volume = AIRTEL_KR_Cashin_INT_Volume + 1
                AIRTEL_KR_Cashin_INT_Sum = AIRTEL_KR_Cashin_INT_Sum + float(sheet[amountColumn + str(i)].value)
                init_list.append(str(sheet[int_id_col + str(i)].value))

        print('AIRTEL_KR_Cashin_INT_Volume: ' + str(AIRTEL_KR_Cashin_INT_Volume))
        print('AIRTEL_KR_Cashin_INT_Sum: ' + str(AIRTEL_KR_Cashin_INT_Sum))
        INT_VOLUME16 = AIRTEL_KR_Cashin_INT_Volume
        INT_VALUE16 = AIRTEL_KR_Cashin_INT_Sum
        wb.close()
        wb.save(match_INT)
        # ---------------------------------------- DUPLICATES -------------------------------------------------
        wb = load_workbook(match_INT)
        Owb = load_workbook(match_OVA)
        sheet = wb[wb.sheetnames[0]]
        wb.active = wb[wb.sheetnames[0]]
        file = pd.read_excel(match_INT)
        duplicates = []
        count = 1
        list = []
        counter = 0
        headerChecker = True
        duplicates_value_sum = 0.00
        for id in init_list:
            if id not in list:
                list.append(id)
            elif id not in duplicates:
                    data = file[file['Transaction Id'] == id]
                    if data.empty:
                        data = file[file['Transaction Id'] == int(id)]
                    data_set_sum = data.iloc[1:, data.columns.get_loc("Amount")].sum()
                    duplicates_value_sum = duplicates_value_sum + data_set_sum
                    duplicates.append(id)
                    if counter > 0:
                        headerChecker = False
                    with pd.ExcelWriter(match_OVA, mode='a', engine="openpyxl", if_sheet_exists='overlay') as writer:
                        data.to_excel(writer, sheet_name='Duplicates', startrow=count, header=headerChecker)
                        if headerChecker is False:
                            count += 2
                        else:
                            count += (data.shape[0] +1)
                        counter += 1
            else:
                if id in duplicates:
                    counter += 1
                continue
        wb.close()
        wb.save(match_INT)
        Owb.close()

        wb = load_workbook(match_OVA)
        wb.active = wb['Duplicates']
        dup_sheet = wb['Duplicates']
        sheet = dup_sheet

        first_col = wb.active.min_column
        last_col = wb.active.max_column
        first_row = wb.active.min_row
        last_row = wb.active.max_row

        dup_col = column_index_from_string(amountColumn) + 1

        print(f"Number of duplicates: {counter}")
        dup_sheet[get_column_letter(dup_col) + str(last_row + 4)].value = str(round(duplicates_value_sum, 2))
        dup_sheet[get_column_letter(dup_col) + str(last_row + 5)].value = str(counter)

        print(f"Duplicates: {duplicates}")

        DUPLICATES_VOLUME16 = counter
        DUPLICATES_VALUE16 = duplicates_value_sum

        wb.close()
        wb.save(match_OVA)

        if matchOVAFound is True and matchINTFound is True:
            # ########################################### MISSING INT TRANSACTIONS ###############################################
            Owb = load_workbook(match_OVA)
            Osheet = Owb[Owb.sheetnames[0]]
            Owb.active = Owb[Owb.sheetnames[0]]
            file = pd.read_excel(match_OVA)

            Iwb = load_workbook(match_INT)
            Isheet = Iwb[Iwb.sheetnames[0]]
            Iwb.active = Iwb[Iwb.sheetnames[0]]

            Ofirst_col = Owb.active.min_column
            Olast_col = Owb.active.max_column
            Ofirst_row = Owb.active.min_row
            Olast_row = Owb.active.max_row

            Ifirst_col = Iwb.active.min_column
            Ilast_col = Iwb.active.max_column
            Ifirst_row = Iwb.active.min_row
            Ilast_row = Iwb.active.max_row

            missing_INT_list = []
            missing_OVA_list = []
            matchFound = False
            headerChecker = True
            counter = 0

            for i in range(Ofirst_row + 1, Olast_row + 1):
                if str(Osheet[serviceTypeCol + str(i)].value) == "Merchant Payment":
                    for j in range(Ifirst_row + 1, Ilast_row + 1):
                        if str(Osheet[ova_id_col + str(i)].value) == str(Isheet[int_id_col + str(j)].value):
                            matchFound = True
                            break
                        else:
                            matchFound = False
                    if matchFound is False:
                        missing_INT_list.append(Osheet[ova_id_col + str(i)].value)
            print(f"Missing integrator transactions: {missing_INT_list}")
            count = 0
            for transaction in missing_INT_list:
                data = file[file['External Transaction Id'] == transaction]
                if counter > 0:
                    headerChecker = False
                with pd.ExcelWriter(match_OVA, mode='a', engine="openpyxl", if_sheet_exists='overlay') as writer:
                    data.to_excel(writer, sheet_name='Missing Integrator Transactions', startrow=count,header=headerChecker)
                    if headerChecker is False:
                        count += 1
                    else:
                        count += 2
                    counter += 1
            Iwb.close()
            Owb.close()

            # ######################################### MISSING OVA TRANSACTIONS ##############################################
            Owb = load_workbook(match_OVA)
            Osheet = Owb[Owb.sheetnames[0]]
            Owb.active = Owb[Owb.sheetnames[0]]

            Iwb = load_workbook(match_INT)
            Isheet = Iwb[Iwb.sheetnames[0]]
            Iwb.active = Iwb[Iwb.sheetnames[0]]
            file = pd.read_excel(match_INT)

            Ofirst_col = Owb.active.min_column
            Olast_col = Owb.active.max_column
            Ofirst_row = Owb.active.min_row
            Olast_row = Owb.active.max_row

            Ifirst_col = Iwb.active.min_column
            Ilast_col = Iwb.active.max_column
            Ifirst_row = Iwb.active.min_row
            Ilast_row = Iwb.active.max_row

            missing_INT_list = []
            missing_OVA_list = []
            matchFound = False
            headerChecker = True
            counter = 0

            for i in init_list:
                for j in range(Ofirst_row + 1, Olast_row + 1):
                    if str(Osheet[serviceTypeCol + str(j)].value) == "Merchant Payment":
                        if str(i) == str(Osheet[ova_id_col + str(j)].value):
                            matchFound = True
                            break
                        else:
                            matchFound = False
                    else:
                        continue
                if matchFound is False:
                    missing_OVA_list.append(i)
            print(f"Missing OVA transactions: {missing_OVA_list}")
            count = 0
            for transaction in missing_OVA_list:
                data = file[file['Transaction Id'] == transaction]
                if counter > 0:
                    headerChecker = False
                with pd.ExcelWriter(match_OVA, mode='a', engine="openpyxl", if_sheet_exists='overlay') as writer:
                    data.to_excel(writer, sheet_name='Missing OVA Transactions', startrow=count, header=headerChecker)
                    if headerChecker is False:
                        count += 1
                    else:
                        count += 2
                    counter += 1
            Iwb.close()
            Owb.close()
        break
    else:
        fileFound = False

if fileFound is False:
    INT_VOLUME16 = 0
    INT_VALUE16 = 0
    DUPLICATES_VOLUME16 = 0
    DUPLICATES_VALUE16 = 0

#
#

matchOVAFound = False
matchINTFound = False
fileFound = False
# ##################################### AIRTEL KR CASHOUT OVA ############################################################
for file in dir_list:
    if file == 'KR AirtelTigo' + yesterday + '.xlsx':
        fileFound = True
        matchOVAFound = True
        match_OVA = 'KR AirtelTigo' + yesterday + '.xlsx'
        wb = load_workbook(match_OVA)
        sheet = wb[wb.sheetnames[0]]
        wb.active = wb[wb.sheetnames[0]]
        file = pd.read_excel(match_OVA)

        first_col = wb.active.min_column
        last_col = wb.active.max_column
        first_row = wb.active.min_row
        last_row = wb.active.max_row

        AIRTEL_KR_Cashout_OVA_Volume = 0
        AIRTEL_KR_Cashout_OVA_Sum = 0.00


        serviceTypeCol = get_column("Service Type")
        amountColumn = get_column('Transaction Amount')
        ova_id_col = get_column('External Transaction Id')

        for i in range(first_row+1, last_row+1):
            if sheet[serviceTypeCol+str(i)].value == "Cash in":
                AIRTEL_KR_Cashout_OVA_Volume = AIRTEL_KR_Cashout_OVA_Volume + 1
                AIRTEL_KR_Cashout_OVA_Sum = AIRTEL_KR_Cashout_OVA_Sum + sheet[amountColumn+str(i)].value

        OVA_VOLUME17 = AIRTEL_KR_Cashout_OVA_Volume
        print('AIRTEL_KR_Cashout_OVA_Volume: ' + str(AIRTEL_KR_Cashout_OVA_Volume))

        print('AIRTEL_KR_Cashout_OVA_Sum: ' + str(AIRTEL_KR_Cashout_OVA_Sum))
        OVA_VALUE17 = AIRTEL_KR_Cashout_OVA_Sum
        wb.close()
        wb.save('KR AirtelTigo Cashout' + yesterday + '- Recons.xlsx')
        match_OVA = 'KR AirtelTigo Cashout' + yesterday + '- Recons.xlsx'

        wb = load_workbook(match_OVA)
        ws1 = wb.create_sheet("Duplicates")
        ws1.title = 'Duplicates'

        ws1 = wb.create_sheet("Missing Integrator Transactions")
        ws1.title = 'Missing Integrator Transactions'

        ws1 = wb.create_sheet("Missing Integrator Transactions")
        ws1.title = 'Missing Integrator Transactions'

        wb.close()
        wb.save(match_OVA)
        break
    else:
        fileFound = False

if fileFound is False:
    OVA_VOLUME17 = 0
    OVA_VALUE17 = 0

# ####################################### AIRTEL KR CASHOUT INT #####################################################
for file in dir_list:
    if file == 'KR AirtelTigo Disb_mBase' + yesterday + '.xlsx':
        fileFound = True
        matchINTFound = True
        match_INT = 'KR AirtelTigo Disb_mBase' + yesterday + '.xlsx'
        wb = load_workbook(match_INT)
        sheet = wb[wb.sheetnames[0]]
        wb.active = wb[wb.sheetnames[0]]
        file = pd.read_excel(match_INT)

        # DEFINE MAX AND MIN COLUMNS AND ROWS
        first_col = wb.active.min_column
        last_col = wb.active.max_column
        first_row = wb.active.min_row
        last_row = wb.active.max_row

        AIRTEL_KR_Cashout_INT_Volume = (last_row + 1) - (first_row + 1)
        INT_VOLUME17 = AIRTEL_KR_Cashout_INT_Volume
        print('AIRTEL_KR_Cashout_INT_Volume: ' + str(AIRTEL_KR_Cashout_INT_Volume))

        AIRTEL_KR_Cashout_INT_Sum = 0.00

        amountColumn = get_column('Amount')
        int_id_col = get_column('Transaction Id')
        for i in range(first_row + 1, last_row + 1):
            AIRTEL_KR_Cashout_INT_Sum = AIRTEL_KR_Cashout_INT_Sum + float(sheet[amountColumn + str(i)].value)

        print('AIRTEL_KR_Cashout_INT_Sum: ' + str(AIRTEL_KR_Cashout_INT_Sum))
        INT_VALUE17 = AIRTEL_KR_Cashout_INT_Sum
        wb.close()
        wb.save(match_INT)
        # ---------------------------------------- DUPLICATES -------------------------------------------------
        wb = load_workbook(match_INT)
        Owb = load_workbook(match_OVA)
        sheet = wb[wb.sheetnames[0]]
        wb.active = wb[wb.sheetnames[0]]
        list = []
        duplicates = []
        count = 1
        counter = 0
        headerChecker = True
        duplicates_value_sum = 0.00
        for id in range(first_row + 1, last_row + 1):
            if sheet[int_id_col + str(id)].value not in list:
                list.append(sheet[int_id_col + str(id)].value)
            elif sheet[int_id_col + str(id)].value not in duplicates:
                    data = file[file['Transaction Id'] == (sheet[int_id_col + str(id)].value)]
                    data_set_sum = data.iloc[1:, data.columns.get_loc("Amount")].sum()
                    duplicates_value_sum = duplicates_value_sum + data_set_sum
                    duplicates.append(sheet[int_id_col + str(id)].value)
                    if counter > 0:
                        headerChecker = False
                    with pd.ExcelWriter(match_OVA, mode='a', engine="openpyxl", if_sheet_exists='overlay') as writer:
                        data.to_excel(writer, sheet_name='Duplicates', startrow=count, header=headerChecker)
                        if headerChecker is False:
                            count += 2
                        else:
                            count += (data.shape[0] +1)
                        counter += 1
            else:
                if sheet[int_id_col + str(id)].value in duplicates:
                    counter += 1
                continue
        wb.close()
        wb.save(match_INT)
        Owb.close()
        # Owb.save(match_OVA)

        wb = load_workbook(match_OVA)
        wb.active = wb['Duplicates']
        dup_sheet = wb['Duplicates']
        sheet = dup_sheet

        first_col = wb.active.min_column
        last_col = wb.active.max_column
        first_row = wb.active.min_row
        last_row = wb.active.max_row

        dup_col = column_index_from_string(amountColumn) + 1

        print(f"Number of duplicates: {counter}")
        dup_sheet[get_column_letter(dup_col) + str(last_row + 4)].value = str(round(duplicates_value_sum, 2))
        dup_sheet[get_column_letter(dup_col) + str(last_row + 5)].value = str(counter)

        print(f"Duplicates: {duplicates}")
        DUPLICATES_VOLUME17 = counter
        DUPLICATES_VALUE17 = duplicates_value_sum
        wb.close()
        wb.save(match_OVA)

        if matchOVAFound is True and matchINTFound is True:
            # ########################################### MISSING INT TRANSACTIONS ###############################################
            Owb = load_workbook(match_OVA)
            Osheet = Owb[Owb.sheetnames[0]]
            Owb.active = Owb[Owb.sheetnames[0]]
            file = pd.read_excel(match_OVA)

            Iwb = load_workbook(match_INT)
            Isheet = Iwb[Iwb.sheetnames[0]]
            Iwb.active = Iwb[Iwb.sheetnames[0]]

            Ofirst_col = Owb.active.min_column
            Olast_col = Owb.active.max_column
            Ofirst_row = Owb.active.min_row
            Olast_row = Owb.active.max_row

            Ifirst_col = Iwb.active.min_column
            Ilast_col = Iwb.active.max_column
            Ifirst_row = Iwb.active.min_row
            Ilast_row = Iwb.active.max_row

            missing_INT_list = []
            missing_OVA_list = []
            matchFound = False
            headerChecker = True
            counter = 0

            for i in range(Ofirst_row + 1, Olast_row + 1):
                if str(Osheet[serviceTypeCol + str(i)].value) == "Cash in":
                    for j in range(Ifirst_row + 1, Ilast_row + 1):
                        if str(Osheet[ova_id_col + str(i)].value) == str(Isheet[int_id_col + str(j)].value):
                            matchFound = True
                            break
                        else:
                            matchFound = False
                    if matchFound is False:
                        missing_INT_list.append(Osheet[ova_id_col + str(i)].value)
            print(f"Missing integrator transactions: {missing_INT_list}")
            count = 0
            for transaction in missing_INT_list:
                data = file[file['External Transaction Id'] == transaction]
                if counter > 0:
                    headerChecker = False
                with pd.ExcelWriter(match_OVA, mode='a', engine="openpyxl", if_sheet_exists='overlay') as writer:
                    data.to_excel(writer, sheet_name='Missing Integrator Transactions', startrow=count,
                                  header=headerChecker)
                    if headerChecker is False:
                        count += 1
                    else:
                        count += 2
                    counter += 1

            Iwb.close()
            Owb.close()

            # ######################################### MISSING OVA TRANSACTIONS ##############################################
            Owb = load_workbook(match_OVA)
            Osheet = Owb[Owb.sheetnames[0]]
            Owb.active = Owb[Owb.sheetnames[0]]

            Iwb = load_workbook(match_INT)
            Isheet = Iwb[Iwb.sheetnames[0]]
            Iwb.active = Iwb[Iwb.sheetnames[0]]
            file = pd.read_excel(match_INT)

            Ofirst_col = Owb.active.min_column
            Olast_col = Owb.active.max_column
            Ofirst_row = Owb.active.min_row
            Olast_row = Owb.active.max_row

            Ifirst_col = Iwb.active.min_column
            Ilast_col = Iwb.active.max_column
            Ifirst_row = Iwb.active.min_row
            Ilast_row = Iwb.active.max_row

            missing_INT_list = []
            missing_OVA_list = []
            matchFound = False
            headerChecker = True
            counter = 0

            for i in range(Ifirst_row + 1, Ilast_row + 1):
                for j in range(Ofirst_row + 1, Olast_row + 1):
                    if str(Osheet[serviceTypeCol + str(j)].value) == "Cash in":
                        if str(Isheet[int_id_col + str(i)].value) == str(Osheet[ova_id_col + str(j)].value):
                            matchFound = True
                            break
                        else:
                            matchFound = False
                    else:
                        continue
                if matchFound is False:
                    missing_OVA_list.append(Isheet[int_id_col + str(i)].value)
            print(f"Missing OVA transactions: {missing_OVA_list}")
            count = 0
            for transaction in missing_OVA_list:
                data = file[file['Transaction Id'] == transaction]
                if counter > 0:
                    headerChecker = False
                with pd.ExcelWriter(match_OVA, mode='a', engine="openpyxl", if_sheet_exists='overlay') as writer:
                    data.to_excel(writer, sheet_name='Missing OVA Transactions', startrow=count, header=headerChecker)
                    if headerChecker is False:
                        count += 1
                    else:
                        count += 2
                    counter += 1
            Iwb.close()
            Owb.close()
        break
    else:
        fileFound = False

if fileFound is False:
    INT_VOLUME17 = 0
    INT_VALUE17 = 0
    DUPLICATES_VOLUME17 = 0
    DUPLICATES_VALUE17 = 0
#
#
fileFound = False
matchOVAFound = False
matchINTFound = False
# ############################################## VODA KR CASHIN OVA ##############################################
for file in dir_list:
    if file == 'KR Vodafone Cashin' + yesterday + '.xlsx':
        fileFound = True
        matchOVAFound = True
        match_OVA = 'KR Vodafone Cashin' + yesterday + '.xlsx'
        wb = load_workbook(match_OVA, read_only=False)
        sheet = wb[wb.sheetnames[0]]
        wb.active = wb[wb.sheetnames[0]]
        ids = []

        # DEFINE MAX AND MIN COLUMNS AND ROWS
        first_col = wb.active.min_column
        last_col = wb.active.max_column
        first_row = wb.active.min_row
        last_row = wb.active.max_row
        amountColumn = get_column('Paid In')
        VODA_CASHIN_KR_OVA_Volume = 0
        for i in range(first_row + 6, last_row + 1):
            if (str(sheet[amountColumn + str(i)].value) == '') or ((sheet[amountColumn + str(i)].value) is None):
                continue
            VODA_CASHIN_KR_OVA_Volume = VODA_CASHIN_KR_OVA_Volume + 1
        print('VODA_CASHIN_KR_OVA_Volume: ' + str(VODA_CASHIN_KR_OVA_Volume))
        OVA_VOLUME18 = VODA_CASHIN_KR_OVA_Volume
        VODA_KR_Cashin_OVA_Sum = 0.00
        for i in range(first_row + 6, last_row + 1):
            if (str(sheet[amountColumn + str(i)].value) == '') or ((sheet[amountColumn + str(i)].value) is None):
                continue
            VODA_KR_Cashin_OVA_Sum = VODA_KR_Cashin_OVA_Sum + float(sheet[amountColumn + str(i)].value)
        ova_id_col = get_column('TransId')
        tmp_id_col = ova_id_col
        alternate_id_col = get_column("Receipt No.")
        for i in range(first_row + 6, last_row + 1):
            if sheet[ova_id_col + str(i)].value is None:
                ids.append((sheet[ova_id_col + str(i)].value))
            else:
                ids.append(str(sheet[ova_id_col + str(i)].value))
        print('VODA_Cashin_KR_OVA_Sum: ' + str(VODA_KR_Cashin_OVA_Sum))
        OVA_VALUE18 = VODA_KR_Cashin_OVA_Sum
        wb.close()
        wb.save('KR Vodafone Cashin' + yesterday + '- Recons.xlsx')
        match_OVA = 'KR Vodafone Cashin' + yesterday + '- Recons.xlsx'

        wb = load_workbook(match_OVA)
        ws1 = wb.create_sheet("Duplicates")
        ws1.title = 'Duplicates'

        ws1 = wb.create_sheet("Missing Integrator Transactions")
        ws1.title = 'Missing Integrator Transactions'

        ws1 = wb.create_sheet("Missing Integrator Transactions")
        ws1.title = 'Missing Integrator Transactions'

        wb.close()
        wb.save(match_OVA)
        break
    else:
        fileFound = False

if fileFound is False:
    OVA_VALUE18 = 0
    OVA_VOLUME18 = 0
#
#
fileFound = False
# ########################################### VODA KR CASHIN INT ######################################################3
for file in dir_list:
    if file == 'KR Vodafone Coll_mBase' + yesterday + '.xlsx':
        fileFound = True
        matchINTFound = True
        match_INT = 'KR Vodafone Coll_mBase' + yesterday + '.xlsx'
        wb = load_workbook(match_INT)
        sheet = wb[wb.sheetnames[0]]
        wb.active = wb[wb.sheetnames[0]]
        file = pd.read_excel(match_INT)

        # DEFINE MAX AND MIN COLUMNS AND ROWS
        first_col = wb.active.min_column
        last_col = wb.active.max_column
        first_row = wb.active.min_row
        last_row = wb.active.max_row

        Vodafone_KR_Cashin_INT_Sum = 0.00
        Vodafone_KR_Cashin_INT_Volume = 0

        statusFlagColumn = get_column('Status')
        amountColumn = get_column('Amount')
        int_id_col = get_column('Transaction Id')
        for i in range(first_row + 1, last_row + 1):
            if (sheet[statusFlagColumn + str(i)].value) == 'CONFIRMED':
                Vodafone_KR_Cashin_INT_Volume += 1

        INT_VOLUME18 = Vodafone_KR_Cashin_INT_Volume
        print('Vodafone_KR_Cashin_INT_Volume: ' + str(Vodafone_KR_Cashin_INT_Volume))

        for i in range(first_row + 1, last_row + 1):
            if (sheet[statusFlagColumn + str(i)].value) == 'CONFIRMED':
                Vodafone_KR_Cashin_INT_Sum = Vodafone_KR_Cashin_INT_Sum + float(sheet[amountColumn + str(i)].value)

        print('Vodafone_KR_Cashin_INT_Sum: ' + str(Vodafone_KR_Cashin_INT_Sum))
        INT_VALUE18 = Vodafone_KR_Cashin_INT_Sum
        wb.close()
        wb.save(match_INT)
        # --------------------------------------------- Duplicates ------------------------------------------------------
        Owb = load_workbook(match_OVA)
        wb = load_workbook(match_INT)
        sheet = wb[wb.sheetnames[0]]
        wb.active = wb[wb.sheetnames[0]]
        list = []
        duplicates = []
        count = 1
        counter = 0
        headerChecker = True
        duplicates_value_sum = 0.00
        for id in range(first_row + 1, last_row + 1):
            if (sheet[statusFlagColumn + str(id)].value) != 'CONFIRMED':
                continue
            else:
                if sheet[int_id_col + str(id)].value not in list:
                    list.append(sheet[int_id_col + str(id)].value)
                elif sheet[int_id_col + str(id)].value not in duplicates:
                        data = file[file['Transaction Id'] == sheet[int_id_col + str(id)].value]
                        data_set_sum = data.iloc[1:, data.columns.get_loc("Amount")].sum()
                        duplicates_value_sum = duplicates_value_sum + data_set_sum
                        duplicates.append(sheet[int_id_col + str(id)].value)
                        if counter > 0:
                            headerChecker = False
                        with pd.ExcelWriter(match_OVA, mode='a', engine="openpyxl", if_sheet_exists='overlay') as writer:
                            data.to_excel(writer, sheet_name='Duplicates', startrow=count, header=headerChecker)
                            if headerChecker is False:
                                count += 2
                            else:
                                count += (data.shape[0]+1)
                            counter += 1
                else:
                    if sheet[int_id_col + str(id)].value in duplicates:
                        counter += 1
                    continue
        wb.close()
        wb.save(match_INT)
        Owb.close
        # wb.save('Vodafone_Cash_Kowri_Send_Money_OVA_trn'+yesterday+'.xlsx')

        wb = load_workbook(match_OVA)
        wb.active = wb['Duplicates']
        dup_sheet = wb['Duplicates']
        sheet = dup_sheet

        first_col = wb.active.min_column
        last_col = wb.active.max_column
        first_row = wb.active.min_row
        last_row = wb.active.max_row

        dup_col = column_index_from_string(amountColumn) + 1

        print(f"Number of duplicates: {counter}")
        dup_sheet[get_column_letter(dup_col) + str(last_row + 4)].value = str(round(duplicates_value_sum, 2))
        dup_sheet[get_column_letter(dup_col) + str(last_row + 5)].value = str(counter)

        print(f"Duplicates: {duplicates}")
        DUPLICATES_VOLUME18 = counter
        DUPLICATES_VALUE18 = duplicates_value_sum

        wb.close()
        wb.save(match_OVA)

        if matchOVAFound is True and matchINTFound is True:
            # ------------------------------------------- MISSING INT TRANSACTIONS ------------------------------------
            Owb = load_workbook(match_OVA)
            Osheet = Owb[Owb.sheetnames[0]]
            Owb.active = Owb[Owb.sheetnames[0]]
            file = pd.read_excel(match_OVA, header=[5])

            Iwb = load_workbook(match_INT)
            Isheet = Iwb[Iwb.sheetnames[0]]
            Iwb.active = Iwb[Iwb.sheetnames[0]]

            Ofirst_col = Owb.active.min_column
            Olast_col = Owb.active.max_column
            Ofirst_row = Owb.active.min_row
            Olast_row = Owb.active.max_row

            Ifirst_col = Iwb.active.min_column
            Ilast_col = Iwb.active.max_column
            Ifirst_row = Iwb.active.min_row
            Ilast_row = Iwb.active.max_row

            missing_INT_list = []
            missing_OVA_list = []
            matchFound = False
            headerChecker = True
            counter = 0

            for i in range(Ofirst_row + 6, Olast_row + 1):
                for j in range(Ifirst_row + 1, Ilast_row + 1):
                    if (Isheet[statusFlagColumn + str(j)].value) == 'CONFIRMED':
                        if Osheet[ova_id_col + str(i)].value is None or Osheet[ova_id_col + str(i)].value == "#VALUE!":
                            ova_id_col = alternate_id_col
                            matchFound = False
                            break
                        if str(Osheet[ova_id_col + str(i)].value) in str(Isheet[int_id_col + str(j)].value):
                            matchFound = True
                            break
                        else:
                            matchFound = False
                if matchFound is False:
                    missing_INT_list.append(str(Osheet[ova_id_col + str(i)].value))
                    ova_id_col = tmp_id_col
            print(f"Missing integrator transactions: {missing_INT_list}")
            count = 0
            for transaction in missing_INT_list:
                data = file[file['TransId'] == transaction]
                if data.empty:
                    data = file[file['Receipt No.'] == int(transaction)]
                if counter > 0:
                    headerChecker = False
                with pd.ExcelWriter(match_OVA, mode='a', engine="openpyxl", if_sheet_exists='overlay') as writer:
                    data.to_excel(writer, sheet_name='Missing Integrator Transactions', startrow=count,
                                  header=headerChecker)
                    if headerChecker is False:
                        count += 1
                    else:
                        count += 2
                    counter += 1

            Iwb.close()
            Owb.close()

            # ----------------------------------------- MISSING OVA TRANSACTIONS ---------------------------------------
            Owb = load_workbook(match_OVA)
            Osheet = Owb[Owb.sheetnames[0]]
            Owb.active = Owb[Owb.sheetnames[0]]

            Iwb = load_workbook(match_INT)
            Isheet = Iwb[Iwb.sheetnames[0]]
            Iwb.active = Iwb[Iwb.sheetnames[0]]
            file = pd.read_excel(match_INT)

            Ofirst_col = Owb.active.min_column
            Olast_col = Owb.active.max_column
            Ofirst_row = Owb.active.min_row
            Olast_row = Owb.active.max_row

            Ifirst_col = Iwb.active.min_column
            Ilast_col = Iwb.active.max_column
            Ifirst_row = Iwb.active.min_row
            Ilast_row = Iwb.active.max_row

            missing_INT_list = []
            missing_OVA_list = []
            matchFound = False
            headerChecker = True
            counter = 0

            for i in range(Ifirst_row + 1, Ilast_row + 1):
                if (Isheet[statusFlagColumn + str(i)].value) != 'CONFIRMED':
                    break
                else:
                    for j in ids:
                        if str(Isheet[int_id_col + str(i)].value) in str(j):
                            matchFound = True
                            break
                        else:
                            matchFound = False
                    if matchFound is False:
                        missing_OVA_list.append(Isheet[int_id_col + str(i)].value)
            print(f"Missing OVA transactions: {missing_OVA_list}")
            count = 0
            for transaction in missing_OVA_list:
                data = file[file['Transaction Id'] == transaction]
                if counter > 0:
                    headerChecker = False
                with pd.ExcelWriter(match_OVA, mode='a', engine="openpyxl", if_sheet_exists='overlay') as writer:
                    data.to_excel(writer, sheet_name='Missing OVA Transactions', startrow=count, header=headerChecker)
                    if headerChecker is False:
                        count += 1
                    else:
                        count += 2
                    counter += 1
            Iwb.close()
            Owb.close()
        break
    else:
        fileFound = False
if fileFound is False:
    INT_VOLUME18 = 0
    INT_VALUE18 = 0
    DUPLICATES_VOLUME18 = 0
    DUPLICATES_VALUE18 = 0

#
fileFound = False
matchOVAFound = False
matchINTFound = False
#
# ######################################### VODA KR CASHOUT OVA #####################################################
for file in dir_list:
    if file == 'KR Vodafone Cashout' + yesterday + '.xlsx':
        fileFound = True
        matchOVAFound = True
        match_OVA = 'KR Vodafone Cashout' + yesterday + '.xlsx'
        wb = load_workbook(match_OVA, read_only=False, data_only=True)
        sheet = wb[wb.sheetnames[0]]
        wb.active = wb[wb.sheetnames[0]]
        sheetName = wb.sheetnames[0]

        # DEFINE MAX AND MIN COLUMNS AND ROWS
        first_col = wb.active.min_column
        last_col = wb.active.max_column
        first_row = wb.active.min_row
        last_row = wb.active.max_row
        amountColumn = get_column('Withdrawn')
        VODA_CASHOUT_KR_OVA_Volume = 0
        for i in range(first_row + 6, last_row + 1):
            if (str(sheet[amountColumn + str(i)].value) == '') or ((sheet[amountColumn + str(i)].value) is None):
                continue
            VODA_CASHOUT_KR_OVA_Volume = VODA_CASHOUT_KR_OVA_Volume + 1
        OVA_VOLUME19 = VODA_CASHOUT_KR_OVA_Volume
        print('VODA_CASHOUT_OVA_Volume: ' + str(VODA_CASHOUT_KR_OVA_Volume))

        VODA_KR_Cashout_OVA_Sum = 0.00

        ids = []

        ova_id_col = get_column('TransId')
        tmp_id_col = ova_id_col
        alternate_id_col = get_column("Receipt No.")

        for i in range(first_row + 6, last_row + 1):
            if ((sheet[amountColumn + str(i)].value) == '') or ((sheet[amountColumn + str(i)].value) is None):
                continue
            VODA_KR_Cashout_OVA_Sum = VODA_KR_Cashout_OVA_Sum + abs(float(sheet[amountColumn + str(i)].value))
            ids.append(str(sheet[ova_id_col + str(i)].value))

        print('VODA_Cashout_OVA_Sum: ' + str(VODA_KR_Cashout_OVA_Sum))
        OVA_VALUE19 = VODA_KR_Cashout_OVA_Sum
        wb.close()
        wb.save('KR Vodafone Cashout' + yesterday + '- Recons.xlsx')
        match_OVA = 'KR Vodafone Cashout' + yesterday + '- Recons.xlsx'

        wb = load_workbook(match_OVA)
        ws1 = wb.create_sheet("Duplicates")
        ws1.title = 'Duplicates'

        ws1 = wb.create_sheet("Missing Integrator Transactions")
        ws1.title = 'Missing Integrator Transactions'

        ws1 = wb.create_sheet("Missing Integrator Transactions")
        ws1.title = 'Missing Integrator Transactions'

        wb.close()
        wb.save(match_OVA)
        break
    else:
        fileFound = False

if fileFound is False:
    OVA_VOLUME19 = 0
    OVA_VALUE19 = 0

#
#
fileFound = False
# ############################################ VODA KR CASHOUT INT #################################################
for file in dir_list:
    if file == 'KR Vodafone Disb_mBase' + yesterday + '.xlsx':
        fileFound = True
        matchINTFound = True
        match_INT = 'KR Vodafone Disb_mBase' + yesterday + '.xlsx'
        wb = load_workbook(match_INT)
        file = pd.read_excel(match_INT)
        sheet = wb[wb.sheetnames[0]]
        wb.active = wb[wb.sheetnames[0]]

        # DEFINE MAX AND MIN COLUMNS AND ROWS
        first_col = wb.active.min_column
        last_col = wb.active.max_column
        first_row = wb.active.min_row
        last_row = wb.active.max_row

        Vodafone_KR_Cashout_INT_Sum = 0.00
        Vodafone_KR_Cashout_INT_Volume = 0
        statusFlagColumn = get_column("Status")
        Vodafone_KR_Cashout_INT_Volume = 0
        for i in range(first_row + 1, last_row + 1):
            if sheet[statusFlagColumn + str(i)].value == "CONFIRMED":
                Vodafone_KR_Cashout_INT_Volume = Vodafone_KR_Cashout_INT_Volume + 1
        INT_VOLUME19 = Vodafone_KR_Cashout_INT_Volume
        print('Vodafone_KR_Cashout_INT_Volume: ' + str(Vodafone_KR_Cashout_INT_Volume))

        amountColumn = get_column('Amount')
        int_id_col = get_column('Transaction Id')
        for i in range(first_row + 1, last_row + 1):
            if sheet[statusFlagColumn + str(i)].value == "CONFIRMED":
                Vodafone_KR_Cashout_INT_Sum = Vodafone_KR_Cashout_INT_Sum + float(sheet[amountColumn + str(i)].value)
            else:
                continue
        print('Vodafone_KR_Cashout_INT_Sum: ' + str(Vodafone_KR_Cashout_INT_Sum))
        INT_VALUE19 = Vodafone_KR_Cashout_INT_Sum
        wb.close()
        wb.save(match_INT)
        # ------------------------------------------ Duplicates ---------------------------------------------------------
        Owb = load_workbook(match_OVA)
        wb = load_workbook(match_INT)
        sheet = wb[wb.sheetnames[0]]
        wb.active = wb[wb.sheetnames[0]]
        list = []
        duplicates = []
        count = 1
        counter = 0
        headerChecker = True
        duplicates_value_sum = 0.00
        for i in range(first_row + 1, last_row + 1):
            if sheet[int_id_col + str(i)].value not in list:
                list.append(sheet[int_id_col + str(i)].value)
            elif sheet[int_id_col + str(i)].value not in duplicates:
                    data = file[file['Transaction Id'] == sheet[int_id_col + str(i)].value]
                    data_set_sum = data.iloc[1:, data.columns.get_loc("Amount")].sum()
                    duplicates_value_sum = duplicates_value_sum + data_set_sum
                    duplicates.append(sheet[int_id_col + str(i)].value)
                    if counter > 0:
                        headerChecker = False
                    with pd.ExcelWriter(match_OVA, mode='a', engine="openpyxl", if_sheet_exists='overlay') as writer:
                        data.to_excel(writer, sheet_name='Duplicates', startrow=count, header=headerChecker)
                        if headerChecker is False:
                            count += 2
                        else:
                            count += (data.shape[0] +1)
                        counter += 1
            else:
                if sheet[int_id_col + str(i)].value in duplicates:
                    counter += 1
                    continue

        wb.close()
        wb.save(match_INT)
        Owb.close()
        # Owb.save(match_OVA)

        wb = load_workbook(match_OVA)
        wb.active = wb['Duplicates']
        dup_sheet = wb['Duplicates']
        sheet = dup_sheet

        first_col = wb.active.min_column
        last_col = wb.active.max_column
        first_row = wb.active.min_row
        last_row = wb.active.max_row

        dup_col = column_index_from_string(amountColumn) + 1

        print(f"Number of duplicates: {counter}")
        dup_sheet[get_column_letter(dup_col) + str(last_row + 4)].value = str(round(duplicates_value_sum, 2))
        dup_sheet[get_column_letter(dup_col) + str(last_row + 5)].value = str(counter)

        print(f"Duplicates: {duplicates}")
        DUPLICATES_VOLUME19 = counter
        DUPLICATES_VALUE19 = duplicates_value_sum

        wb.close()
        wb.save(match_OVA)

        if matchOVAFound is True and matchINTFound is True:
            # ----------------------------------- MISSING INT TRANSACTIONS ---------------------------------------------
            Owb = load_workbook(match_OVA, data_only=True)
            Osheet = Owb[Owb.sheetnames[0]]
            Owb.active = Owb[Owb.sheetnames[0]]
            file = pd.read_excel(match_OVA, header=[5])

            Iwb = load_workbook(match_INT)
            Isheet = Iwb[Iwb.sheetnames[0]]
            Iwb.active = Iwb[Iwb.sheetnames[0]]

            Ofirst_col = Owb.active.min_column
            Olast_col = Owb.active.max_column
            Ofirst_row = Owb.active.min_row
            Olast_row = Owb.active.max_row

            Ifirst_col = Iwb.active.min_column
            Ilast_col = Iwb.active.max_column
            Ifirst_row = Iwb.active.min_row
            Ilast_row = Iwb.active.max_row

            missing_INT_list = []
            missing_OVA_list = []
            matchFound = False
            headerChecker = True
            counter = 0

            for i in range(Ofirst_row + 6, Olast_row + 1):
                for j in range(Ifirst_row + 1, Ilast_row + 1):
                    if (Isheet[statusFlagColumn + str(j)].value) == 'CONFIRMED':
                        if Osheet[ova_id_col + str(i)].value is None or Osheet[ova_id_col + str(i)].value == "#VALUE!":
                            ova_id_col = alternate_id_col
                            matchFound = False
                            break
                        if str(Osheet[ova_id_col + str(i)].value) in str(Isheet[int_id_col + str(j)].value):
                            matchFound = True
                            break
                        else:
                            matchFound = False
                if matchFound is False:
                    missing_INT_list.append(str(Osheet[ova_id_col + str(i)].value))
                    ova_id_col = tmp_id_col
            print(f"Missing integrator transactions: {missing_INT_list}")
            count = 0
            for transaction in missing_INT_list:
                data = file[file['TransId'] == transaction]
                if data.empty:
                    data = file[file['Receipt No.'] == int(transaction)]
                if counter > 0:
                    headerChecker = False
                with pd.ExcelWriter(match_OVA, mode='a', engine="openpyxl", if_sheet_exists='overlay') as writer:
                    data.to_excel(writer, sheet_name='Missing Integrator Transactions', startrow=count,
                                  header=headerChecker)
                    if headerChecker is False:
                        count += 1
                    else:
                        count += 2
                    counter += 1
            Iwb.close()
            Owb.close()

            # ------------------------------------------ MISSING OVA TRANSACTIONS ---------------------------------------
            Owb = load_workbook(match_OVA)
            Osheet = Owb[Owb.sheetnames[0]]
            Owb.active = Owb[Owb.sheetnames[0]]

            Iwb = load_workbook(match_INT)
            Isheet = Iwb[Iwb.sheetnames[0]]
            Iwb.active = Iwb[Iwb.sheetnames[0]]
            file = pd.read_excel(match_INT)

            Ofirst_col = Owb.active.min_column
            Olast_col = Owb.active.max_column
            Ofirst_row = Owb.active.min_row
            Olast_row = Owb.active.max_row

            Ifirst_col = Iwb.active.min_column
            Ilast_col = Iwb.active.max_column
            Ifirst_row = Iwb.active.min_row
            Ilast_row = Iwb.active.max_row

            missing_INT_list = []
            missing_OVA_list = []
            matchFound = False
            headerChecker = True
            counter = 0

            for i in range(Ifirst_row + 1, Ilast_row + 1):
                for id in ids:
                    if str(Isheet[int_id_col + str(i)].value) in str(id):
                        matchFound = True
                        break
                    else:
                        matchFound = False
                if matchFound is False:
                    missing_OVA_list.append(Isheet[int_id_col + str(i)].value)
            print(f"Missing OVA transactions: {missing_OVA_list}")
            count = 0
            for transaction in missing_OVA_list:
                data = file[file['Transaction Id'] == transaction]
                if counter > 0:
                    headerChecker = False
                with pd.ExcelWriter(match_OVA, mode='a', engine="openpyxl", if_sheet_exists='overlay') as writer:
                    data.to_excel(writer, sheet_name='Missing OVA Transactions', startrow=count, header=headerChecker)
                    if headerChecker is False:
                        count += 1
                    else:
                        count += 2
                    counter += 1
            Iwb.close()
            Owb.close()
        break
    else:
        fileFound = False

if fileFound is False:
    INT_VOLUME19 = 0
    INT_VALUE19 = 0
    DUPLICATES_VOLUME19 = 0
    DUPLICATES_VALUE19 = 0
#
#
#

# VARIANCE COUNTS & VALUES
VARIANCE_VOLUME1 = abs(INT_VOLUME1 - OVA_VOLUME1)
VARIANCE_VOLUME2 = abs(INT_VOLUME2 - OVA_VOLUME2)
VARIANCE_VOLUME3 = abs(INT_VOLUME3 - OVA_VOLUME3)
VARIANCE_VOLUME4 = abs(INT_VOLUME4 - OVA_VOLUME4)
VARIANCE_VOLUME5 = abs(INT_VOLUME5 - OVA_VOLUME5)
VARIANCE_VOLUME6 = abs(INT_VOLUME6 - OVA_VOLUME6)
VARIANCE_VOLUME7 = abs(INT_VOLUME7 - OVA_VOLUME7)
VARIANCE_VOLUME8 = abs(INT_VOLUME8 - OVA_VOLUME8)
VARIANCE_VOLUME9 = abs(INT_VOLUME9 - OVA_VOLUME9)
VARIANCE_VOLUME10 = abs(INT_VOLUME10 - OVA_VOLUME10)
VARIANCE_VOLUME11 = abs(INT_VOLUME11 - OVA_VOLUME11)
VARIANCE_VOLUME12 = abs(INT_VOLUME12 - OVA_VOLUME12)
VARIANCE_VOLUME13 = abs(INT_VOLUME13 - OVA_VOLUME13)
VARIANCE_VOLUME14 = abs(INT_VOLUME14 - OVA_VOLUME14)
VARIANCE_VOLUME15 = abs(INT_VOLUME15 - OVA_VOLUME15)
VARIANCE_VOLUME16 = abs(INT_VOLUME16 - OVA_VOLUME16)
VARIANCE_VOLUME17 = abs(INT_VOLUME17 - OVA_VOLUME17)
VARIANCE_VOLUME18 = abs(INT_VOLUME18 - OVA_VOLUME18)
VARIANCE_VOLUME19 = abs(INT_VOLUME19 - OVA_VOLUME19)

VARIANCE_VOLUME_LIST1 = [VARIANCE_VOLUME1, VARIANCE_VOLUME2, VARIANCE_VOLUME3, VARIANCE_VOLUME4, VARIANCE_VOLUME5,
                         VARIANCE_VOLUME6, VARIANCE_VOLUME7, VARIANCE_VOLUME8, VARIANCE_VOLUME9, VARIANCE_VOLUME10,
                         VARIANCE_VOLUME11]
VARIANCE_VOLUME_LIST2 = [VARIANCE_VOLUME12, VARIANCE_VOLUME13, VARIANCE_VOLUME14, VARIANCE_VOLUME15, VARIANCE_VOLUME16,
                         VARIANCE_VOLUME17, VARIANCE_VOLUME18, VARIANCE_VOLUME19]

VARIANCE_VALUE1 = abs(INT_VALUE1 - OVA_VALUE1)
VARIANCE_VALUE2 = abs(INT_VALUE2 - OVA_VALUE2)
VARIANCE_VALUE3 = abs(INT_VALUE3 - OVA_VALUE3)
VARIANCE_VALUE4 = abs(INT_VALUE4 - OVA_VALUE4)
VARIANCE_VALUE5 = abs(INT_VALUE5 - OVA_VALUE5)
VARIANCE_VALUE6 = abs(INT_VALUE6 - OVA_VALUE6)
VARIANCE_VALUE7 = abs(INT_VALUE7 - OVA_VALUE7)
VARIANCE_VALUE8 = abs(INT_VALUE8 - OVA_VALUE8)
VARIANCE_VALUE9 = abs(INT_VALUE9 - OVA_VALUE9)
VARIANCE_VALUE10 = abs(INT_VALUE10 - OVA_VALUE10)
VARIANCE_VALUE11 = abs(INT_VALUE11 - OVA_VALUE11)
VARIANCE_VALUE12 = abs(INT_VALUE12 - OVA_VALUE12)
VARIANCE_VALUE13 = abs(INT_VALUE13 - OVA_VALUE13)
VARIANCE_VALUE14 = abs(INT_VALUE14 - OVA_VALUE14)
VARIANCE_VALUE15 = abs(INT_VALUE15 - OVA_VALUE15)
VARIANCE_VALUE16 = abs(INT_VALUE16 - OVA_VALUE16)
VARIANCE_VALUE17 = abs(INT_VALUE17 - OVA_VALUE17)
VARIANCE_VALUE18 = abs(INT_VALUE18 - OVA_VALUE18)
VARIANCE_VALUE19 = abs(INT_VALUE19 - OVA_VALUE19)

VARIANCE_VALUE_LIST1 = [VARIANCE_VALUE1, VARIANCE_VALUE2, VARIANCE_VALUE3, VARIANCE_VALUE4, VARIANCE_VALUE5,
                        VARIANCE_VALUE6, VARIANCE_VALUE7, VARIANCE_VALUE8, VARIANCE_VALUE9, VARIANCE_VALUE10,
                        VARIANCE_VALUE11]
VARIANCE_VALUE_LIST2 = [VARIANCE_VALUE12, VARIANCE_VALUE13, VARIANCE_VALUE14, VARIANCE_VALUE15, VARIANCE_VALUE16,
                        VARIANCE_VALUE17, VARIANCE_VALUE18, VARIANCE_VALUE19]

# update sheet
fwb = load_workbook('Reconciliations 2023.xlsx')

if date.today().day == 1:
    today = datetime.today()
    prev_month = today.replace(month=today.month - 1) if today.month > 1 else today.replace(month=12,year=today.year - 1)
    prev_month_abbr = prev_month.strftime("%b").upper()
    fsheet = fwb[prev_month_abbr]
    fwb.active = fwb[prev_month_abbr]
else:
    fsheet = fwb[current_month]
    fwb.active = fwb[current_month]

first_col = fwb.active.min_column
last_col = fwb.active.max_column
first_row = fwb.active.min_row
last_row = fwb.active.max_row

start_row = 0
OVA_VOLUME_LIST1 = [OVA_VOLUME1, OVA_VOLUME2, OVA_VOLUME3, OVA_VOLUME4, OVA_VOLUME5, OVA_VOLUME6, OVA_VOLUME7,
                    OVA_VOLUME8, OVA_VOLUME9, OVA_VOLUME10, OVA_VOLUME11]
for row in range(first_row + 1, last_row + 1):
    if str(fsheet['A' + str(row)].value) == recons_yesterday:
        start_row = row
        break

i = start_row
for volume in OVA_VOLUME_LIST1:
    fsheet['E' + str(i)].value = volume
    i = i + 1

OVA_VOLUME_LIST2 = [OVA_VOLUME12, OVA_VOLUME13, OVA_VOLUME14, OVA_VOLUME15, OVA_VOLUME16, OVA_VOLUME17, OVA_VOLUME18,
                    OVA_VOLUME19]
i = i + 10
for volume in OVA_VOLUME_LIST2:
    fsheet['E' + str(i)].value = volume
    i = i + 1

OVA_VALUE_LIST1 = [OVA_VALUE1, OVA_VALUE2, OVA_VALUE3, OVA_VALUE4, OVA_VALUE5, OVA_VALUE6, OVA_VALUE7, OVA_VALUE8,
                   OVA_VALUE9, OVA_VALUE10, OVA_VALUE11]
OVA_VALUE_LIST2 = [OVA_VALUE12, OVA_VALUE13, OVA_VALUE14, OVA_VALUE15, OVA_VALUE16, OVA_VALUE17, OVA_VALUE18,
                   OVA_VALUE19]

i = start_row
for value in OVA_VALUE_LIST1:
    fsheet['F' + str(i)].value = value
    i = i + 1

i = i + 10
for value in OVA_VALUE_LIST2:
    fsheet['F' + str(i)].value = value
    i = i + 1

INT_VOLUME_LIST1 = [INT_VOLUME1, INT_VOLUME2, INT_VOLUME3, INT_VOLUME4, INT_VOLUME5, INT_VOLUME6, INT_VOLUME7,
                    INT_VOLUME8, INT_VOLUME9, INT_VOLUME10, INT_VOLUME11]
INT_VOLUME_LIST2 = [INT_VOLUME12, INT_VOLUME13, INT_VOLUME14, INT_VOLUME15, INT_VOLUME16, INT_VOLUME17, INT_VOLUME18,
                    INT_VOLUME19]

i = start_row
for volume in INT_VOLUME_LIST1:
    fsheet['G' + str(i)].value = volume
    i += 1

i = i + 10
for volume in INT_VOLUME_LIST2:
    fsheet['G' + str(i)].value = volume
    i += 1

INT_VALUE_LIST1 = [INT_VALUE1, INT_VALUE2, INT_VALUE3, INT_VALUE4, INT_VALUE5, INT_VALUE6, INT_VALUE7, INT_VALUE8,
                   INT_VALUE9, INT_VALUE10, INT_VALUE11]
INT_VALUE_LIST2 = [INT_VALUE12, INT_VALUE13, INT_VALUE14, INT_VALUE15, INT_VALUE16, INT_VALUE17, INT_VALUE18,
                   INT_VALUE19]

i = start_row
for value in INT_VALUE_LIST1:
    fsheet['H' + str(i)].value = value
    i += 1

i = i + 10
for value in INT_VALUE_LIST2:
    fsheet['H' + str(i)].value = value
    i += 1

i = start_row
for count in VARIANCE_VOLUME_LIST1:
    fsheet['I' + str(i)].value = count
    i += 1

i = i + 10
for count in VARIANCE_VOLUME_LIST2:
    fsheet['I' + str(i)].value = count
    i += 1

i = start_row
for value in VARIANCE_VALUE_LIST1:
    fsheet['J' + str(i)].value = value
    i += 1



i = i + 10
for value in VARIANCE_VALUE_LIST2:
    fsheet['J' + str(i)].value = value
    i += 1

DUPLICATE_VOLUME_LIST1 = [DUPLICATES_VOLUME1, DUPLICATES_VOLUME2, DUPLICATES_VOLUME3, DUPLICATES_VOLUME4,
                          DUPLICATES_VOLUME5, DUPLICATES_VOLUME6, DUPLICATES_VOLUME7, DUPLICATES_VOLUME8,
                          DUPLICATES_VOLUME9, DUPLICATES_VOLUME10, DUPLICATES_VOLUME11]
DUPLICATE_VOLUME_LIST2 = [DUPLICATES_VOLUME12, DUPLICATES_VOLUME13, DUPLICATES_VOLUME14, DUPLICATES_VOLUME15,
                          DUPLICATES_VOLUME16, DUPLICATES_VOLUME17, DUPLICATES_VOLUME19]

i = start_row
for value in DUPLICATE_VOLUME_LIST1:
    fsheet['M' + str(i)].value = value
    i += 1

i = i + 10
for value in DUPLICATE_VOLUME_LIST2:
    fsheet['M' + str(i)].value = value
    i += 1

DUPLICATE_VALUE_LIST1 = [DUPLICATES_VALUE1, DUPLICATES_VALUE2, DUPLICATES_VALUE3, DUPLICATES_VALUE4, DUPLICATES_VALUE5,
                         DUPLICATES_VALUE6, DUPLICATES_VALUE7, DUPLICATES_VALUE8, DUPLICATES_VALUE9, DUPLICATES_VALUE10,
                         DUPLICATES_VALUE11]
DUPLICATE_VALUE_LIST2 = [DUPLICATES_VALUE12, DUPLICATES_VALUE13, DUPLICATES_VALUE14, DUPLICATES_VALUE15,
                         DUPLICATES_VALUE16, DUPLICATES_VALUE17, DUPLICATES_VALUE19]

i = start_row
for value in DUPLICATE_VALUE_LIST1:
    fsheet['N' + str(i)].value = value
    i += 1

i = i + 10
for value in DUPLICATE_VALUE_LIST2:
    fsheet['N' + str(i)].value = value
    i += 1

fwb.close()
fwb.save('Reconciliations 2023.xlsx')

