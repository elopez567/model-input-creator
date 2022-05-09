import pyodbc
import numpy as np
import pandas as pd
import tkinter as tk
from tkinter import *
from tkinter import filedialog
from openpyxl import load_workbook
from openpyxl.workbook import Workbook
from datetime import date


class Mainapp:
    def __init__(self):
        app = Tk()
        app.title('BestExInputCreator v1.0')
        app.geometry("300x180")
        app['background'] = '#003087'

        # Labels
        file_label = Label(app, bg='#003087', fg='white', text="Showout File")

        space = Label(app, bg='#003087', text="      ")
        space2 = Label(app, bg='#003087', text="      ")
        space3 = Label(app, bg='#003087', text="      ")
        space4 = Label(app, bg='#003087', text="      ")
        space5 = Label(app, bg='#003087', text="      ")
        space6 = Label(app, bg='#003087', text="      ")

        filedir = StringVar()
        file_dir = Entry(app, textvariable=filedir, width=25, borderwidth=3).grid(row=8, column=3, columnspan=5,
                                                                                  sticky=tk.W + tk.E)

        # Grids
        space.grid(row=0, column=0)
        space2.grid(row=9, column=0)
        space3.grid(row=11, column=0)
        file_label.grid(row=7, column=5)

        today = date.today()
        day = today.strftime("%m/%d/%Y")
        month = day[:2]
        month = int(month) - 1
        self.orig_date = f'{month}/1/2022'

        def get_sheetnames_xlsx(filepath):
            wb = load_workbook(filepath, read_only=True, keep_links=False)
            return wb.sheetnames

        # Open File Dialog
        def openfile():
            app.filename = filedialog.askopenfilename(initialdir=r"M:\Capital Markets\Investor Loan Analysis\Ritz",
                                                      title="Select Deal Tape")
            self.sheet_names = get_sheetnames_xlsx(app.filename)

            filedir = StringVar()
            filedir.set(app.filename)
            file_dir = Entry(app, textvariable=filedir, width=25, borderwidth=3).grid(row=8, column=3, columnspan=5,
                                                                                      sticky=tk.W + tk.E)

            options = self.sheet_names
            self.clicked = StringVar()
            self.clicked.set('Select Sheet')
            sheet_menu = OptionMenu(app, self.clicked, *options).grid(row=11, column=5)

        # Create Strat Function
        def runstrat():
            self.df = pd.read_excel(io=app.filename, skiprows=1, sheet_name=self.clicked.get())
            self.df2 = pd.read_excel(
                io=r'M:\Capital Markets\Users\Emmanuel Lopez\ADCO\ADCO Templates\ADCO Input Template.xlsx')
            self.df3 = pd.read_excel(io=r"M:\Capital Markets\Users\Emmanuel Lopez\Moody's MILAN\ExampleDeal.xlsx")

            self.df2 = self.df2.drop([0, 0])
            self.df2['Loan_ID'] = self.df['Loan']
            self.df2['Original_Term'] = self.df['TERM']
            self.df2['Age'] = 0
            self.df2['Remaining_Term'] = self.df['TERM']
            self.df2['Lien_Position'] = 1
            self.df2['OrigRate'] = self.df['Rate']
            self.df2['CurrentRate'] = self.df['Rate']
            self.df2['Current_Loan_Size'] = self.df['UPB']
            self.df2['Orig_Loan_Size'] = self.df['UPB']
            self.df2['Orig_LTV'] = self.df['LTV']
            self.df2['Orig_Total_LTV'] = self.df['LTV']
            self.df2['FICO_Score'] = self.df['Fico']
            self.df2['RateFixed'] = 1
            self.df2['Balloon_Months'] = 0
            self.df2['PP_Months'] = 0
            self.df2['IO_Months'] = 0
            self.df2['Delinquency'] = 'C'
            self.df2['First_Reset_Age'] = -1
            self.df2['Months_Between_Reset'] = -1
            self.df2['Index'] = -1
            self.df2['Gross_Margin'] = -1
            self.df2['Life_Cap'] = -1
            self.df2['Life_Floor'] = -1
            self.df2['Periodic_Cap'] = -1
            self.df2['Periodic_Floor'] = -1
            self.df2['Payment_Change_Cap'] = -1
            self.df2['Payment_Reset_Period'] = -1
            self.df2['Payment_Recast_Period'] = -1
            self.df2['Current_Minimum_Payment'] = -1
            self.df2['Max_Negam_Percent'] = -1
            self.df2['Documentation'] = 'F'
            self.df2['MI_Cutoff'] = -1
            self.df2['MI_Premium'] = -1

            self.df2['Occupancy'] = self.df['Occupancy']
            self.df2['LoanPurpose'] = self.df['Purpose']
            self.df2['PropertyType'] = self.df['PropType']
            self.df2['Num_Units'] = self.df['PropType']
            self.df2['State'] = self.df['STATE'].apply(lambda x: x.strip())
            self.df2['Zip'] = self.df['Zip']
            # Deal Name
            self.df2['Group_Number'] = 1
            self.df2['tunestr'] = 'FNMA'
            self.df2['PrevDel'] = 0
            self.df2['Months_Since_Delinq'] = -1
            self.df2['Bankrupt'] = -1
            self.df2['ServicingFee'] = -1

            self.df2['MI_Percent'] = 0
            self.df2.loc[(self.df2['Orig_LTV'] > 80) & (self.df2['Orig_LTV'] <= 85), 'MI_Percent'] = 12
            self.df2.loc[(self.df2['Orig_LTV'] > 85) & (self.df2['Orig_LTV'] <= 90), 'MI_Percent'] = 25
            self.df2.loc[(self.df2['Orig_LTV'] > 90) & (self.df2['Orig_LTV'] <= 95), 'MI_Percent'] = 30
            self.df2.loc[(self.df2['Orig_LTV'] > 95), 'MI_Percent'] = 35

            self.df2.loc[(self.df2['PropertyType'] != '2 Unit') & (self.df2['PropertyType'] != '3 Unit') & (
                        self.df2['PropertyType'] != '4 Unit'), 'Num_Units'] = 1
            self.df2.loc[(self.df2['PropertyType'] == '2 Unit'), 'Num_Units'] = 2
            self.df2.loc[(self.df2['PropertyType'] == '3 Unit'), 'Num_Units'] = 3
            self.df2.loc[(self.df2['PropertyType'] == '4 Unit'), 'Num_Units'] = 4

            self.df2.loc[(self.df2['LoanPurpose'] == 'PURCH'), 'LoanPurpose'] = 'P'
            self.df2.loc[(self.df2['LoanPurpose'] == 'REFINANCE - RATE/TERM'), 'LoanPurpose'] = 'R'
            self.df2.loc[(self.df2['LoanPurpose'] == 'REFINANCE - CASHOUT'), 'LoanPurpose'] = 'E'
            self.df2.loc[(self.df2['PropertyType'] == 'MFD'), 'PropertyType'] = 'MH'
            self.df2.loc[(self.df2['PropertyType'] == '2 Unit'), 'PropertyType'] = 'MFR'
            self.df2.loc[(self.df2['PropertyType'] == '3 Unit'), 'PropertyType'] = 'MFR'
            self.df2.loc[(self.df2['PropertyType'] == '4 Unit'), 'PropertyType'] = 'MFR'
            self.df2.loc[(self.df2['Occupancy'] == 'OWN'), 'Occupancy'] = 'O'
            self.df2.loc[(self.df2['Occupancy'] == '2ND'), 'Occupancy'] = 'S'
            self.df2.loc[(self.df2['Occupancy'] == 'NOO'), 'Occupancy'] = 'I'

            # MOODY'S
            self.df3 = self.df3.drop([0, 0])

            def left(var):
                return str(var)[0]

            self.df3['Loan Number'] = self.df['Loan']
            self.df3['Current Loan Amount'] = self.df['UPB']
            self.df3['Amortization Type'] = 1
            self.df3.loc[(self.df2['LoanPurpose'] == 'P'), 'Loan Purpose'] = 7
            self.df3.loc[(self.df2['LoanPurpose'] == 'R'), 'Loan Purpose'] = 9
            self.df3.loc[(self.df2['LoanPurpose'] == 'E'), 'Loan Purpose'] = 3

            self.df3['Left'] = self.df['Loan'].apply(left)

            self.df3.loc[(self.df3['Left'] == '7'), 'Channel'] = 1
            self.df3.loc[(self.df3['Left'] == '6'), 'Channel'] = 2
            self.df3.loc[(self.df3['Left'] == '8'), 'Channel'] = 3

            self.df3['Origination Date'] = self.orig_date
            self.df3['Original Interest Rate'] = self.df['Rate'].apply(lambda x: x / 100)
            self.df3['Original Amortization Term'] = self.df['TERM']
            self.df3['Original Term to Maturity'] = self.df['TERM']
            self.df3['Original Interest Only Term'] = 0
            self.df3['Prepayment Penalty Total Term'] = 0
            self.df3['FICO'] = self.df['Fico']
            self.df3['Originator DTI'] = self.df['DTI'].apply(lambda x: x / 100)
            self.df3['Postal Code'] = self.df['Zip']

            self.df3.loc[(self.df['PropType'] == 'SFR'), 'Property Type'] = 1
            self.df3.loc[(self.df['PropType'] == 'MFD'), 'Property Type'] = 6
            self.df3.loc[(self.df['PropType'] == 'Condo'), 'Property Type'] = 3
            self.df3.loc[(self.df['PropType'] == 'PUD'), 'Property Type'] = 6
            self.df3.loc[(self.df['PropType'] == '2 Unit'), 'Property Type'] = 13
            self.df3.loc[(self.df['PropType'] == '3 Unit'), 'Property Type'] = 14
            self.df3.loc[(self.df['PropType'] == '4 Unit'), 'Property Type'] = 15

            self.df3.loc[(self.df2['Occupancy'] == 'O'), 'Occupancy'] = 1
            self.df3.loc[(self.df2['Occupancy'] == 'S'), 'Occupancy'] = 2
            self.df3.loc[(self.df2['Occupancy'] == 'I'), 'Occupancy'] = 3
            self.df3['Original Appraised Property Value'] = self.df['UPB'] / (self.df['LTV'].apply(lambda x: x / 100))
            self.df3['Original CLTV'] = self.df['LTV'].apply(lambda x: x / 100)
            self.df3['Original LTV'] = self.df3['Original CLTV']
            self.df3['LGD Model Type'] = 'PrivateLabel'
            self.df3['PD Model Type'] = 'PrivateLabel'
            self.df3['Pool ID'] = 1
            self.df3['Documentation'] = 1
            self.df3['Pre-closing Modification Flag'] = 0

            del self.df

            self.df3 = self.df3.drop(labels='Left', axis=1)
            moodysSQLpull()
            FicoAdjust()
            self.df2.to_excel(fr'M:\Capital Markets\Users\Emmanuel Lopez\ADCO_{self.clicked.get()}.xlsx', index=False)
            self.milan_input.to_excel(fr'M:\Capital Markets\Users\Emmanuel Lopez\Milan_{self.clicked.get()}.xlsx',
                                      index=False)

            del self.df2, self.milan_input

            status = StringVar()
            status.set('Done!')
            status_entry = Entry(app, textvariable=status, width=8, bg='green', fg='white').grid(row=12, column=5,
                                                                                                 sticky=tk.W + tk.E)

        def moodysSQLpull():
            # Open SQL Connection
            conn = pyodbc.connect('Driver={SQL Server Native Client 11.0};'
                                  'Server=AWWDCORPSQLP02;'
                                  'Database=clg_reporting;'
                                  'Trusted_Connection=yes;')

            list_of_ids = list(self.df3['Loan Number'].apply(lambda x: f'{x}'))
            tup = tuple(list_of_ids)

            # Executing SQL queries and exporting results into dataframes
            bor_corr = pd.read_sql_query(f'''SELECT l.loannum as 'Loan Number', l.BorCount as 'Total Number of Borrowers' 
                                          FROM clg_reporting.dbo.vw_LOS_PCG_LoanDetails l
                                          WHERE l.LOS_Name = 'em' and l.TestLoanFlag = 'n'
                                          and loannum in {tup}
                                          ''', conn)
            bor_bdl = pd.read_sql_query(f'''SELECT LoanId as 'Loan Number', BorrowersCount as 'Total Number of Borrowers' 
                                            from rrd..HUB_BDL_LoanDetail
                                            where LoanId in {tup}
                                            ''', conn)
            bor_cdl = pd.read_sql_query(f'''SELECT LoanId as 'Loan Number', BorrowersCount as 'Total Number of Borrowers' 
                                            from rrd..HUB_CDL_LoanDetail
                                            where LoanId in {tup}
                                            ''', conn)
            piw = pd.read_sql_query(f'''SELECT [LoanNumber] as 'Loan Number',[property_inspection_waiver_indicator] as 'PIW' 
                                            FROM cm_inventory_management.ods_readonly.vw_pooling_data 
                                            WHERE [LoanNumber] in {tup}
                                            ''', conn)
            # Closes SQL connections
            conn.close()

            # Combines BDL CDL  & Correspondent BorCount and merges into Milan input dataframe
            bor_combined = pd.concat([bor_bdl, bor_cdl, bor_corr])
            del bor_bdl, bor_cdl, bor_corr
            bor_combined["Loan Number"] = bor_combined["Loan Number"].apply(lambda x: int(x))
            bor_combined.loc[(bor_combined['Total Number of Borrowers'].isnull()), 'Total Number of Borrowers'] = 1
            bor_combined["Total Number of Borrowers"] = bor_combined["Total Number of Borrowers"].apply(
                lambda x: int(x))
            merge_first = pd.merge(self.df3, bor_combined, left_on="Loan Number", right_on="Loan Number", how="left")
            del self.df3, bor_combined
            merge_first["Total Number of Borrowers_x"] = merge_first["Total Number of Borrowers_y"]
            merge_first = merge_first.drop(columns=["Total Number of Borrowers_y"])
            merge_first = merge_first.rename(columns={'Total Number of Borrowers_x': 'Total Number of Borrowers'})
            merge_first.loc[(merge_first['Total Number of Borrowers'].isna()), 'Total Number of Borrowers'] = 1
            merge_first['xxx'] = ''
            piw["Loan Number"] = piw["Loan Number"].apply(lambda x: int(x))
            self.milan_input = pd.merge(merge_first, piw, on="Loan Number", how="left")
            del merge_first, piw
            self.milan_input.loc[(self.milan_input['PIW'] == True), 'PIW'] = 'Y'
            self.milan_input.loc[(self.milan_input['PIW'] != 'Y'), 'PIW'] = 'N'

        def FicoAdjust():
            # DTI FICO Adjustments
            dti = self.milan_input['Originator DTI']
            self.df2.loc[(dti <= 0.30), 'FICO_Score'] = self.df2['FICO_Score'].apply(lambda x: x + 53.01)
            self.df2.loc[(dti > 0.30) & (dti <= 0.35), 'FICO_Score'] = self.df2['FICO_Score'].apply(lambda x: x + 22.33)
            self.df2.loc[(dti > 0.35) & (dti <= 0.40), 'FICO_Score'] = self.df2['FICO_Score'].apply(lambda x: x + 4.52)
            self.df2.loc[(dti > 0.40) & (dti <= 0.45), 'FICO_Score'] = self.df2['FICO_Score'].apply(lambda x: x - 7.42)
            self.df2.loc[(dti > 0.45), 'FICO_Score'] = self.df2['FICO_Score'].apply(lambda x: x - 51.85)

            # Purpose FICO Adjustments
            purp = self.df2['LoanPurpose']
            self.df2.loc[(purp == 'P'), 'FICO_Score'] = self.df2['FICO_Score'].apply(lambda x: x + 12.38)
            self.df2.loc[(purp == 'E'), 'FICO_Score'] = self.df2['FICO_Score'].apply(lambda x: x - 68.55)
            self.df2.loc[(purp == 'R'), 'FICO_Score'] = self.df2['FICO_Score'].apply(lambda x: x - 21.27)

            # Occupancy FICO Adjustments
            occ = self.df2['Occupancy']
            self.df2.loc[(occ == 'O'), 'FICO_Score'] = self.df2['FICO_Score'].apply(lambda x: x + 16.99)
            self.df2.loc[(occ == 'S'), 'FICO_Score'] = self.df2['FICO_Score'].apply(lambda x: x - 25.33)
            self.df2.loc[(occ == 'I'), 'FICO_Score'] = self.df2['FICO_Score'].apply(lambda x: x - 88.74)

            # Property Type FICO Adjustments
            prop = self.df2['PropertyType']
            self.df2.loc[(prop == 'SFR'), 'FICO_Score'] = self.df2['FICO_Score'].apply(lambda x: x + 3.25)
            self.df2.loc[(prop == 'Condo'), 'FICO_Score'] = self.df2['FICO_Score'].apply(lambda x: x - 12.49)
            self.df2.loc[(prop == 'MH'), 'FICO_Score'] = self.df2['FICO_Score'].apply(lambda x: x - 87.42)
            self.df2.loc[(prop == 'PUD'), 'FICO_Score'] = self.df2['FICO_Score'].apply(lambda x: x - 2.07)
            self.df2.loc[(prop == 'Coop'), 'FICO_Score'] = self.df2['FICO_Score'].apply(lambda x: x + 91.5)

            # Borrower Count FICO Adjustments
            borCount = self.milan_input['Total Number of Borrowers']
            self.df2.loc[(borCount == 1), 'FICO_Score'] = self.df2['FICO_Score'].apply(lambda x: x - 36.49)
            self.df2.loc[(borCount == 2), 'FICO_Score'] = self.df2['FICO_Score'].apply(lambda x: x + 38.4)
            self.df2.loc[(borCount > 2), 'FICO_Score'] = self.df2['FICO_Score'].apply(lambda x: x + 72.18)

            self.df2.loc[(self.df2['FICO_Score'] > 850), 'FICO_Score'] = 850
            self.df2.loc[(self.df2['FICO_Score'] < 350), 'FICO_Score'] = 350

            del dti, borCount, purp, occ, prop

        # Buttons
        openfile_btn = Button(app, bg='#31AFDF', fg='white', activebackground='#F1C400', activeforeground='white',
                              command=openfile, text="Open File").grid(row=10, column=4)
        runstrat_btn = Button(app, bg='#31AFDF', fg='white', activebackground='#F1C400', activeforeground='white',
                              command=runstrat, text="Create Inputs").grid(row=10, column=6)

        app.mainloop()


Mainapp()
