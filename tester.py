from data_processor import convert_pd
from data_processor import create_report
from datetime import datetime
import pandas as pd

BL=r"C:\Users\raymo\OneDrive\Documents\Paychex proj\test data\BL_time_rn.xlsx"
FG=r"C:\Users\raymo\OneDrive\Documents\Paychex proj\test data\Invoice_Status_rn.xlsx"
CP=r"C:\Users\raymo\OneDrive\Documents\Paychex proj\test data\\CW_rn.xlsx"


PD=r"C:\Users\raymo\OneDrive\Documents\Paychex proj\test data\Full-2021-plus.xlsx"
rep=r"C:\Users\raymo\OneDrive\Documents\Paychex proj\test data"


# rep = convert_pd(CP, datetime(2022, 1, 1), datetime(2022, 3, 10),"CP")
# print(rep)

out = create_report({"BL":BL, "FG": FG, "CP": CP}, PD, rep, datetime(2020, 1, 1), datetime(2022, 6, 1))
print(out)

# print(rep)
# print(rep.index[0]==pd.Timestamp('1/10/2022'))