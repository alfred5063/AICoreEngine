
import openpyxl
import pandas as pd
from openpyxl import Workbook

#download vendor transaction, 
def vendor_transaction(vt_excel):
	vt_df = pd.read_excel(vt_excel, header = None)
  #copy column d + e to the last column
	vt_df[11]=vt_df[3] 
	vt_df[12]=vt_df[4]
	df = pd.DataFrame(vt_df.values,
    columns=['Invoice No', 'Doc Types','Client','GL No', 'Register Date','Due Date',
             'Batch-Entry','Paid Invoice No/Payment Ref','Invoice Amount','Paid Amount',
             'Balance','Payment Batch','Payment Date']
    )
	df.to_excel(vt_excel, index=False)
	return df


