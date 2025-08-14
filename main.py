from fetchers.secur_doc import paste_to_excel_secur_doc
from fetchers.grp_9000 import paste_to_excel_9000grp
from fetchers.dz_spot import paste_to_excel_dz_spot
from fetchers.balance_nrk import paste_to_excel_balance_nrk
from fetchers.dz_spot_diff import paste_to_excel_diff_spot
from mail.forecast_nrk import generate_forecast_email
from fetchers.diff_acc import paste_to_excel_diff_acc
from fetchers.fz_ccf_6jx import paste_to_excel_fz_ccf_6jx
from fetchers.doc_acc import paste_to_excel_doc_acc

# Основные вызовы (вызываются из Excel через xlwings)
def run_secur_doc():
    paste_to_excel_secur_doc()

def run_grp_9000():
    paste_to_excel_9000grp()

def run_dz_spot():    
    paste_to_excel_dz_spot()
    
def run_balance_nrk():    
    paste_to_excel_balance_nrk()

def run_diff_spot():    
    paste_to_excel_diff_spot()

def run_forecast_mail():
    generate_forecast_email()
    
def run_diff_acc():
    paste_to_excel_diff_acc()

def run_fz_ccf_6jx():
    paste_to_excel_fz_ccf_6jx()

def run_doc_acc():
    paste_to_excel_doc_acc()
