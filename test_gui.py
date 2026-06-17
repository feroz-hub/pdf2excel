import tkinter as tk
from pdf_to_excel_gui import App

app = App()
app.pdf_var.set("NIST.SP.800-53r5.pdf")
app.fmt_var.set("Standard Assessment")
app.mode_var.set("Auto")
app.profile_var.set("Auto")
app.gen_review_var.set(True)
app.show_issues_var.set(True)
app.force_export_var.set(False)
app.out_var.set("NIST_fixed_gui.xlsx")

# Call extract
app._on_extract()

import time
while app.extract_btn['state'] == 'disabled':
    app.update()
    time.sleep(0.1)
    
app._poll_events()
while app.extract_btn['state'] == 'disabled':
    app.update()
    app._poll_events()
    time.sleep(0.1)

# Call export
app._on_export()

print("GUI export completed.")
