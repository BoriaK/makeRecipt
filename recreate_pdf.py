"""Re-export the existing .docx receipt as PDF (no field changes)."""
import os, win32com.client

DOCX = r'C:\Users\Administrator\Documents\קבלות_על_אימונים\example_receipt.docx'
PDF  = r'C:\Users\Administrator\Documents\קבלות_על_אימונים\recipt_16_April_2026.pdf'

word = win32com.client.Dispatch("Word.Application")
word.Visible = False
try:
    doc = word.Documents.Open(os.path.abspath(DOCX))
    doc.ExportAsFixedFormat(os.path.abspath(PDF), 17)
    doc.Close(False)
    print(f"PDF created: {PDF}")
finally:
    word.Quit()

