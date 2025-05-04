import pdfplumber
import pandas as pd
import tkinter as tk
from tkinter import filedialog, simpledialog


def get_pdf_path_gui():
  """Opens a file selection dialog and returns the selected PDF path."""
  root = tk.Tk()
  root.withdraw()  # Hide the main window
  file_path = filedialog.askopenfilename(
      title="Select Password-Protected PDF File",
      filetypes=[("PDF files", "*.pdf"), ("All files", "*.*")],
  )
  return file_path


def get_password_gui():
  """Opens a simple dialog to prompt for the PDF password."""
  root = tk.Tk()
  root.withdraw()  # Hide the main window
  password = simpledialog.askstring(
      "Password Required", "Enter the password for the PDF:", show="*"
  )
  return password


def read_password_protected_pdf_table_pdfplumber(pdf_path, password, table_settings=None):
  """
  Reads a table from a specific page of a password-protected PDF file using pdfplumber.

  Args:
    pdf_path (str): The path to the password-protected PDF file.
    password (str): The password to decrypt the PDF.
    table_settings (dict, optional): Dictionary of table extraction settings for pdfplumber.
                                      See pdfplumber documentation for available options.
                                      Defaults to None, which uses pdfplumber's auto-detection.
    page_number (int, optional): The page number where the table is located (1-based). Defaults to 1.

  Returns:
    pandas.DataFrame or None: A pandas DataFrame containing the extracted table,
                              or None if the table is not found or an error occurs.
  """
  try:
    with pdfplumber.open(pdf_path, password=password) as pdf:
      def remove_none_in_list(l):
        return [item for item in l if item is not None]
      def is_good_row(row):
        # Expect 9 columns for income table
        # Filter only income rows
        return len(row) >= 9 and row[1] and row[1].endswith('所得')

      income_table = []
      for page in pdf.pages:
        if table := page.extract_table(table_settings=table_settings):
          # Chances are that we get some None from a regular row
          income_table.extend([remove_none_in_list(row) for row in table if is_good_row(row)])
      return income_table
  except FileNotFoundError:
    print(f'Error: File not found at "{pdf_path}".')
    return None
  except Exception as e:
    print(f'An error occurred: {e}')
    return None


if __name__ == '__main__':
    # Get PDF file path using a file selection dialog
    pdf_file = get_pdf_path_gui()
    if not pdf_file:
      print("No PDF file selected. Exiting.")
      exit()

    # Get password using a simple dialog
    pdf_password = get_password_gui()
    if not pdf_password:
      print("Password not entered. Exiting.")
      exit()

    # Optional table settings for more precise extraction
    # Explore pdfplumber documentation for options like 'vertical_strategy',
    # 'horizontal_strategy', 'explicit_vertical_lines', 'explicit_horizontal_lines', etc.
    custom_table_settings = {
      'horizontal_strategy': 'lines',
      'vertical_strategy': 'lines',
      'snap_tolerance': 3,
    }

    if table := read_password_protected_pdf_table_pdfplumber(
      pdf_file, pdf_password, table_settings=custom_table_settings
    ):
      table_df = pd.DataFrame(table, columns=['所得人', '所得類別', '格式代號', '所得發生處所名稱', '統一編號', '給付/收入總額', '必要費用及成本', '所得總額', '扣繳稅額'])
      breakdown_df = table_df[['所得人', '所得類別', '給付/收入總額', '扣繳稅額']]
      print('Filtered Table (using pdfplumber):')
      print(breakdown_df)
      breakdown_df['給付/收入總額'] = breakdown_df['給付/收入總額'].str.replace(',', '').astype(int)
      breakdown_df['扣繳稅額'] = breakdown_df['扣繳稅額'].str.replace(',', '').astype(int)
      print()
      print('Income Breakdown')
      print(breakdown_df.groupby(['所得人', '所得類別']).sum().to_csv(sep=';'))
      print('Income Summary')
      summary_df = breakdown_df.drop('所得類別', axis='columns')
      print(summary_df.groupby('所得人').sum().to_csv(sep=';'))
