from collections import defaultdict
import io
from typing import List

import numpy as np
import pandas as pd
import pdfplumber
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows

from utils.embedded import csv_str
import os

def pencil():
    return "✏️"

def to_camel_case(text: str):
    text = text.replace("\n", " ")
    words = text.split()
    return words[0].lower() + "".join(word.capitalize() for word in words[1:])

import re

def extract_payment_id(payment_string):
    match = re.search(r'Payment #:\s*(\d+)', payment_string)

    if match:
        return True, match.group(1)
    else:
        return False, None

def extract_mm_dd(date_string):
    match = re.search(r'(\d{2}/\d{2})/\d{4}', date_string)

    if match:
        return True, match.group(1).replace('/', '-')
    else:
        return False, None

class CostcoTree(object):
    def __init__(self, dir_path: str, pdf_files: List[str], output_path: str) -> None:
        self.dir_path = dir_path
        self.list_of_pdfs = pdf_files
        self.output_path = output_path
        self.store_names = self.get_costco_store_names()

    def monthly_loop(self):
        wb = Workbook()
        wb.remove(wb.active)
        for idx, pdf_path in enumerate(self.list_of_pdfs):
            df1, df2, tab_name = self.get_table_from_pdf(pdf_path=pdf_path)
            pdf_name = os.path.basename(pdf_path)
            self.draw(df1, df2, tab_name=tab_name, wb=wb)

    def draw(self, df1, df2, tab_name, wb):
        sheetname = f'{tab_name[0]} #{tab_name[1]}'
        # output_path = os.path.join("/content/", f"{self.dir_path}-output.xlsx")
        ws = wb.create_sheet(title=sheetname)
        wb.active = ws

        for row in dataframe_to_rows(df1, index=False, header=True):
            if len(row[0]) <= 10:
              row[-1] = np.nan
            ws.append(row)

        for _ in range(2):
            ws.append([])

        for row in dataframe_to_rows(df2, index=False, header=True):
            ws.append(row)

        total = df2["amount"].sum()
        print(f"{sheetname} meta: {tab_name}")
        ws.append([])
        ws.append(["Total", total])
        ws.append(["Date", tab_name[0]])
        ws.append(["check number", tab_name[1]])

        wb.save(self.output_path)
        print("Finished drawing, " + sheetname)

    def get_costco_store_names(self):
        def key_formatter(s: str) -> str:
            if not s:
                raise ValueError("key cannot be empty.")

            if s.startswith("#") and len(s) <= 5:
                s = s.lstrip("#")
                return s.zfill(4)

            return "-1"

        df = pd.read_csv(io.StringIO(csv_str), usecols=[0, 2], header=None)
        df.columns = ["long", "short"]
        df["short"] = df["short"].apply(lambda x: key_formatter(x))

        store_names = defaultdict(str)
        for x, y in zip(df["long"], df["short"]):
            store_names[y] = x

        # add missed out fields
        missed = {"1997": "C991997", '0000': 'Unknown'}
        for x, y in missed.items():
            store_names[x] = y

        return store_names

    def __extract_key(self, s: str, n=-6):
          z = ''.zfill(4)
          if not s:
              return z


          res = s[:n]
          res = re.findall(r'\d+', res)

          if not res:
            print(f"No digits found in '{s}', returning default key '{z}'")
            return z
          res = res[0].lstrip('0').zfill(4)

          if res not in self.store_names:
              lres = res.lstrip('0')
              if lres in self.store_names:
                  return lres
              rres = res.rstrip('0').zfill(4)
              if rres in self.store_names:
                  return rres
              return z

          return res

    def get_table_from_pdf(self, pdf_path):
        with pdfplumber.open(pdf_path) as pdf:
            pages = pdf.pages
            data, tab_name = [], []
            for i, page in enumerate(pages):
              lines = page.extract_text_lines()
              for line in lines:
                if line['text'].startswith('Date'):
                  matched, res = extract_mm_dd(line['text'])
                  if matched:
                    tab_name.append(res)
                else:
                  matched, res = extract_payment_id(line['text'])
                  if matched:
                    tab_name.append(res)
                    break

              table = page.extract_table()
              if table:
                  if all(not tr for tr in table[-1]):
                      table = table[:-1]
                  if i:
                      table.pop(0)
                  data.extend(table)

        df = pd.DataFrame(data[1:], columns=[data[0]])
        df = df.rename(columns=lambda x: to_camel_case(x))

        # convert multi-index to single-index to enable groupby
        df.columns = df.columns.get_level_values(0)

        df["storeKey"] = df["invoiceNumber"].apply(self.__extract_key)
        df["storeName"] = df["storeKey"].map(
            lambda key: self.store_names.get(key, "-1")
        )
        for idx in np.where(df["storeName"] == "-1")[0]:
            inv = df.loc[idx, "invoiceNumber"]
            n = -7
            if len(inv) < 11:
                n = -6
            skey = self.__extract_key(inv, n=n)
            sval = self.store_names[skey]
            df.loc[idx, "storeKey"] = skey
            df.loc[idx, "storeName"] = sval

        missed = np.where(df["storeName"] == "-1")[0]
        if len(missed):
            raise AssertionError("invalid key.")

        df["amount"] = df["amount"].replace(",", "", regex=True).astype(float)

        df2 = df[["storeName", "amount"]].copy()
        df2 = df2.groupby("storeName", as_index=False).sum()
        return df, df2, tab_name


