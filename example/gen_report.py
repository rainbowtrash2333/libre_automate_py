from workbook import Workbook
from word import Word
from myutil import array2df, process_value_to_str
import pandas as pd
from typing import Sequence
import os
import shutil
from officeLoader import OfficeLoader

if __name__ == '__main__':
    data_path = r"F:\客户风险\202503"
    wb = Workbook(read_only=True, filepath=rf"{data_path}\data.xls", visible=False)
    data = wb.get_used_value(0, range_name='A1:C121')
    df = array2df(data)
    processed_values = df.apply(process_value_to_str, axis=1)

    wb.close()
    date_str = "2025年3月"
    template_file = r"F:\客户风险\teml\teml.doc"
    result_file = rf"{data_path}\{date_str}昭通市银行业客户风险监测报告.doc"
    if not os.path.exists(result_file):
        shutil.copy2(template_file, result_file)
    word = Word(read_only=False, filepath=result_file, visible=True)
    df['label'] = df['label'].apply(lambda x: f"$({x})" if pd.notna(x) and x != '' else x)
    labels_list: Sequence[str] = df['label'].tolist()
    values_list: Sequence[str] = processed_values.tolist()

    word.replace_words(labels_list, values_list)
   # Word.save()
    office_loader = OfficeLoader()
    office_loader.close()
