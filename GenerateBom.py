import sys
import traceback

import pandas as pd
from PySide2.QtCore import *
from PySide2.QtWidgets import *
from PySide2.QtUiTools import *
import datetime
import os
import xlrd

class MyWidget(QWidget):
    def __init__(self):
        super().__init__()

        # Load UI file
        loader = QUiLoader()
        self.ui = loader.load('bom.ui', self)

        # Connect signals
        self.ui.uploadFile.clicked.connect(self.select_file)
        self.ui.analysis.clicked.connect(self.parse_file)

        # Show the UI
        self.ui.show()

    def pase(self):
        now = datetime.datetime.now()
        date_str = now.strftime('%Y-%m-%d_%H-%M-%S')
        # date_str = "123"
        dir_path, file_name = os.path.split(self.file_path)
        result_path = os.path.join(dir_path, 'result_' + date_str + '.xlsx')
        # result_path = '123.xlsx'
        sheet_names = ['00-菜单Bom', '02-MF流程']
        dfs = pd.read_excel(self.file_path, sheet_name=sheet_names)

        # 读取第一个sheet的内容，并处理合并单元格
        # df_recipe_bom = dfs[excel_file.sheet_names[0]]
        df_recipe_bom = dfs['00-菜单Bom']
        df_recipe_bom.fillna(method='ffill', inplace=True)
        df_process = dfs['02-MF流程']
        df_process.fillna(method='ffill', inplace=True)

        # 获取工艺flow及每个step上使用的recipe
        df_step_recipe = df_process.iloc[:, [1, 2, 4]]

        df_recipe_material = df_recipe_bom.iloc[:, [1, 2, 3, 4, 5, 12, 14]]

        # 按 Recipe编号 列进行合并，使用左连接
        merged_df = pd.merge(df_step_recipe, df_recipe_material, on='Recipe编号', how='left')

        # 获取元件料号为NaN的行
        not_nan_df = merged_df.loc[merged_df['元件料号'].isna()]
        # 打印工艺flow中存在但是bom统计中不存在的内容
        self.ui.textEdit.append(">>>>>>>>>>如下信息存在于工艺flow中但不存在于菜单BOM中<<<<<<<<<<")
        self.ui.textEdit.append(not_nan_df.iloc[:, [0, 1, 2]].to_string(index=False))

        # 添加新列 主件底数，值为1
        merged_df['主件底数'] = 1
        grouped = merged_df.groupby(['品名', '规格', '单位', 'layer', '元件料号', '主件底数'])[['合计', '宽放后']].sum()
        # append_to_excel(file_path, grouped, 'new_sheet1')
        # 将grouped写入Excel文件的新sheet中
        # with pd.ExcelWriter(file_path, engine='openpyxl', mode='a') as writer:
        #     grouped.to_excel(writer, sheet_name='new_sheet10', index=True)

        # 创建ExcelWriter
        writer = pd.ExcelWriter(result_path, engine='openpyxl')

        # 将DataFrame写入Excel文件
        grouped.to_excel(writer, sheet_name='Sheet1', index=True)

        # 关闭ExcelWriter
        writer.close()
        return result_path

    def select_file(self):
        file_path, _ = QFileDialog.getOpenFileName(self, 'Select file', QDir.homePath(), 'Excel Files (*.xlsx *.xls)')
        self.ui.textEdit.setText(file_path)
        self.file_path = file_path

    def parse_file(self):
        self.ui.textEdit.append("解析中.....")
        file_path = self.file_path
        if not file_path:
            QMessageBox.warning(self, 'Warning', 'Please select a file.')
            return
        try:
            result_path = self.pase()
            self.ui.textEdit.append(">>>>>>>>>>文件"+result_path+"生成完成<<<<<<<<<<")
        except Exception as e:
            # self.ui.textEdit.append(str(traceback.format_exc()))
            QMessageBox.warning(self, 'Warning', str(traceback.format_exc()))


if __name__ == '__main__':
    app = QApplication()
    widget = MyWidget()
    widget.show()
    sys.exit(app.exec_())
