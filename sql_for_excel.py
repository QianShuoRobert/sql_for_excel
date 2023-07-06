#!/usr/bin/env python
# - *- coding: utf-8 -*-

# @File: sql_for_excel.py
# @Time: 2023/06/26 09:35:00
# @Author: Robert
# @Contact: 891581750@qq.com
# @Licence: MIT
# @Desc: None

import sys
import os
import pandas
from pathlib import Path
import sqlite3
# import markdown
import logging
import time
from PyQt5.QtWidgets import QMainWindow, QApplication, QFileDialog, QTreeWidgetItem, QTableWidgetItem, QMessageBox, QMenu, QAction
from PyQt5.QtCore import Qt
from PyQt5.QtGui import QCursor, QIcon
from collections import namedtuple
from enum import Enum
from main_window import Ui_MainWindow

# 设置日志参数
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

# 树节点类型定义
class TreeNodeType(Enum):
    File = 0 # 文件
    Sheet = 1 # 表格
    Field = 2 # 字段

# 挂在树节点中的自定义内容
# node_type: TreeNodeType
TreeNodeData = namedtuple('TreeNodeData', ['node_type', 'value'])

class MyApp(QMainWindow, Ui_MainWindow):
    def __init__(self) -> None:
        QApplication.__init__(self)
        self.setupUi(self)
        self._base_path = Path(__file__).parent
        self._icons_path = self._base_path / 'icons'
        self.setWindowIcon(QIcon(str(self._icons_path / 'main.svg')))
        self._bind_ui_events()
        # 创建数据库
        self._conn = sqlite3.connect(':memory:')
        # self._conn = sqlite3.connect(self._base_path / 'test.db') # 调试先用本地数据库好查看 debug
        self._cursor = self._conn.cursor()
        self._query_columns = None # 保存当前的查询结果的列，用于导出结果
        self._query_result = None # 保存当前的查询结果，用于导出结果
    
    def __del__(self) -> None:
        self._conn.close()
    
    def _bind_ui_events(self) -> None:
        # 文件和表树
        self.treeWidgetExcelsAndSheets.setColumnCount(2)
        self.treeWidgetExcelsAndSheets.setHeaderLabels(['名称', '类型'])
        self.treeWidgetExcelsAndSheets.setColumnWidth(0, 300)
        self.treeWidgetExcelsAndSheets.setContextMenuPolicy(Qt.CustomContextMenu) # 打开右键菜单策略
        # 绑定菜单事件
        self.treeWidgetExcelsAndSheets.customContextMenuRequested.connect(self._treeWidgetItem_popContextMenu)
        # 菜单选项
        self._contextMenu = {
            TreeNodeType.File: [
                ('在文件夹中查看文件', self._treeWidgetItem_popContextMenu_ShowInDir),
                ('从列表中移除此文件', self._treeWidgetItem_popContextMenu_RemoveFileFromTree),
            ],
            TreeNodeType.Sheet: [
                ('[表名] 插入到SQL', self._treeWidgetItem_popContextMenu_InsertSheetName),
                ('查看表格数据', self._treeWidgetItem_popContextMenu_ShowSheetData),
                ('从列表中移除此表', self._treeWidgetItem_popContextMenu_RemoveSheetFromTree),
            ],
            TreeNodeType.Field: [
                ('"字段名称" 插入到SQL', self._treeWidgetItem_popContextMenu_InsertFieldName),
                ('"字段名称", 插入到SQL', self._treeWidgetItem_popContextMenu_InsertFieldNameWithComma),
            ],
        }
        # Excel文件和表
        self.pushButtonImportFile.clicked.connect(self.pushButtonImportFile_clicked)
        # SQL
        self.pushButtonFormatSql.clicked.connect(self.pushButtonFormatSql_clicked)
        self.pushButtonRunSql.clicked.connect(self.pushButtonRunSql_clicked)
        self.textEditSql.textChanged.connect(self.textEditSql_textChanged)
        # 执行结果
        self.pushButtonExportResult.clicked.connect(self.pushButtonExportResult_clicked)

    def _treeWidgetItem_popContextMenu_ShowInDir(self, currentItem) -> None:
            '''在文件夹中查看文件'''
            pname = currentItem.data(0, Qt.UserRole).value
            self._show_file_in_folder(pname)

    def _remove_excel_node(self, currentItem) -> None:
        for index in range(currentItem.childCount()):
            child = currentItem.child(index)
            self._remove_sheet_node(child)

    def _treeWidgetItem_popContextMenu_RemoveFileFromTree(self, currentItem) -> None:
        '''从列表中移除此文件'''
        file_name = currentItem.data(0, Qt.UserRole).value
        self._remove_excel_node(currentItem)
        self.statusbar.showMessage(f'从列表中移除文件成功！【{file_name}】')
    
    def _treeWidgetItem_popContextMenu_InsertSheetName(self, currentItem) -> None:
        '''表名称插入到SQL'''
        sheet_name = currentItem.data(0, Qt.UserRole).value
        self.textEditSql.insertPlainText(f' {sheet_name} ')
    
    def _treeWidgetItem_popContextMenu_ShowSheetData(self, currentItem) -> None:
        '''查看表格数据'''
        sheet_name = currentItem.data(0, Qt.UserRole).value
        sql = f'select * from {sheet_name}'
        self._cursor.execute(sql)
        self._query_columns = [desc[0] for desc in self._cursor.description]
        self._query_result = self._cursor.fetchall()
        self._show_query_result()
        self.statusbar.showMessage(f'查看表格数据：{sheet_name}')
    
    def _remove_sheet_node(self, currentItem) -> None:
        sheet_name = currentItem.data(0, Qt.UserRole).value
        sql = f'drop table {sheet_name}'
        self._cursor.execute(sql)
        # 从树上删除
        parentItem = currentItem.parent()
        parentItem.removeChild(currentItem)
        if parentItem.childCount() == 0:
            # 不起作用
            # self.treeWidgetExcelsAndSheets.removeItemWidget(parentItem, 0) 
            # 网上找的也不起作用，没有takeItem 接口
            # self.treeWidgetExcelsAndSheets.takeItem(self.treeWidgetExcelsAndSheets.row(parentItem))
            # 先隐藏起来吧，没什么坏影响
            parentItem.setHidden(True)

    def _treeWidgetItem_popContextMenu_RemoveSheetFromTree(self, currentItem) -> None:
        '''从列表中移除此表'''
        sheet_name = currentItem.data(0, Qt.UserRole).value
        self._remove_sheet_node(currentItem)
        self.statusbar.showMessage(f'从列表中移除此表成功：{sheet_name}')
    
    def _treeWidgetItem_popContextMenu_InsertFieldName(self, currentItem) -> None:
        '''字段名称插入到SQL'''
        field_name = currentItem.data(0, Qt.UserRole).value
        self.textEditSql.insertPlainText(f' {field_name} ')
    
    def _treeWidgetItem_popContextMenu_InsertFieldNameWithComma(self, currentItem) -> None:
        '''字段名称插入到SQL'''
        field_name = currentItem.data(0, Qt.UserRole).value
        self.textEditSql.insertPlainText(f' {field_name}, ')

    def _treeWidgetItem_popContextMenu(self, pos) -> None:
        currentItem = self.treeWidgetExcelsAndSheets.currentItem()
        pos_item = self.treeWidgetExcelsAndSheets.itemAt(pos)
        if currentItem and pos_item:
            pop_menu = QMenu(self.treeWidgetExcelsAndSheets)
            node_type: TreeNodeType = currentItem.data(0, Qt.UserRole).node_type
            if node_type in self._contextMenu:
                for item in self._contextMenu[node_type]:
                    action = QAction(item[0], self) # 名称
                    # 也可以用action.trigger.connect() 直接绑定处理函数(不带入参)，
                    # 这里用setData和data，后面pop_menu.triggered[QAction].connect绑定所有action相同的入口
                    action.setData(item[1]) # 处理函数
                    pop_menu.addAction(action)
            else:
                err_info = f'_treeWidgetItem_popContextMenu unknown node_type{node_type}'
                logging.error(err_info)
                raise Exception(err_info)
            pop_menu.triggered[QAction].connect(self._processContextMenu)
            pop_menu.exec(QCursor.pos())
    
    def _processContextMenu(self, qaction: QAction) -> None:
        currentItem = self.treeWidgetExcelsAndSheets.currentItem()
        if currentItem:
            try:
                qaction.data()(currentItem)
            except Exception as ex:
                # 加一层异常处理，防止因为一下UI异常崩溃
                error_info = f'【{qaction.text()}】发生异常！{str(ex)}'
                logging.error(error_info)
                self.statusbar.showMessage(error_info)
    
    def _add_excel_node(self, pname: Path) -> QTreeWidgetItem:
        new_file_node = QTreeWidgetItem(self.treeWidgetExcelsAndSheets)
        new_file_node.setIcon(0, QIcon(str(self._icons_path / 'excel.svg')))
        new_file_node.setData(0, Qt.UserRole, TreeNodeData(TreeNodeType.File, str(pname)))
        new_file_node.setText(0, pname.stem)
        new_file_node.setText(1, pname.suffix)
        new_file_node.setExpanded(True)
        return new_file_node
    
    def _get_db_tables_name(self) -> list[str]:
        tables_name = []
        self._cursor.execute("SELECT name FROM sqlite_master WHERE type='table'")
        query_result = self._cursor.fetchall()
        if query_result:
            tables_name = [row[0] for row in query_result]
        return tables_name

    def _add_sheet_node(self, sheet_name, parent, df: pandas.DataFrame) -> None:
        # 创建一个表，先判断表名是否冲突，如果冲突分配一个新名字
        table_name = sheet_name
        table_name_index = 1
        db_tables_name = self._get_db_tables_name()
        while table_name in db_tables_name:
            table_name = f'{table_name}_{table_name_index}'
        # 插入到树上
        new_sheet_node = QTreeWidgetItem(parent)
        new_sheet_node.setIcon(0, QIcon(str(self._icons_path / 'table.svg')))
        new_sheet_node.setData(0, Qt.UserRole, TreeNodeData(TreeNodeType.Sheet, f'[{table_name}]'))
        new_sheet_node.setText(0, table_name)
        new_sheet_node.setExpanded(False)
        # df = dfs[sheet]
        # logging.info(df.dtypes)
        # logging.info(list(df.dtypes))
        db_colunms = []
        for col in df.columns:
            new_field_node = QTreeWidgetItem(new_sheet_node)
            new_field_node.setIcon(0, QIcon(str(self._icons_path / 'field.svg')))
            new_field_node.setData(0, Qt.UserRole, TreeNodeData(TreeNodeType.Field, f'"{col}"'))
            # new_field_node.setText(0, f'{col}: {df.dtypes[col]}')
            col_dtype = 'TEXT'
            new_field_node.setText(0, col)
            new_field_node.setText(1, col_dtype)
            db_colunms.append([f'"{col}"', col_dtype])
        # 拼装sql 创建表
        fields = ', '.join([f'{db_col[0]} {db_col[1]}' for db_col in db_colunms])
        sql = f'''CREATE TABLE [{table_name}] ({fields})'''
        self._cursor.execute(sql)
        # 把数据加进去
        sql = f'''insert into [{table_name}] values({','.join(['?'] * len(db_colunms))})'''
        self._cursor.executemany(sql, df.values)

    def pushButtonImportFile_clicked(self):
        fname, *_ = QFileDialog.getOpenFileName(self, '导入Excel', '', 'Excel Files (*.xlsx *.xls)')
        if fname:
            try:
                t1 = time.time()
                pname = Path(fname)
                dfs = pandas.read_excel(pname, sheet_name=None, dtype=str)
                sheet_names = list(dfs.keys())
                # 加到树上
                new_field_node = self._add_excel_node(pname)
                for sheet in sheet_names:
                    self._add_sheet_node(sheet, new_field_node, dfs[sheet])
                self._conn.commit()
                t2 = time.time()
                self.statusbar.showMessage(f'导入Excel文件【{pname}】成功！用时[{(t2 - t1):.2f}s]')
            except Exception as ex:
                error_info = f'导入Excel文件【{pname}】失败！ {str(ex)}'
                logging.error(error_info)
                self.statusbar.showMessage(error_info)
        else:
            self.statusbar.showMessage(f'未选择有效的Excel文件！')


    def textEditSql_textChanged(self):
        # logging.info('textEditSql_textChanged')
        # 先解绑处理函数，重新赋值后再绑定
        # self.textEditSql.textChanged.disconnect(self.textEditSql_textChanged)
        # text = self.textEditSql.toPlainText()
        # md_text = f'``` sql\n{text}\n```'
        # html = markdown.markdown(md_text)
        # self.textEditSql.setHtml(html)
        # # self.textEditSql.setPlainText(text)
        # self.textEditSql.textChanged.connect(self.textEditSql_textChanged)
        pass

    def pushButtonFormatSql_clicked(self):
        self.statusbar.showMessage('todo') # todo

    def _show_query_result(self) -> None:
        self.tableWidgetSqlResult.setColumnCount(len(self._query_columns))
        self.tableWidgetSqlResult.setRowCount(0)
        self.tableWidgetSqlResult.setHorizontalHeaderLabels(self._query_columns)
        for index, row in enumerate(self._query_result):
            self.tableWidgetSqlResult.setRowCount(index + 1)
            cells = [QTableWidgetItem(str(value) if value else '') for value in row]
            [self.tableWidgetSqlResult.setItem(index, i, v) for i,v in enumerate(cells)]

    def pushButtonRunSql_clicked(self):
        self._query_columns = None
        self._query_result = None
        self.tableWidgetSqlResult.clear()
        self.tableWidgetSqlResult.setColumnCount(0)
        self.tableWidgetSqlResult.setRowCount(0)
        sql = self.textEditSql.toPlainText().strip()
        if not sql:
            QMessageBox.information(self, '执行SQL', 'SQL内容为空！', QMessageBox.Yes, QMessageBox.Yes)
            return
        t1, t2, t3, t4 = 0.0, 0.0, 0.0, 0.0
        try:
            t1 = time.time()
            self._cursor.execute(sql)
            t2 = time.time()
        except sqlite3.OperationalError as ex:
            error_info = f'执行SQL失败！ {str(ex)}'
            self.statusbar.showMessage(error_info)
            return
        if not self._cursor.description:
            self.statusbar.showMessage(f'执行SQL结果为空！')
            return
        self._query_result = self._cursor.fetchall()
        t3 = time.time()
        self._query_columns = [desc[0] for desc in self._cursor.description]
        self._show_query_result()
        t4 = time.time()
        self.statusbar.showMessage(f'执行SQL成功！执行用时[{(t2 - t1):.2f}s]，获取数据用时[{(t3 - t2):.2f}s]，显示数据用时[{(t4 - t3):.2f}s]')

    def pushButtonExportResult_clicked(self):
        if not self._query_result:
            QMessageBox.information(self, '导出执行结果', '当前没有执行结果！', QMessageBox.Yes, QMessageBox.Yes)
            return
        currentItem = self.treeWidgetExcelsAndSheets.currentItem()
        if not currentItem:
            QMessageBox.information(self, '导出执行结果', '请先在Excel文件和表中选择一个目标Excel文件！', QMessageBox.Yes, QMessageBox.Yes)
            return
        while currentItem.parent():
            currentItem = currentItem.parent()
        pname = currentItem.data(0, Qt.UserRole).value
        new_sheet_name: str = self.lineEditNewSheetName.text()
        if new_sheet_name in self._get_db_tables_name():
            info = '已存在相同的表格名称，请填写不同的表格名称！'
            QMessageBox.information(self, 'Excel表格名称无效', info, QMessageBox.Yes, QMessageBox.Yes)
            return
        if new_sheet_name == '' or len(new_sheet_name.encode('gbk')) > 31 or next(filter(lambda x: x in ['\/?*[]'], new_sheet_name), None):
            info = '请确保：\n\n名称不多于31个字符。\n名称不包含下列任一字符： : \\ / ? * [ 或 ]\n名称不为空。'
            QMessageBox.information(self, 'Excel表格名称无效', info, QMessageBox.Yes, QMessageBox.Yes)
            return
        try:
            excel_writer = pandas.ExcelWriter(pname, mode='a')
            new_df = pandas.DataFrame(data=self._query_result, columns=self._query_columns)
            new_df.to_excel(excel_writer, new_sheet_name, index=False)
            excel_writer.close()
            # 导出成功后，在文件树上增加这个表
            self._add_sheet_node(new_sheet_name, currentItem, new_df)
            # 弹窗提示
            self.statusbar.showMessage(f'导出成功！已导出为文件【{pname}】的表[{new_sheet_name}]')
            yes_or_no = QMessageBox.question(self, '导出执行结果', '导出成功！是否在文件夹中查看文件？', QMessageBox.Yes | QMessageBox.No, QMessageBox.Yes)
            if yes_or_no == QMessageBox.Yes:
                self._show_file_in_folder(pname)
        except Exception as ex:
            error_info = f'导出执行结果失败！异常信息：{str(ex)}'
            logging.error(error_info)
            self.statusbar.showMessage(error_info)

    def _show_file_in_folder(self, fpath) -> None:
        cmd = f'explorer /select,"{Path(fpath)}"' # qt的path是/格式的，需要转换成windows的\格式
        os.popen(cmd) # 打开文件目录，这里不用os.system因为会有一个黑框一闪而过

if __name__ == '__main__':
    qt_app = QApplication(sys.argv)
    my_app = MyApp()
    my_app.show()
    sys.exit(qt_app.exec())
