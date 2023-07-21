#!/usr/bin/env python
# - *- coding: utf-8 -*-

# @File: SqlHighlighter.py
# @Time: 2023/07/20 19:05:00
# @Author: robertqian
# @Contact: robertqian@live.com
# @Licence: MIT
# @Desc: None

import re
from PyQt5.QtCore import Qt
from PyQt5.QtGui import (QColor, QFont, QFontDatabase,
                         QSyntaxHighlighter, QTextCharFormat)


class SqlHighlighter(QSyntaxHighlighter):
    def __init__(self, editor, parent=None):
        QSyntaxHighlighter.__init__(self, parent)
        self._keywords = ['ABORT', 'ACTION', 'ADD', 'AFTER', 'ALL',
            'ALTER', 'ANALYZE', 'AND', 'AS', 'ASC', 'ATTACH', 'AUTOINCREMENT',
            'BEFORE', 'BEGIN', 'BETWEEN', 'BY', 'CASCADE', 'CASE', 'CAST', 'CHECK',
            'COLLATE', 'COLUMN', 'COMMIT', 'CONFLICT', 'CONSTRAINT', 'CREATE',
            'CROSS', 'CURRENT_DATE', 'CURRENT_TIME', 'CURRENT_TIMESTAMP',
            'DATABASE', 'DEFAULT', 'DEFERRABLE', 'DEFERRED', 'DELETE', 'DESC',
            'DETACH', 'DISTINCT', 'DROP', 'EACH', 'ELSE', 'END', 'ESCAPE', 'EXCEPT',
            'EXCLUSIVE', 'EXISTS', 'EXPLAIN', 'FAIL', 'FOR', 'FOREIGN', 'FROM',
            'FULL', 'GLOB', 'GROUP', 'HAVING', 'IF', 'IGNORE', 'IMMEDIATE', 'IN', 'INDEX',
            'INDEXED', 'INITIALLY', 'INNER', 'INSERT', 'INSTEAD',
            'INTERSECT', 'INTO', 'IS', 'ISNULL', 'JOIN', 'KEY', 'LEFT', 'LIKE',
            'LIMIT', 'MATCH', 'NATURAL', 'NO', 'NOT', 'NOTNULL', 'NULL', 'OF',
            'OFFSET', 'ON', 'OR', 'ORDER', 'OUTER', 'PLAN', 'PRAGMA', 'PRIMARY', 
            'QUERY', 'RAISE', 'RECURSIVE', 'REFERENCES', 'REGEXP', 'REINDEX', 'RELEASE',
            'RENAME', 'REPLACE', 'RESTRICT', 'RIGHT', 'ROLLBACK', 'ROW', 'SAVEPOINT', 
            'SELECT', 'SET', 'TABLE', 'TEMP', 'TEMPORARY', 'THEN', 'TO', 'TRANSACTION', 
            'TRIGGER', 'UNION', 'UNIQUE', 'UPDATE', 'USING', 'VACUUM', 'VALUES', 
            'VIEW', 'VIRTUAL', 'WHEN', 'WHERE', 'WITH', 'WITHOUT']
        self._mappings = {}
        self._tables_name = []
        self._setup_editor(editor)

    def update_tables_name(self, tables_name: list[str]) -> None:
        '''
        更新所有的tables表名，用于高亮显示
        '''
        self._tables_name = tables_name

    def _add_mapping(self, pattern, format):
        self._mappings[pattern] = format

    def highlightBlock(self, text):
        # 关键词和固定格式高亮
        for pattern, format in self._mappings.items():
            for match in re.finditer(pattern, text):
                start, end = match.span()
                self.setFormat(start, end - start, format)
        # 动态添加的表格名称高亮
        tables_pattern = r'\b(?:{})\b'.format('|'.join(self._tables_name))
        tables_format = QTextCharFormat()
        tables_format.setForeground(Qt.blue)
        for match in re.finditer(tables_pattern, text):
            start, end = match.span()
            self.setFormat(start, end - start, tables_format)

    def _setup_editor(self, editor):
        # keywords
        class_format = QTextCharFormat()
        class_format.setFontWeight(QFont.Bold)
        class_format.setForeground(Qt.darkRed)
        pattern = r'\b(?i:{})\s'.format('|'.join(self._keywords))
        self._add_mapping(pattern, class_format)
        # [xxxx]
        function_format = QTextCharFormat()
        # function_format.setFontItalic(True)
        function_format.setForeground(Qt.blue)
        pattern = r'\[([^\]]*)\]'
        self._add_mapping(pattern, function_format)
        # 'xxx' 和 "xxx"
        function_format = QTextCharFormat()
        # function_format.setFontItalic(False)
        function_format.setForeground(Qt.darkGreen)
        pattern = r'\'([^\]]*)\''
        self._add_mapping(pattern, function_format)
        pattern = r'\"([^\]]*)\"'
        self._add_mapping(pattern, function_format)
        # comment_format = QTextCharFormat()
        # comment_format.setBackground(QColor("#77ff77"))
        # self._add_mapping(r'^\s*#.*$', comment_format)

        font = QFontDatabase.systemFont(QFontDatabase.FixedFont)
        editor.setFont(font)
        self.setDocument(editor.document())
