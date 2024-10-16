# hook-openpyxl.py
hiddenimports = [
    'openpyxl.cell._writer',
    'openpyxl.styles.styleable',
    'openpyxl.styles.fills',
    'openpyxl.styles.numbers',
    'openpyxl.styles.protection',
    'openpyxl.styles.borders',
    'openpyxl.styles.colors',
    'openpyxl.styles.fonts',
    'openpyxl.styles.namedstyles',
    'openpyxl.styles.alignment',
    'openpyxl.styles.differential',
    'openpyxl.styles.mergedcell',
    'openpyxl.styles.styleable',
    'openpyxl.styles.styleable',
    'openpyxl.workbook._writer',
    'openpyxl.chart._writer',
    'openpyxl.drawing._writer',
    'openpyxl.utils.datetime',
    # 添加更多的子模块，如果需要的话
]

# 如果有需要，还可以添加其他隐藏导入