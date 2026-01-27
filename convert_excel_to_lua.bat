@echo off
chcp 65001 >nul
title Excel转Lua转换器
color 0A

echo ========================================
echo      Excel转Lua转换工具
echo ========================================
echo.

REM 检查Python是否安装
python --version >nul 2>&1
if errorlevel 1 (
    echo 错误: 未找到Python，请先安装Python 3.x
    echo 可以从 https://www.python.org/downloads/ 下载
    pause
    exit /b 1
)

REM 检查openpyxl是否安装
python -c "import openpyxl" >nul 2>&1
if errorlevel 1 (
    echo 检测到未安装openpyxl库
    echo 正在自动安装openpyxl...
    echo.
    pip install openpyxl
    if errorlevel 1 (
        echo 安装失败，请手动运行: pip install openpyxl
        pause
        exit /b 1
    )
    echo.
)

REM 检查目录结构
if not exist "excel" (
    echo 创建excel目录...
    mkdir excel
    echo 请将Excel文件放入excel目录中
)

if not exist "lua" (
    echo 创建lua目录...
    mkdir lua
)

echo.
echo 正在转换Excel文件...
echo.

REM 运行Python脚本
python excel_to_lua.py

echo.
echo ========================================
echo 转换完成！
echo ========================================
echo.

REM 检查是否有文件被转换
set "has_lua_files=0"
if exist lua\*.lua (
    set "has_lua_files=1"
    echo 生成的Lua文件:
    dir /b lua\*.lua
    echo.
)

REM 检查是否有Excel文件
set "has_excel_files=0"
if exist excel\*.xlsx (
    set "has_excel_files=1"
    echo Excel目录中的文件:
    dir /b excel\*.xlsx
    echo.
)

if %has_lua_files%==0 (
    echo 提示: 未生成任何Lua文件
    echo 请确保:
    echo   1. Excel文件已放入excel目录
    echo   2. Excel文件为.xlsx格式
    echo   3. 文件不包含$或#字符
    echo.
)

REM 如果有重复ID警告，高亮显示
if exist "excel" (
    python -c "
import os
import re
from openpyxl import load_workbook

excel_dir = 'excel'
duplicate_warnings = []

for filename in os.listdir(excel_dir):
    if filename.endswith('.xlsx') and '$' not in filename and '#' not in filename:
        filepath = os.path.join(excel_dir, filename)
        try:
            wb = load_workbook(filename=filepath, data_only=True)
            ws = wb.active
            
            # 查找ID列
            id_col = None
            for col in range(1, ws.max_column + 1):
                key_cell = ws.cell(row=3, column=col).value
                if key_cell and str(key_cell).strip().lower() == 'id':
                    id_col = col
                    break
            
            if id_col:
                id_values = []
                for row in range(4, ws.max_row + 1):
                    first_cell = ws.cell(row=row, column=1).value
                    if first_cell and str(first_cell).strip().startswith(('/', '?')):
                        continue
                    
                    id_cell = ws.cell(row=row, column=id_col).value
                    if id_cell is not None:
                        id_values.append(id_cell)
                
                # 检查重复
                from collections import Counter
                id_counts = Counter(id_values)
                duplicates = [id_val for id_val, count in id_counts.items() if count > 1]
                
                if duplicates:
                    duplicate_warnings.append((filename, duplicates))
                    
        except Exception as e:
            pass

if duplicate_warnings:
    print('\n[警告] 发现以下重复ID:')
    for filename, dup_ids in duplicate_warnings:
        print(f'  {filename}: {set(dup_ids)}')
    print('\n注意: 重复ID可能导致数据覆盖，请检查Excel文件！')
" 2>nul
)

echo.
echo 操作:
echo 1. 按回车键重新运行转换
echo 2. 输入 q 退出
echo 3. 输入 o 打开lua目录

:choice
set /p choice=请选择: 
if "%choice%"=="q" exit /b 0
if "%choice%"=="Q" exit /b 0
if "%choice%"=="o" goto open_folder
if "%choice%"=="O" goto open_folder
goto rerun

:open_folder
if exist lua (
    explorer lua
)
goto choice

:rerun
cls
call %0