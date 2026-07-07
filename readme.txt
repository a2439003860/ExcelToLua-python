Excel转Lua转换工具

工具简介
这是一个将Excel文件自动转换为Lua配置文件的工具，适用于游戏开发、配置管理等场景。

核心功能
自动转换：将Excel表格转换为Lua表结构
数据类型支持：支持int、float、bool、string等基本类型
数组支持：支持一维数组（用[]表示）和二维数组（用{}表示）

智能跳过：自动跳过注释行和空列
智能检测
重复ID检测：自动检测并警告重复的ID值
空列过滤：跳过第三行参数为空的列
文件过滤：跳过包含$或#的文件名

项目根目录/
├── convert_excel_to_lua.bat  # 主批处理文件
├── excel_to_lua.py           # Python转换脚本
├── config.txt                # Lua复制开关和目标路径配置文件
├── excel/                    # Excel文件存放目录（自动创建）
├── log/                      # 日志输出目录（开启日志后自动创建）
└── lua/                      # 生成的Lua文件目录（自动创建）

使用说明
环境要求
Windows操作系统
Python 3.x
openpyxl库（首次运行会自动安装）

快速开始

双击运行 convert_excel_to_lua.bat

将Excel文件放入自动创建的 excel目录

工具会自动转换并在 lua目录生成对应文件

如果存在 config.txt 文件，且复制开关为1，工具会在转换完成后，将本次成功生成的Lua文件再复制一份到 config.txt 指定的目录。
复制时会默认覆盖目标目录中的同名文件。

当前 config.txt 默认内容为：
# 是否复制生成的Lua文件：0不复制，1复制
copy=1

# Lua复制目标目录，支持相对路径和绝对路径
target=..\APR\Assets\Data\Config

# 是否显示具体跳过了哪一列、哪一行：0不显示，1显示
show_skip_details=0

# 是否输出日志：0不输出，1输出
output_log=0

该路径为相对路径，表示复制到同级Unity工程 APR 的 Assets\Data\Config 目录。

修改复制开关：
直接打开 config.txt 文件，将 copy 改为0或1。
copy=0 表示不复制，只生成到本工具的 lua 目录。
copy=1 表示复制到 target 指定目录。

修改复制目标路径：
直接打开 config.txt 文件，将 target 改为需要复制到的目录即可。
支持相对路径和绝对路径。
相对路径以本工具所在目录（配置表目录）为基准。

示例：
copy=1
target=..\APR\Assets\Data\Config

copy=1
target=D:\Project\MyUnityGame\Assets\Data\Config

如果不需要额外复制，推荐将 config.txt 中的 copy 改为0。

跳过详情开关：
show_skip_details=0 表示只显示跳过统计。
show_skip_details=1 表示额外显示具体跳过了哪一列、哪一行。

日志开关：
output_log=0 表示不输出日志。
output_log=1 表示将本次转换时控制台显示的内容同时输出到 log 文件夹。
日志文件名格式为：系统日期_系统时间.log，精确到秒，例如 20260707_183045.log。

Excel文件格式要求
文件格式：.xlsx格式
文件命名：不能包含 $或 #字符

表格结构：
第1行：忽略（可用于注释）
第2行：数据类型定义
第3行：字段名（key）
第4行开始：数据行

目前支持：int、string、float、bool
支持数组、二维数组
当数据类型定义加入中括号时（例如[int]）转换数据时会视该数据列为一维数组，此时该数据列由英文逗号切分
当数据类型定义加入大括号时（例如{int}）转换数据时会视该数据列为二维数组，此时该数据列由英文逗号与分号切分

转换示例：
-- 自动生成的Lua配置文件
-- 源文件: example.xlsx

example = {
    [1] = {
        ID = 1,
        Name = "Alice",
        Scores = {90, 85, 95}
    },
    [2] = {
        ID = 2,
        Name = "Bob", 
        Scores = {80, 75, 85}
    }
}

return example

注意事项

转换规则
行跳过规则：第一列以/或?开头的行会被跳过
列跳过规则：第三行参数为空的列会被跳过
ID列检测：自动识别字段名为"ID"的列并进行重复检查

错误处理
工具会自动检测并显示重复ID警告
转换失败时会显示具体错误信息
开启 output_log 后，转换日志会保存到 log 文件夹
支持重新运行和查看输出目录

常见问题
Q: 工具无法运行怎么办？
A: 确保已安装Python 3.x，且网络连接正常（用于自动安装openpyxl）
Q: 生成的Lua文件有误？
A: 检查Excel文件格式是否符合要求，特别是数据类型标记是否正确
Q: 如何批量转换？
A: 只需将多个Excel文件放入excel目录，工具会自动批量处理

技术支持
如遇问题，请检查：
Excel文件格式是否符合要求
Python环境是否正常
控制台输出的错误信息
工具会在转换完成后显示详细的操作日志和警告信息，帮助您快速定位问题。
