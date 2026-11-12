# xlsm/xlsx 内部结构与解析指南

openpyxl 报错、超慢或直接崩溃时，用本指南配合 `scripts/parse_xlsm.py`（纯标准库）完成解析。

## 1. 什么时候绕开 openpyxl

| 症状 | 常见原因 | 处理 |
| --- | --- | --- |
| `InvalidFileException` / `BadZipFile` | 加密文件（OLE 容器）、旧版 .xls、文件损坏 | 先看文件头 8 字节：`D0 CF 11 E0` = OLE（加密或 .xls），`PK` = zip。加密文件让用户去密码后另存 |
| 打开 10 秒以上 / 内存暴涨 | 大型财务模型（>1MB 的 xlsm 常见），openpyxl 会构建完整对象模型 | 用 parse_xlsm.py，iterparse 流式读，内存只保留当前行 |
| `KeyError: xl/_rels/workbook.xml.rels` 等关系错误 | 关系部件缺失/损坏（异常保存、同步盘截断） | parse_xlsm.py 会回退到按编号顺序猜 worksheet 部件 |
| 要看"文件里有哪些宏" | openpyxl 只给 keep_vba 布尔值，不给模块清单 | `parse_xlsm.py 文件.xlsm --vba` |
| WPS 产生的怪文件 | 命名空间变化、非标准扩展部件 | 解析器统一按"本地名"匹配标签，不依赖具体命名空间 |

## 2. xlsx/xlsm 就是一个 zip

```
 workbook.xlsx (zip)
├── [Content_Types].xml        部件类型声明
├── _rels/.rels                根关系 → xl/workbook.xml
└── xl/
    ├── workbook.xml           工作表清单 + 定义名称
    ├── _rels/workbook.xml.rels  rId → worksheets/sheet1.xml 等映射
    ├── sharedStrings.xml      共享字符串池（t="s" 的下标指向这里）
    ├── styles.xml             样式：cellXfs 索引 → numFmtId → 格式码
    ├── worksheets/sheet1.xml  单元格 <c r="A1" t=".." s=".."><v>/<f>/<is>
    └── vbaProject.bin         仅 xlsm/xltm：VBA 工程（OLE 复合文档）
```

单元格的完整取值流程：`<c>` 的 `r` 定位（列字母→列号），`t` 定类型
（`s`=共享字符串下标 / `str`=公式字符串结果 / `inlineStr`→`<is><t>` /
`b`=布尔 / `e`=错误如 `#DIV/0!` / 缺省=数字），`s` 定样式 → 查 styles.xml
判断是否日期格式。

## 3. 日期：Excel 序列号 → datetime（1900 系统）

Excel 把日期存为序列号：1 = 1900-01-01。注意它虚构了 1900-02-29：

- 序列号 ≥ 61：`datetime(1899,12,30) + timedelta(days=serial)`
- 序列号 = 60：不存在的 1900-02-29（按 1900-03-01 处理并留注释）
- 序列号 1..59：`datetime(1899,12,31) + timedelta(days=serial)`

是否按日期解释由样式决定：内置格式 id 14-22、27-36、45-47、50-58 是日期
时间；自定义格式码剥离引号字面量、`\` 转义与 `[..]` 前缀后含 `ymdhs` 的
也视为日期（parse_xlsm.py 已内置该判断）。

## 4. VBA：vbaProject.bin 是 OLE 复合文档

结构要点（CFB/OLE）：

- 文件头 512 字节含扇区大小、FAT 位置、目录起始扇区
- 目录是红黑树，条目 128 字节：UTF-16LE 名字、类型（1=存储 2=流）、起始扇区、大小
- <4096 字节的流放在"迷你流"里，经 miniFAT 寻址；parse_xlsm.py 实现了完整寻址
- 宏模块名可以直接读明文流 `PROJECT`（`Module=名字; \名字`、`Document=Sheet1/&H00000000` 等行）
- 模块**代码**在 `VBA/<模块名>` 流里，经 MS-OVBA 压缩，本工具不解压；
  需要代码内容时用 Excel/WPS 打开按 Alt+F11 查看，或在用户明确要求下
  安装 olefile 进一步处理

安全红线：只读解析，绝不执行、不启用宏，不引导用户"先启用宏再试试"。

## 5. 大文件：iterparse 流式读取

```python
import zipfile
from xml.etree import ElementTree as ET

with zipfile.ZipFile("大模型.xlsm") as zf:
    with zf.open("xl/worksheets/sheet1.xml") as f:
        for _, el in ET.iterparse(f, events=("end",)):
            if el.tag.rsplit('}',1)[-1] == "row":
                ...  # 处理行
                el.clear()   # 必须清理，否则内存照样涨
```

parse_xlsm.py 即按此模式实现：百 MB 级工作表也能读，峰值内存只与单行宽度有关。

## 6. 共享公式与已知局限

- `<f t="shared" si="3">` 从属单元格只存下标：值 `<v>` 一般仍在，取值不受影响；
  公式文本需回溯主单元格，parse_xlsm.py 输出"主单元格公式"，引用未做平移
- 数据透视表、图表只体现在 zip 部件清单里，不解析内容
- 加密工作簿无法也不应破解：提示用户去密码另存

## 7. parse_xlsm.py 命令速查

```bash
python3 scripts/parse_xlsm.py 文件.xlsm --list            # 总览
python3 scripts/parse_xlsm.py 文件.xlsm --sheet Sheet1 --head 20
python3 scripts/parse_xlsm.py 文件.xlsm --json --formulas > 数据.json
python3 scripts/parse_xlsm.py 文件.xlsm --all --outdir 解析结果/
python3 scripts/parse_xlsm.py 文件.xlsm --vba
```

JSON 结构：`{file, size_bytes, sheets:[{name, max_row, max_col,
merged_cells, rows(二维数组，日期为 ISO 字符串), formulas?}],
defined_names, vba:{has_vba, modules, note}, zip_entries}`。
