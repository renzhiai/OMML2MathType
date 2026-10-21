# OMML2MathType

把 Word 内置公式（OMML / 「插入 → 公式」写的那种）批量转换成 **MathType 公式对象**，双击即可用 MathType 编辑。

一条命令、33 个公式、60 秒，0 失败。

---

## 1. 环境要求

| 项 | 要求 |
|---|---|
| 系统 | Windows |
| Office | Word（2010 及以上，装了 MathType 的 Word 加载项） |
| MathType | 6.x / 7.x，且 Word 启动时有 `MathType Commands` 加载项 |
| Python | 3.9+，需要 `python-docx`、`pywin32`、`lxml` |

缺包的话：

```bat
python -m pip install python-docx pywin32 lxml
```

---

## 2. 三种用法

### 2.1 拖拽（最省事）

把 `.docx` 文件直接拖到 **`转换公式.bat`** 上（可以一次拖多个）。

- 原文件会被就地覆盖，同时自动生成备份 `原名.docx.bak-时间戳.docx`
- 转换完窗口不会自动关，能看到结果和告警

### 2.2 命令行

```bat
python omml2mathtype.py 论文.docx
```

默认输出 `论文_MathType.docx`，原文件不动。

常用组合：

```bat
:: 覆盖原文件（自动备份）
python omml2mathtype.py 论文.docx --inplace

:: 指定输出路径
python omml2mathtype.py 论文.docx -o D:\out\论文_mt.docx

:: 先看看公式都能转成什么 TeX（不改任何文件）
python omml2mathtype.py 论文.docx --dry-run

:: 自定义超宽告警阈值（磅），默认取版心的 82%
python omml2mathtype.py 论文.docx --max-width 340

:: 保留中间文件，便于排查
python omml2mathtype.py 论文.docx --keep-work
```

| 参数 | 说明 |
|---|---|
| `input` | 输入 .docx |
| `-o, --output` | 输出路径 |
| `--inplace` | 覆盖原文件（默认备份，`--no-backup` 可关） |
| `--dry-run` | 只打印每个公式的 TeX，不动文件 |
| `--max-width` | 公式宽度告警阈值（磅） |
| `--progid` | MathType 的 OLE ProgID，默认 `Equation.DSMT4` |
| `--keep-work` | 保留中间文件 `_work_*.docx` |
| `--report` | 审计报告 JSON 路径 |
| `-q, --quiet` | 精简输出 |

### 2.3 当库用

```python
from omml2tex import omml_to_tex
from docx import Document

M = "{http://schemas.openxmlformats.org/officeDocument/2006/math}"
doc = Document("论文.docx")
for om in doc.element.body.iter(M + "oMath"):
    print(omml_to_tex(om))       # → LaTeX
```

`omml_to_tex.py` 单独也能跑：

```bat
python omml2tex.py 论文.docx
```

只依赖 `lxml`，不需要 Word，不需要 MathType —— 适合做公式文本检索、查重、导出 Markdown/LaTeX。

---

## 3. 为什么必须分两段

这是整个工具的关键，绕不开：

1. **在 Word 里不能直接改公式的文字。** 公式里的 `Range.Text` 仍处在数学区内，改完还是 Word 公式，触发不了 MathType 的 TeX 切换。所以先回到 XML 层，用 `python-docx` 把每个 `m:oMath` 整个替换成普通文本 run `$TeX$`，让它彻底脱离数学区。
2. **再由 Word 逐段扫 `$...$`**，选中后执行 MathType 的宏 `MathTypeCommands.UILib.MTCommand_TeXToggle`，就地生成 `Equation.DSMT4` 对象。

顺带解决了三个坑：

- **运算符丢失**：OMML 会把 `\left(a-b\right)` 里的 `−`、`+` 当成 delimiter 的 `sepChr` 并拆成多个 `m:e`，按常规拼接这些加减号会消失。工具按 `sepChr` 还原。
- **多行结构**：`cases`（分段函数）、两行 `array`、矩阵、根式、重音（`\dot` / `\ddot` / `\bar`）都能正确转成 MathType 的分段/数组结构。
- **正文里的 `$`**：如果正文本身有美元符号，会被误当公式定界符。工具会临时换成全角 `＄`，转换完自动还原。

---

## 4. 输出与自检

跑完会打印审计结果，并写一份 JSON：

```
[3/4] 审计公式尺寸 …
      页数 19｜公式对象 33｜版心 413.9 pt｜告警阈值 339.4 pt
      ⚠ 有 1 个公式超出版心，可能把式号挤到下一行：
         式 (4.14) 宽 420.0 pt，约超出 80.6 pt
```

**告警了怎么办**：MathType 渲染普遍比 Word 公式宽 5%~8%，超宽会把右侧式号挤到下一行。MathType 对象改宽度存盘会被复位，只能改内容 —— 用 `--dry-run` 找到那条 TeX，手动折成两行：

```latex
\begin{array}{l}第一行 \\ 第二行\end{array}
```

改完再跑一次即可。

---

## 5. 已知限制

- **页眉页脚、脚注、文本框**里的公式不处理（只扫正文 body）。
- **MathType 语法报错会弹窗**，需要手动点掉；TeX 拼错时才会出现。
- **只做单向**：OMML → MathType。反向（MathType → Word 公式）请用 MathType 自带的「Convert Equations」。
- **转换不可逆**：`--inplace` 前务必保留备份（默认已备份）。
- 文档若正被 Word 打开，保存会冲突 —— 先关掉，或转换后选「不保存」再重开。

---

## 6. 排错

| 现象 | 原因 / 处理 |
|---|---|
| `未找到 Word 公式（OMML）` | 文档里本来就没有 Word 公式，或公式是图片/MathType 对象 |
| 全部失败、提示宏不存在 | Word 没装 MathType 加载项。检查 Word「文件 → 选项 → 加载项」里是否有 MathType Commands |
| 部分失败，日志里有 TeX | 该条 TeX 语法 MathType 不认，用 `--dry-run` 对照修改 |
| 卡住不动 | MathType 弹了错误对话框，切到那个窗口点掉 |
| 页数变了 | 正常，MathType 对象行高与 Word 公式略有差异 |

多次运行的失败明细追加在 `_run.log`，报告在 `_report_<时间戳>.json`。

---

## 7. 目录

```
OMML2MathType/
├── omml2mathtype.py      主程序（提取 → 转换 → 审计 → 落盘）
├── omml2tex.py           OMML → LaTeX 转换核心（可独立使用）
├── 转换公式.bat           拖拽入口（就地转换 + 备份）
├── 安装右键菜单.bat       在 .docx 右键菜单加一项
├── 卸载右键菜单.bat       移除该项
└── README.md
```
