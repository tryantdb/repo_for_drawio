要把 make_mipi.py 打包成可执行的 .exe 且尽量小，核心思路是：减少依赖体积 + 用合适的打包器 + 开启压缩/剔除选项。下面给你一套“能用即小”的实操方案，以及“更小”的替代方案。


---

方案 A（最省事）：PyInstaller（保留 pandas 版本）

> 适合：不想改代码。
体积：一般几十到上百 MB（pandas+numpy 体积本身较大）。



1) 在干净虚拟环境打包

python -m venv .venv
.venv\Scripts\activate
pip install --upgrade pip
pip install pandas openpyxl pyinstaller

2) 一次性命令（单文件 + 压缩尽量精简）

pyinstaller make_mipi.py ^
  --onefile ^
  --noconsole ^
  --clean ^
  --name mipi ^
  --exclude-module numpy.random._examples ^
  --exclude-module tests ^
  --exclude-module pytest ^
  --exclude-module IPython ^
  --exclude-module pip

可选：安装 UPX 后再加 --upx-dir <UPX目录> 开启可执行文件压缩（能再缩一些体积）。

3) 精细化（可选，用 .spec）

生成一次 pyinstaller make_mipi.py 得到 make_mipi.spec，然后在 .spec 里加 excludes=[...]，把用不上的模块剔除（如 dateutil.tzwin, matplotlib, scipy 等），再：


pyinstaller --clean --onefile --noconsole make_mipi.spec

4) 小技巧

使用较新的 Python（3.11/3.12）和干净 venv，能避免历史残留包被打进来。

--noconsole 仅隐藏黑窗，不影响体积；体积关键在依赖。

如需日志窗口可去掉 --noconsole。



---

方案 B（更小）：改用 openpyxl 实现（避免 pandas+numpy 体积）

> 适合：能小改代码以显著缩小 exe。
体积：通常可压到 10–20MB 左右（随 Python 版本与 UPX 有差异）。



把 make_mipi.py 改写为只用 openpyxl（功能不变：读取 register，写到新表，表名用默认 Sheet1）。

# make_mipi.py （openpyxl 版，更精简）
from pathlib import Path
import sys
from openpyxl import load_workbook, Workbook

def main(cfg_path="config.xlsx", src_sheet="register", out_path="mipi.xlsx"):
    cfg = Path(cfg_path)
    if not cfg.exists():
        sys.exit(f"找不到文件：{cfg.resolve()}")

    wb_in = load_workbook(cfg, data_only=True)  # 读取计算后的值
    if src_sheet not in wb_in.sheetnames:
        sys.exit(f"读取工作表失败：找不到名为 {src_sheet} 的表")

    ws_in = wb_in[src_sheet]

    # 新建工作簿（默认有 Sheet，正是 Excel 默认名）
    wb_out = Workbook()
    ws_out = wb_out.active  # 默认 Sheet
    ws_out.title = "Sheet"  # 注意：openpyxl 默认是 "Sheet"，Excel 打开会显示 "Sheet1"
    # 若明确希望叫 "Sheet1"，请改为：ws_out.title = "Sheet1"

    # 复制值（不复制格式/合并/批注）
    for row in ws_in.iter_rows(values_only=True):
        ws_out.append(list(row))

    wb_out.save(out_path)
    print(f"已生成：{Path(out_path).resolve()}（工作表名：{ws_out.title}）")

if __name__ == "__main__":
    import sys
    args = sys.argv[1:]
    if len(args) >= 3:
        main(args[0], args[1], args[2])
    elif len(args) == 2:
        main(args[0], args[1])
    elif len(args) == 1:
        main(args[0])
    else:
        main()

> 说明：

openpyxl 默认工作表名是 "Sheet"。如果你必须是 "Sheet1"，把上面标注处改成 ws_out.title = "Sheet1" 即可。

你原来用 pandas 的默认写出名是 "Sheet1"，这里需要手动对齐。




打包命令（体积显著更小）：

python -m venv .venv
.venv\Scripts\activate
pip install --upgrade pip
pip install openpyxl pyinstaller

pyinstaller make_mipi.py ^
  --onefile ^
  --noconsole ^
  --clean ^
  --name mipi ^
  --exclude-module tests ^
  --exclude-module pytest ^
  --exclude-module pip

如果装了 UPX，再加 --upx-dir <UPX目录>。


---

方案 C（更极致）：Nuitka 原生编译

> 适合：追求更小或性能更好，能接受更长编译时间。
体积：通常比 PyInstaller 更小，尤其在不使用 pandas 的情况下。



步骤：

python -m venv .venv
.venv\Scripts\activate
pip install --upgrade pip
pip install nuitka openpyxl  # 若用 pandas 也装上，但体积会大

python -m nuitka --onefile --standalone ^
  --nofollow-imports ^
  --clang --lto=yes ^
  --remove-output ^
  --output-filename=mipi.exe ^
  make_mipi.py

--nofollow-imports 会严格只打进被静态可达的依赖（对“精简代码依赖”很友好）。

用 pandas 时不建议加 --nofollow-imports（pandas 的动态导入较多，可能需 --include-module= 指明）。

可选：加 --windows-console-mode=disable 隐藏控制台。



---

选型建议（快速决策）

马上能跑且不改代码：用 PyInstaller + pandas（方案 A）。

明显缩小体积：把脚本改成 openpyxl（方案 B）→ 再用 PyInstaller。

进一步压缩/优化：改完 openpyxl 后用 Nuitka（方案 C）。



---

常见问题 & 排坑

体积还是很大？
九成是因为 pandas/numpy。能不用就不用；或加 UPX；或切到 Nuitka。

杀软误报？
给 exe 做代码签名能大幅降低误报率；或用企业内部分发白名单。

运行时缺 DLL？
确保用的打包器在同一台机器打包和测试；必要时用 Dependency Walker/ Process Monitor 排查。

需要复制格式/合并单元格？
openpyxl 也能做，但要额外拷列宽、行高、合并范围、样式，这会增加代码复杂度（对体积影响不大）。



---

如果你确定要保持默认工作表名为 Sheet1，我可以把上面 openpyxl 代码微调成 Sheet1 并给你一条完整的 Nuitka 命令行，做到更小。你更倾向于哪条路线？

