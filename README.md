# 签名小工具

按 Excel 名单批量生成手写风签名图片（随机字体 + 随机笔迹），并导出内嵌签名的 Excel。

![icon](assets/icon.png)

## Windows 下载

两种方式（GitHub Actions 自动构建）：

1. **安装包**：到 [Releases](../../releases) 下载 `SignatureTool-Setup-x.x.x.exe`，双击安装
2. **绿色版**：Releases 下载 `SignatureTool-...-win64-portable.zip`，解压双击 `SignatureTool.exe` 即可，无需安装

> 每次推送到 main 分支都会自动重新构建，可在 [Actions](../../actions) 页面下载最新产物。

## macOS 运行

```bash
pip3 install -r requirements.txt
python3 signature_tool.py
```

或直接双击 `start.command`。

## 使用步骤

1. **选择 Excel 模板** —— 仅支持 `.xlsx`，第一个 sheet 的表头里需要有姓名列（工具会自动识别"姓名/名字/Name"等关键字，也可手动下拉选择）。
2. **调整参数**（可全部保持默认）：
   - 画布尺寸：签名图片的像素大小
   - 字体：勾选参与随机的字体，每个姓名随机抽一款
   - 墨色：随机 / 纯黑 / 碳黑 / 蓝黑 / 深蓝
   - 倾斜、连笔紧凑、笔画加粗、随机下划线
3. **生成预览** —— 预览显示前 24 个姓名；点【重新随机 ⟳】换一批笔迹。同名重渲染结果一致（预览即所得）。
4. **导出**：
   - `导出 Excel（内嵌签名）`：输出 `<原名>_含签名.xlsx`，在表格末尾新增"签名"列并嵌入图片，行高列宽自动适配；同时输出图片文件夹。
   - `导出 Word（套模板）`：选择含「序号 / 签到」表格的 Word 模板（如签到表），签名按**序号**自动填入对应行的"签到"格。签名高度默认 0.8cm（受模板固定行高限制，可 0.4~1.2cm 调整）。
   - `仅导出图片文件夹`：每行一个 `序号_姓名.png`（透明底，可直接盖到文档上）。

### Word 模板说明

- 模板表格需有表头行，包含「序号」列和「签到 / 签名」列（支持一行多组，如左栏 1-25、右栏 26-50 的双栏签到表）
- 名单 Excel 中如有「序号」列会自动识别并按序号落位；没有则按名单顺序 1, 2, 3… 对号
- 若模板行高为固定值，签名会自动压缩到行高内，不会被裁切

## 字体

内置 5 款开源手写字体（SIL OFL 协议，来自 Google Fonts），已放在 `fonts/` 目录：

| 文件 | 风格 |
|---|---|
| MaShanZheng-Regular.ttf | 马善政毛笔楷书 |
| ZhiMangXing-Regular.ttf | 志莽行书 |
| LongCang-Regular.ttf | 龙藏行草 |
| LiuJianMaoCao-Regular.ttf | 柳建毛草 |
| ZCOOLXiaoWei-Regular.ttf | 站酷小薇 |

想扩充字体池：把任意 `.ttf/.otf/.ttc` 中文手写体丢进 `fonts/` 目录，重启工具即可被扫描到。

## 注意

- 输出 Excel 由模板数据重建，**原文件中已有的图片/图表不会保留**（openpyxl 限制）。
- 姓名、序号会作为图片文件名，姓名中的非法字符会自动替换为 `_`。
- 建议直接用 Excel/WPS 打开输出文件检查签名效果。

## 依赖

- Python 3.9+
- pillow、openpyxl
