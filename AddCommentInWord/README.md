# 批量作业评语工具

项目简介：  
根据 Excel 表格中的学号和评语，自动将其插入到对应的 Word 页眉或 PDF 水印中。

目录结构
- homeworks/: 【这里放输入文件】 把你的 Excel 和学生作业扔到这里。
- processed/: 【这里是输出结果】 程序运行后，处理好的文件会出现在这里。
- config.py: 修改文件路径配置。
- auto_comment.py: 主程序。
  
使用步骤
1. 安装依赖库: pip install -r requirements.txt
2. 确保 examples 文件夹里有 comments.xlsx 和作业文件。
3. 运行 auto_comment.py。
4. 去 processed 文件夹查看结果。

## 输入文件要求  
为了确保程序能正确识别文件并匹配评语，请严格遵守以下规则：
Excel 评语表 (.xlsx 或 .xls)

    格式：必须包含至少两列数据
    列顺序：
        第 1 列 (A列)：学生学号（作为唯一识别码）
        第 2 列 (B列)：评语内容
    注意：
        程序默认读取 Excel 的第一个 Sheet
        请去除所有合并单元格
        学号格式（文本或数字）均可，程序会自动处理

学生作业文件

    支持格式：.docx (Word文档) 和 .pdf
    命名规则：文件名中必须包含学号数字
        ✅ 正确示例：
            202301_张三.docx
            张三-202301-期末.pdf
            202301.docx
        ❌ 错误示例：
            张三.docx (程序无法识别这是哪个学号)
            期末作业.pdf
    存放位置：所有待处理文件需集中存放在同一个文件夹内

### 🛠️ 高级配置（修改字体与位置）
如果对评语的位置、字体或颜色不满意，可用记事本或 PyCharm 打开 auto_comment.py 进行微调。
1. 修改 PDF 设置
PDF 的修改集中在 process_pdf 函数中：

    更换字体（解决乱码）
    在代码顶部找到 FONT_PATH 变量，默认指向黑体：C:\Windows\Fonts\simhei.ttf。
    想换字体只需去 C:\Windows\Fonts 复制对应字体文件的路径替换即可。
    调整评语位置
    搜索 c.drawString(300, 750, …) 这一行：
        第一个数字越大，文字越靠右；
        第二个数字越大，文字越靠上（0 为页面最底部）。
        可边改边跑，实时预览效果。

2. 修改 Word 设置
Word 的修改集中在 process_word 函数中：

    对齐方式
    找到 WD_ALIGN_PARAGRAPH.RIGHT，可改为 .CENTER（居中）或 .LEFT（左对齐）。
    颜色与字号
        颜色：RGBColor(255, 0, 0) 为红色，(0, 0, 0) 为黑色，三值对应 R/G/B。
        字号：Pt(12) 中的数字即为磅值，越大字越大。
