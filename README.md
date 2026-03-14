# 凭证生成工具

将发票和银行流水一键转换为金蝶/用友标准凭证导入格式。

## 快速启动

### Windows
双击 `run.bat`，浏览器会自动打开工具页面。

### Mac / Linux
```bash
chmod +x run.sh
./run.sh
```

### 手动启动
```bash
pip install -r requirements.txt
streamlit run app.py
```
然后在浏览器打开 http://localhost:8501

---

## 使用步骤

1. **上传文件**
   - 科目列表（必须）
   - 全量发票导出（应收，可选）
   - 进项发票信息（应付，可选）
   - 活期交易明细（银行，可选）

2. **填写凭证号**（日期自动从文件提取，无需手动填）

3. **点击"开始生成"**

4. **下载 Excel 文件**，黄色行需手动确认科目

---

## 文件结构

```
voucher_app/
├── app.py                    # 主界面
├── processor/
│   ├── ar.py                 # 应收凭证逻辑
│   ├── ap.py                 # 应付凭证逻辑
│   └── bank.py               # 银行凭证逻辑
├── utils/
│   ├── subject.py            # 科目匹配
│   ├── excel_writer.py       # Excel 输出（含黄色标注）
│   └── date_utils.py         # 日期工具
├── requirements.txt
├── run.bat                   # Windows 启动脚本
└── run.sh                    # Mac/Linux 启动脚本
```

---

## 注意事项

- 科目列表必须包含"编码"和"名称"两列
- 银行流水需为农商行标准格式，**表头在第5行**
- 应付凭证费用科目根据"用途"列或供应商名称关键词自动匹配
- 无法匹配的行会在输出 Excel 中**标黄**，请下载后手动核对

## 环境要求

- Python 3.10+
- 依赖见 requirements.txt（streamlit, pandas, openpyxl, xlrd）
