# 报关资料生成器

根据进口Invoice PDF自动生成报关资料（合同、报关单、申报要素）的工具。

## 功能特性

- **PDF Invoice解析**: 自动从PDF中提取发票信息
- **模板生成**: 使用Word模板生成合同、报关单和申报要素
- **智能匹配**: 自动匹配供应商和产品信息
- **可配置**: 通过YAML配置文件管理供应商和产品映射

## 项目结构

```
customs_doc_generator/
├── main.py                    # 主程序
├── pdf_parser.py              # PDF解析模块
├── docx_generator.py           # Word文档生成模块
├── config_loader.py           # 配置加载模块
├── requirements.txt           # 依赖包
├── data/
│   ├── supplier_product_mapping_import.yaml  # 进口供应商和产品配置
│   ├── 进口商品识别规则.xlsx                    # 进口供应商描述/料号到内部 ITEM 的规则
│   ├── item和品名对应表.xlsx                    # 与出口申报要素对应表同结构的内部 ITEM 表
│   └── supplier_product_mapping_export.yaml  # 出口供应商和产品配置
├── templates/
    ├── CONTRACT_CAMARI_TEMPLATE.xlsx  # 合同（买方固定 CAMARI，优先使用）
    ├── CONTRACT.xlsx                  # 旧版合同（兼容）
    ├── FedEx报关单模板.xlsx
    └── 申报要素.docx
└── export_templates/
    ├── 出口申报要素对应表.xlsx
    ├── export_declaration.xlsx
    └── 申报要素总汇.docx
```

## 安装

```bash
pip install -r requirements.txt
```

## Invoice 可编辑 PDF 模板（macOS 预览可修改）

为了以后在 macOS `预览`里直接点进去修改（不再每次解析 PDF），项目内已新增一个生成脚本，可生成带表单字段的 `invoice_form_template.pdf`（AcroForm）。

模板位置：
- `templates/invoice_form_template.pdf`（最多支持 10 行明细）

如果你想重新生成模板：
```bash
cd /Users/michael/customs_doc_generator
.venv/bin/python scripts/generate_invoice_form_template_pdf.py --output templates/invoice_form_template.pdf --max-rows 10
```


## 使用方法

### 交互模式

```bash
cd customs_doc_generator
python main.py -i
```

### 命令行模式

```bash
python main.py /path/to/invoice.pdf -s supplier_001 -p product_001 -o output
```

### 参数说明

- `invoice`: Invoice PDF文件路径
- `-c, --config`: 配置文件路径（默认: data/supplier_product_mapping_import.yaml）
- `-t, --templates`: 模板目录路径（默认: templates）
- `-s, --supplier`: 供应商代码
- `-p, --product`: 产品代码
- `-o, --output`: 输出目录（默认: output）
- `-i, --interactive`: 交互模式

## 配置说明

进口建议按“两层”维护：

1. `data/进口商品识别规则.xlsx`：维护供应商发票里的描述、料号、成分、hide type 如何匹配到内部 ITEM。
2. `export_templates/出口申报要素对应表.xlsx`：维护内部 ITEM 对应的申报品名、HS Code、申报要素。
3. `data/supplier_product_mapping_import.yaml`：维护进口供应商资料，以及少量仍未迁移到内部 ITEM 表的进口专用兜底品类。

`data/item和品名对应表.xlsx` 与出口申报要素对应表保持同结构，作为旧代码兼容表；新的进口/出口逻辑都优先围绕内部 ITEM 表工作。

### 供应商信息

```yaml
suppliers:
  supplier_001:
    name: "供应商名称"
    country: "国家"
    address: "地址"
    contact: "联系人"
    trade_term: "CIF"  # 贸易条款
    payment_term: "T/T 30天"  # 结算方式
```

### 产品信息

```yaml
product_categories:
  product_001:
    name: "产品名称"
    category: "产品类别"
    sub_category: "子类别"
    declaration_elements:
      brand_type: "无"
      model: "见货物"
      usage: "用途"
      composition: "成份"
      voltage: "电压"
    legal_unit: "个"  # 法定单位
    second_unit: "千克"  # 第二法定单位
    supervision: "无"  # 监管条件
    inspection: "L"  # 检验检疫类别
```

### 进口商品识别规则

`data/进口商品识别规则.xlsx` 的常用列：

- `供应商`: 填 YAML 里的供应商代码，可用逗号分隔；留空表示所有供应商
- `匹配字段`: `item_code`、`item_code_prefix`、`description`、`composition`、`hide_type`、`article_name`、`任意文本`
- `匹配方式`: `包含`、`等于`、`前缀`、`正则`
- `匹配内容`: 多个关键词可用英文分号或逗号分隔
- `内部ITEM / YAML品类`: 优先填写 `export_templates/出口申报要素对应表.xlsx` 的 `ITEM on Invoice`；也兼容 `supplier_product_mapping_import.yaml` 里的 `product_categories` key
- `优先级`: 数字越大越先匹配

如果进口商品没有识别到 HS Code 或申报品名，程序会在输出目录生成 `待维护进口商品_*.xlsx`，把里面的信息补成新规则即可。

## 合同 Excel 模板（CONTRACT_CAMARI_TEMPLATE.xlsx）

- 买方固定为 **CAMARI TRADING (ZHEJIANG) CO., LTD**（第4行）。
- **合同编号** = Invoice 号（`S/C NO` 右侧单元格）。
- 第 **12** 行填写 **Currency / Incoterms / Payment terms**（由程序根据发票与 YAML 写入）。
- 第 **15** 行起为商品明细（品名、数量、单价、金额）；**K 列**为总金额行 `TOT AMOUNT: 金额 币种`。
- 保留原模板 **(5)～(8) 条款**及底部 **买方/卖方** 签章区。

若你更新了 `CONTRACT.xlsx` 版式，可重新生成 CAMARI 模板：

```bash
python3 scripts/build_contract_camari_template.py
```

## Word模板准备

在 `templates` 目录下准备以下模板：

1. **合同模板.docx** - 使用 `{{字段名}}` 占位符（若仅使用 Excel 合同可忽略）
2. **报关单模板.docx** - 同上
3. **申报要素模板.docx** - 同上

支持的占位符包括：
- `{{INVOICE_NO}}` - 发票号
- `{{INVOICE_DATE}}` - 发票日期
- `{{SUPPLIER_NAME}}` - 供应商名称
- `{{TRADE_TERM}}` - 贸易条款
- `{{COUNTRY_OF_ORIGIN}}` - 原产国
- `{{PRODUCT_NAME}}` - 产品名称
- `{{BRAND_TYPE}}` - 品牌类型
- 等...

## 注意事项

1. PDF解析基于正则匹配，可能需要根据实际Invoice格式调整解析规则
2. 模板中的占位符需要与配置和产品信息匹配
3. 首次使用建议先在交互模式下测试
