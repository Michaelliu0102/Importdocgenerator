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
│   └── supplier_product_mapping.yaml  # 供应商和产品配置
└── templates/
    ├── CONTRACT_CAMARI_TEMPLATE.xlsx  # 合同（买方固定 CAMARI，优先使用）
    ├── CONTRACT.xlsx                  # 旧版合同（兼容）
    ├── FedEx报关单模板.xlsx
    └── 申报要素.docx
```

## 安装

```bash
pip install -r requirements.txt
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
- `-c, --config`: 配置文件路径（默认: data/supplier_product_mapping.yaml）
- `-t, --templates`: 模板目录路径（默认: templates）
- `-s, --supplier`: 供应商代码
- `-p, --product`: 产品代码
- `-o, --output`: 输出目录（默认: output）
- `-i, --interactive`: 交互模式

## 配置说明

编辑 `data/supplier_product_mapping.yaml` 来配置：

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
