# macOS Intel 打包说明

这份源码包用于在 Intel 芯片 Mac 上生成对应的 `ClearanceOS.app`。

## 前置条件

1. 安装 Python 3.11 或更新版本。
2. 如果需要 OCR 功能，Intel Mac 上还需要安装 `ocrmypdf` 和 `tesseract`。

## 打包步骤

1. 解压 `mac_intel_build_source.zip`。
2. 打开 Terminal，进入解压后的文件夹。
3. 执行：

```bash
chmod +x build_app.sh
./build_app.sh
```

4. 生成文件位于：

```text
dist/ClearanceOS.app
```

把这个 `.app` 发给 Intel Mac 员工即可。

## 首次打开

如果 macOS 提示无法打开，右键 `ClearanceOS.app`，选择“打开”，再确认一次。

## 更新模板

如需更新 mapping 或模板，重新替换源码里的这些文件后再运行 `./build_app.sh`：

- `data/supplier_product_mapping_import.yaml`
- `data/进口商品识别规则.xlsx`
- `data/supplier_product_mapping_export.yaml`
- `export_templates/出口申报要素对应表.xlsx`
- `templates/`
- `export_templates/`
