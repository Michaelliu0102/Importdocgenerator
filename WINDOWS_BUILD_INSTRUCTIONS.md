# Windows 打包说明

在 Windows 电脑上执行一次即可生成 `ClearanceOS.exe`。

## 前置条件

1. 安装 Python 3.11 或更新版本。
2. 安装时勾选 `Add python.exe to PATH`。

## 打包步骤

1. 解压 `windows_build_source.zip`。
2. 双击 `build_windows.bat`。
3. 等待命令窗口显示 `Build complete`。
4. 生成文件位于：

```text
dist\ClearanceOS\ClearanceOS.exe
```

把整个 `dist\ClearanceOS` 文件夹发给 Windows 员工，不要只发单个 exe，因为 `_internal` 目录里包含运行依赖、模板和配置文件。

## 更新模板

如需更新 mapping 或模板，重新替换源码里的这些文件后再运行 `build_windows.bat`：

- `data\supplier_product_mapping_import.yaml`
- `data\进口商品识别规则.xlsx`
- `data\supplier_product_mapping_export.yaml`
- `export_templates\出口申报要素对应表.xlsx`
- `templates\`
- `export_templates\`
