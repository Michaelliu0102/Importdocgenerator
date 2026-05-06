# macOS Apple Silicon 打包说明

这份源码包也可以在 Apple Silicon Mac 上生成对应的 `ClearanceOS.app`，适用于 `M1 / M2 / M3 / M4` 机型。

如果目标机器是 Intel Mac，请改用 [MAC_INTEL_BUILD_INSTRUCTIONS.md](/Users/michael/customs_doc_generator/MAC_INTEL_BUILD_INSTRUCTIONS.md) 对应流程，或直接分发 `ClearanceOS-mac-intel.zip`。

## 前置条件

1. 安装带 `Tk 8.6+` 的 Python，推荐 `Python 3.11+`。
2. 如果需要 OCR 功能，还需要安装 `ocrmypdf` 和 `tesseract`。

## 打包步骤

1. 打开 Terminal，进入项目目录。
2. 执行：

```bash
chmod +x build_app.sh
./build_app.sh
```

3. 生成文件位于：

```text
dist/ClearanceOS.app
```

4. 打包完成后建议先检查架构：

```bash
chmod +x scripts/inspect_macos_app.sh
./scripts/inspect_macos_app.sh dist/ClearanceOS.app
```

看到 `Architectures` 里包含 `arm64`，就说明这是 Apple Silicon 版本。

如果你手里已经有 GitHub Actions 产物，优先直接使用 `ClearanceOS-mac-arm64.zip`。

## 首次打开

如果 macOS 提示无法打开，右键 `ClearanceOS.app`，选择“打开”，再确认一次。
