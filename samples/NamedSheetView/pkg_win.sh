#!/bin/bash

echo "============================================"
echo "Building Windows Executable in WSL"
echo "============================================"

# 清理之前的发布文件
if [ -d "publish" ]; then
    rm -rf publish
fi

# 发布为 Windows x64 自包含可执行文件
echo ""
echo "Publishing for Windows x64 (net8.0)..."
dotnet publish -c Release -r win-x64 \
    --framework net8.0 \
    --self-contained true \
    -p:PublishSingleFile=true \
    -p:IncludeNativeLibrariesForSelfExtract=true \
    -p:EnableCompressionInSingleFile=true \
    -o publish/win-x64

if [ $? -ne 0 ]; then
    echo ""
    echo "Build failed!"
    exit 1
fi

echo ""
echo "============================================"
echo "Build completed successfully!"
echo "Output directory: publish/win-x64"
echo "Executable: publish/win-x64/AddNamedSheetView.exe"
echo "============================================"
echo ""

# 列出发布的文件
ls -lh publish/win-x64/*.exe

echo ""
echo "You can now run this exe in Windows:"
echo "  publish/win-x64/AddNamedSheetView.exe"
echo ""

