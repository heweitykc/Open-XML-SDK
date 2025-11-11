#!/bin/bash

echo "============================================"
echo "Building Self-Contained Executables"
echo "============================================"

# 禁用 NuGet 漏洞扫描以避免离线环境报错
export DOTNET_NUGET_VULNERABILITYAUDITING=0
export NUGET_DISABLE_VULNERABILITY_CHECKS=true

# 清理之前的发布文件
if [ -d "publish" ]; then
    rm -rf publish
fi

发布为 Windows x64 自包含可执行文件
echo ""
echo "Publishing for Windows x64 (net8.0)..."
DOTNET_DISABLE_VULNERABILITY_CHECKS=1 dotnet publish -c Release -r win-x64 \
    --framework net8.0 \
    --self-contained true \
    -p:PublishSingleFile=true \
    -p:IncludeNativeLibrariesForSelfExtract=true \
    -p:EnableCompressionInSingleFile=false \
    -p:PublishReadyToRun=true \
    -p:EnablePackageVulnerabilityAudit=false \
    -p:NuGetAudit=false \
    -o publish/win-x64

if [ $? -ne 0 ]; then
    echo ""
    echo "Windows build failed!"
    exit 1
fi

