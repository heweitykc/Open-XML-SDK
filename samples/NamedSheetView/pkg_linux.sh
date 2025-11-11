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

echo ""
echo "Publishing for Linux x64 (net8.0)..."
# DOTNET_DISABLE_VULNERABILITY_CHECKS=1 dotnet publish -c Release -r linux-x64 \
#     --framework net8.0 \
#     --self-contained true \
#     -p:PublishSingleFile=true \
#     -p:IncludeNativeLibrariesForSelfExtract=true \
#     -p:EnableCompressionInSingleFile=false \
#     -p:PublishReadyToRun=true \
#     -p:EnablePackageVulnerabilityAudit=false \
#     -p:NuGetAudit=false \
#     -o publish/linux-x64

DOTNET_DISABLE_VULNERABILITY_CHECKS=1 dotnet publish -c Release -r linux-x64 \
  --framework net8.0 --self-contained false \
  -p:UseAppHost=true -p:PublishReadyToRun=true \
  -p:PublishSingleFile=false -p:ExecutableName=ppt_gen \
  -o publish/linux-fdd


# 将 Linux 可执行文件复制到 /usr/bin
echo "Copying executable to /usr/bin (may require sudo)..."
sudo cp ./publish/linux-x64/AddNamedSheetView /usr/bin/ppt_gen


if [ $? -ne 0 ]; then
    echo ""
    echo "Linux build failed!"
    exit 1
fi

