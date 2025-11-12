#!/bin/bash

# 禁用 NuGet 漏洞扫描以避免离线环境报错
export DOTNET_NUGET_VULNERABILITYAUDITING=0
export NUGET_DISABLE_VULNERABILITY_CHECKS=true

# 清理之前的发布文件
if [ -d "publish" ]; then
    rm -rf publish/*
fi

echo ""
echo "Publishing for Linux x64 (net8.0)..."
DOTNET_DISABLE_VULNERABILITY_CHECKS=1 dotnet publish -c Release -r linux-x64 \
    --framework net8.0 \
    --self-contained true \
    -p:PublishSingleFile=true \
    -p:IncludeNativeLibrariesForSelfExtract=true \
    -p:EnableCompressionInSingleFile=true \
    -p:PublishReadyToRun=false \
    -p:DebuggerSupport=false \
    -p:DebugSymbols=false \
    -p:DebugType=None \
    -p:EnablePackageVulnerabilityAudit=false \
    -p:NuGetAudit=false \
    -p:TreatWarningsAsErrors=false \
    -o publish/linux-x64
