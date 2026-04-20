# ============================================================
# 自动环境配置脚本 / Automatic Environment Setup Script
# ============================================================
# 用途: 自动检测并安装项目所需的所有依赖
# Usage: Automatically detect and install all project dependencies
# 
# 运行方式 / How to run:
#   .\setup_env.ps1
#   或 / or
#   .\setup_env.ps1 -FullInstall
# ============================================================

param(
    [switch]$FullInstall = $false,  # 安装所有可选依赖
    [switch]$SkipOCR = $false,      # 跳过 OCR 依赖
    [switch]$SkipAI = $false        # 跳过 AI 依赖
)

Write-Host "============================================================" -ForegroundColor Cyan
Write-Host "  文档处理环境自动配置脚本" -ForegroundColor Cyan
Write-Host "  Document Processing Environment Setup" -ForegroundColor Cyan
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host ""

# 检测 Python 版本
Write-Host "[1/6] 检测 Python 环境..." -ForegroundColor Yellow
$pythonVersion = python --version 2>&1
if ($LASTEXITCODE -eq 0) {
    Write-Host "  ✓ Python 已安装: $pythonVersion" -ForegroundColor Green
} else {
    Write-Host "  ✗ 错误: 未检测到 Python" -ForegroundColor Red
    Write-Host "  请先安装 Python 3.8+ : https://www.python.org/downloads/" -ForegroundColor Red
    exit 1
}

# 检测虚拟环境
Write-Host ""
Write-Host "[2/6] 检测虚拟环境..." -ForegroundColor Yellow
$condaEnv = conda env list 2>&1 | Select-String "office"
if ($condaEnv) {
    Write-Host "  ✓ 检测到 'office' 虚拟环境" -ForegroundColor Green
    Write-Host "  提示: 请确保已激活环境 (conda activate office)" -ForegroundColor Cyan
} else {
    Write-Host "  ⚠ 未检测到 'office' 虚拟环境" -ForegroundColor Yellow
    Write-Host "  将在当前环境中安装依赖" -ForegroundColor Yellow
}

# 安装核心依赖
Write-Host ""
Write-Host "[3/6] 安装核心依赖..." -ForegroundColor Yellow

# 检查并安装 document-builder
Write-Host "  检查 document-builder..." -NoNewline
python -c "import docbuilder" 2>$null
if ($LASTEXITCODE -eq 0) {
    Write-Host " ✓ 已安装" -ForegroundColor Green
} else {
    Write-Host " ✗ 未安装" -ForegroundColor Red
    Write-Host "  正在安装 document-builder..." -ForegroundColor Cyan
    pip install document-builder
    if ($LASTEXITCODE -eq 0) {
        Write-Host "  ✓ document-builder 安装成功" -ForegroundColor Green
    } else {
        Write-Host "  ⚠ document-builder 安装失败" -ForegroundColor Yellow
        Write-Host "  请手动下载 ONLYOFFICE Document Builder SDK" -ForegroundColor Yellow
        Write-Host "  下载地址: https://api.onlyoffice.com/docbuilder/installation" -ForegroundColor Cyan
    }
}

# 检查并安装 lxml
Write-Host "  检查 lxml..." -NoNewline
python -c "import lxml" 2>$null
if ($LASTEXITCODE -eq 0) {
    Write-Host " ✓ 已安装" -ForegroundColor Green
} else {
    Write-Host " ✗ 未安装" -ForegroundColor Red
    Write-Host "  正在安装 lxml..." -ForegroundColor Cyan
    pip install lxml
}

# 检查并安装 Pillow
Write-Host "  检查 Pillow..." -NoNewline
python -c "import PIL" 2>$null
if ($LASTEXITCODE -eq 0) {
    Write-Host " ✓ 已安装" -ForegroundColor Green
} else {
    Write-Host " ✗ 未安装" -ForegroundColor Red
    Write-Host "  正在安装 Pillow..." -ForegroundColor Cyan
    pip install Pillow
}

# 安装可选依赖 - OCR
if (-not $SkipOCR -and $FullInstall) {
    Write-Host ""
    Write-Host "[4/6] 安装 OCR 依赖 (可选)..." -ForegroundColor Yellow
    
    Write-Host "  检查 pix2tex..." -NoNewline
    python -c "import pix2tex" 2>$null
    if ($LASTEXITCODE -eq 0) {
        Write-Host " ✓ 已安装" -ForegroundColor Green
    } else {
        Write-Host " ✗ 未安装" -ForegroundColor Red
        Write-Host "  正在安装 pix2tex..." -ForegroundColor Cyan
        pip install pix2tex
    }
    
    Write-Host "  检查 latex2mathml..." -NoNewline
    python -c "import latex2mathml" 2>$null
    if ($LASTEXITCODE -eq 0) {
        Write-Host " ✓ 已安装" -ForegroundColor Green
    } else {
        Write-Host " ✗ 未安装" -ForegroundColor Red
        Write-Host "  正在安装 latex2mathml..." -ForegroundColor Cyan
        pip install latex2mathml
    }
    
    # 检查 ImageMagick
    Write-Host "  检查 ImageMagick..." -NoNewline
    $imageMagick = Get-Command convert -ErrorAction SilentlyContinue
    if ($imageMagick) {
        Write-Host " ✓ 已安装" -ForegroundColor Green
    } else {
        Write-Host " ✗ 未安装" -ForegroundColor Red
        Write-Host "  ⚠ ImageMagick 未安装 (OCR 模式需要)" -ForegroundColor Yellow
        Write-Host "  下载地址: https://imagemagick.org/script/download.php" -ForegroundColor Cyan
    }
} else {
    Write-Host ""
    Write-Host "[4/6] 跳过 OCR 依赖安装 (使用 -FullInstall 启用)" -ForegroundColor Gray
}

# 安装可选依赖 - AI 文本校对
if (-not $SkipAI -and $FullInstall) {
    Write-Host ""
    Write-Host "[5/6] 安装 AI 文本校对依赖 (可选)..." -ForegroundColor Yellow
    
    Write-Host "  检查 pycorrector..." -NoNewline
    python -c "import pycorrector" 2>$null
    if ($LASTEXITCODE -eq 0) {
        Write-Host " ✓ 已安装" -ForegroundColor Green
    } else {
        Write-Host " ✗ 未安装" -ForegroundColor Red
        Write-Host "  正在安装 pycorrector (本地校对)..." -ForegroundColor Cyan
        pip install pycorrector
    }
} else {
    Write-Host ""
    Write-Host "[5/6] 跳过 AI 依赖安装 (使用 -FullInstall 启用)" -ForegroundColor Gray
}

# 验证安装
Write-Host ""
Write-Host "[6/6] 验证安装..." -ForegroundColor Yellow

if (Test-Path "verify_installation.py") {
    Write-Host "  运行验证脚本..." -ForegroundColor Cyan
    python verify_installation.py
} else {
    Write-Host "  ⚠ 未找到 verify_installation.py" -ForegroundColor Yellow
}

# 总结
Write-Host ""
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host "  环境配置完成!" -ForegroundColor Green
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "下一步:" -ForegroundColor Yellow
Write-Host "  1. 确保已激活 'office' 虚拟环境: conda activate office" -ForegroundColor White
Write-Host "  2. 运行文档合并: python run_merge.py --help" -ForegroundColor White
Write-Host "  3. 运行公式转换: python convert_docx_formulas.py --help" -ForegroundColor White
Write-Host ""
Write-Host "文档:" -ForegroundColor Yellow
Write-Host "  - setup_environment.md (环境配置指南)" -ForegroundColor White
Write-Host "  - QUICK_START.md (快速开始)" -ForegroundColor White
Write-Host "  - INSTALLATION_GUIDE.md (详细安装说明)" -ForegroundColor White
Write-Host ""
