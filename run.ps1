# 设置输出颜色函数
function Write-ColorOutput {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Message,
        [Parameter(Mandatory = $false)]
        [string]$Color = "White"
    )
    Write-Host $Message -ForegroundColor $Color
}

# 检查 Python 是否安装
Write-ColorOutput "检查 Python 环境..." "Green"
if (-not (Get-Command python -ErrorAction SilentlyContinue)) {
    Write-ColorOutput "未找到 Python，请先安装 Python 3" "Red"
    exit 1
}

# 检查 pip 是否安装
Write-ColorOutput "检查 pip..." "Green"
if (-not (Get-Command pip -ErrorAction SilentlyContinue)) {
    Write-ColorOutput "未找到 pip，请先安装 pip" "Red"
    exit 1
}

# 检查 requirements.txt 是否存在
Write-ColorOutput "检查 requirements.txt..." "Green"
if (-not (Test-Path "requirements.txt")) {
    Write-ColorOutput "未找到 requirements.txt 文件" "Red"
    exit 1
}

# 检查 main.py 是否存在
Write-ColorOutput "检查 main.py..." "Green"
if (-not (Test-Path "main.py")) {
    Write-ColorOutput "未找到 main.py 文件" "Red"
    exit 1
}

# 检查 data 目录是否存在，如果不存在则创建
Write-ColorOutput "检查 data 目录..." "Green"
if (-not (Test-Path "data")) {
    Write-ColorOutput "创建 data 目录..." "Green"
    New-Item -ItemType Directory -Path "data"
}

# 检查 data 目录中是否有 Excel 文件
Write-ColorOutput "检查 Excel 文件..." "Green"
$excelFiles = Get-ChildItem -Path "data" -Filter *.xls* -File
if ($excelFiles.Count -eq 0) {
    Write-ColorOutput "data 目录下没有找到 Excel 文件" "Red"
    Write-ColorOutput "请将 Excel 文件放入 data 目录中" "Green"
    exit 1
}

# 安装依赖
Write-ColorOutput "安装依赖包..." "Green"
$result = pip install -r requirements.txt
if ($LASTEXITCODE -eq 0) {
    Write-ColorOutput "开始运行数据分析脚本..." "Green"
    python main.py
}
else {
    Write-ColorOutput "依赖包安装失败" "Red"
    exit 1
} 