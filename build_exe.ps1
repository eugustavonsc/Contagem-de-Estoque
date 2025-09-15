# Script PowerShell para criar ambiente virtual, instalar dependências e gerar .exe com PyInstaller
# Execute no diretório do projeto (powershell.exe)

python -m venv .venv
.\\.venv\\Scripts\\Activate.ps1
pip install --upgrade pip
pip install -r requirements.txt

# Gera um executável único. Ajuste o nome se quiser
# Se houver um logo.png na raiz, converte para .ico e usa como ícone do exe
$logoPng = Join-Path (Get-Location) 'logo.png'
$logoIco = Join-Path (Get-Location) 'logo.ico'
if (Test-Path $logoPng) {
	Write-Host "Convertendo logo.png para logo.ico..."
	# Converte usando Python/Pillow
	$pyFile = Join-Path (Get-Location) '._convert_logo_temp.py'
	$pyContent = @"
from PIL import Image
img = Image.open(r'{0}')
img.save(r'{1}', format='ICO', sizes=[(256,256)])
"@ -f $logoPng, $logoIco
	Set-Content -Path $pyFile -Value $pyContent -Encoding UTF8
	python $pyFile
	Remove-Item $pyFile -ErrorAction SilentlyContinue
	$iconArg = @('--icon', 'logo.ico')
} else {
	Write-Host "logo.png não encontrado; o executável será gerado sem ícone personalizado." -ForegroundColor Yellow
	$iconArg = ''
}

# Gera o executável com possível ícone personalizado
	$args = @('--noconfirm', '--onefile', '--add-binary', 'chromedriver.exe;.')
if ($iconArg -ne $null -and $iconArg.Length -gt 0) { $args += $iconArg }
$args += @('--name', 'ContagemDeEstoque', 'main.py')
Write-Host "Executando: pyinstaller $($args -join ' ')"
Start-Process -FilePath pyinstaller -ArgumentList $args -NoNewWindow -Wait

Write-Host "Build concluída. Verifique a pasta dist\\ContagemDeEstoque.exe" -ForegroundColor Green
