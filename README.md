# Contagem-de-Estoque

Projeto Python que automatiza a inserção de chassis via Selenium.

Pré-requisitos
- Windows
- Python 3.8+
- Conexão com internet (para webdriver_manager fazer download do ChromeDriver, a menos que use um chromedriver.exe local)

Instalação e execução (modo desenvolvimento)
1. Criar um ambiente virtual e ativar:

```powershell
python -m venv .venv
.\\.venv\\Scripts\\Activate.ps1
```

2. Instalar dependências:

```powershell
pip install -r requirements.txt
```

3. Criar um arquivo `.env` baseado em `.env.example` com suas credenciais.
4. Executar o script:

```powershell
python main.py
```

Gerando um executável (.exe) com PyInstaller
1. Opcional: colocar `chromedriver.exe` na raiz do projeto (já existe no repositório) para que o executável use-o localmente.
2. Com o ambiente ativo e dependências instaladas, rodar:

```powershell
pyinstaller --noconfirm --onefile --add-binary "chromedriver.exe;." --name ContagemDeEstoque main.py
```

- `--onefile` gera um único .exe.
- `--add-binary "chromedriver.exe;."` embute o chromedriver na pasta temporária do executável e faz com que o código detecte e use o chromedriver local.
- Se quiser que o console apareça para ver logs, remova `--windowed`.

Observações importantes
- Ao empacotar com PyInstaller, o caminho `_MEIPASS` é usado para localizar arquivos embutidos — o `main.py` já tenta detectar isso.
- Se preferir não embutir o chromedriver, copie `chromedriver.exe` para a mesma pasta do `.exe` após a build.
- Teste o `.exe` em uma máquina Windows onde o Chrome esteja instalado na versão compatível com o chromedriver.

Solução de problemas
- Erros de compatibilidade do Chrome/ChromeDriver: atualize `chromedriver.exe` para a versão correspondente ao Chrome instalado.
- O Selenium pode precisar de permissões para abrir janelas; execute o .exe com permissões de usuário apropriadas.

Script auxiliar
- Incluí `build_exe.ps1` para automatizar a criação do ambiente e do .exe.

Publicar no GitHub e criar Release
1. Inicialize o repositório (se ainda não):

```powershell
git init
git add .
git commit -m "Initial commit - ContagemDeEstoque v$(Get-Content VERSION -Raw).Trim()"
git branch -M main
git remote add origin <url-do-seu-repo>
git push -u origin main
```

2. Criar uma Release (usando GitHub CLI `gh`):

```powershell
gh release create $(Get-Content VERSION -Raw).Trim() dist\ContagemDeEstoque.exe -t "ContagemDeEstoque $(Get-Content VERSION -Raw).Trim()" -n "Primeira release"
```

Observações sobre o que subir ao GitHub
- Não faça commit de: `dist/`, `build/`, `.venv/`, `*.exe`, `.env` (consulte `.gitignore`).
- Suba código-fonte, scripts de build (`build_exe.ps1`), `requirements.txt`, `README.md` e `VERSION`.

