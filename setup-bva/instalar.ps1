<#
  instalar.ps1 - Setup de Claude para el Estudio BVA (Windows 10/11).

  Deja la computadora lista para usar los plugins del marketplace
  estudiobva-skills tal cual se desarrollaron:

    1. Programas: Git, Node.js LTS, Python 3.12 y GitHub CLI (con winget),
       Claude Code (instalador oficial) y, si falta, la app de Claude.
    2. GitHub: login con el navegador y acceso al repo privado del estudio.
    3. Marketplace: lo agrega (o actualiza) e instala TODOS sus plugins.
    4. Dependencias de cada plugin: npm, Chromium de Playwright, openpyxl.
    5. Configuracion: ~/.fisco-ar/.env con la unidad compartida, la planilla de
       claves (la contrasena se escribe aca, oculta) y el Python.
    6. Diagnostico final.

  Se puede correr las veces que haga falta: lo que ya esta hecho lo saltea.
  Uso: doble clic en "Instalar Claude BVA.cmd", o
       powershell -ExecutionPolicy Bypass -File instalar.ps1 [-SinProgramas] [-Unidad "G:\Unidades compartidas\BVA - Sociedades Farmaceuticas"]

  Este archivo va sin acentos a proposito: Windows PowerShell 5.1 lee los .ps1
  sin BOM como ANSI.
#>
param(
  [switch]$SinProgramas,     # no instala programas con winget (solo configura)
  [string]$Unidad = ''       # ruta de la unidad compartida, si no se detecta sola
)

$ErrorActionPreference = 'Continue'   # con 'Stop', PS 5.1 corta ante cualquier stderr de un programa
$REPO   = 'estudiobva/estudio-claude-marketplace'
$MARKET = 'estudiobva-skills'
$ClaudeDir = Join-Path $env:USERPROFILE '.claude'
$env:CLAUDE_CODE_PLUGIN_PREFER_HTTPS = '1'   # bajar el repo por HTTPS con las credenciales de gh

try { [Console]::OutputEncoding = [Text.Encoding]::UTF8 } catch {}
$log = Join-Path $env:USERPROFILE 'Documents\BVA-salidas\setup-bva.log'
New-Item -ItemType Directory -Force -Path (Split-Path $log) | Out-Null
try { Start-Transcript -Path $log -Append | Out-Null } catch {}

function Titulo($t) { Write-Host ''; Write-Host "== $t" -ForegroundColor Cyan }
function Ok($t)     { Write-Host "  OK    $t" -ForegroundColor Green }
function Aviso($t)  { Write-Host "  AVISO $t" -ForegroundColor Yellow }
function Info($t)   { Write-Host "        $t" }
function Cortar($t) {
  Write-Host "  ERROR $t" -ForegroundColor Red
  Write-Host ''
  Write-Host 'El setup se detuvo. Cuando lo resuelvas, volve a correrlo: retoma donde quedo.'
  Write-Host "Registro completo: $log"
  try { Stop-Transcript | Out-Null } catch {}
  exit 1
}

function Refrescar-Path {
  $m = [Environment]::GetEnvironmentVariable('Path', 'Machine')
  $u = [Environment]::GetEnvironmentVariable('Path', 'User')
  $env:Path = "$m;$u;$env:USERPROFILE\.local\bin"
}

function Tiene($cmd) { return [bool](Get-Command $cmd -ErrorAction SilentlyContinue) }

function Instalar-Winget($id, $nombre, $override) {
  if ($SinProgramas) { Cortar "Falta $nombre y se pidio -SinProgramas." }
  if (-not (Tiene 'winget')) {
    Cortar "Falta $nombre y no esta winget. Instala 'Instalador de aplicacion' (App Installer) desde Microsoft Store y volve a correr esto."
  }
  Info "Instalando $nombre... (si Windows pregunta si permitis cambios, acepta)"
  $wargs = @('install', '--id', $id, '-e', '--accept-package-agreements', '--accept-source-agreements')
  if ($override) { $wargs += @('--override', $override) } else { $wargs += '--silent' }
  & winget @wargs
  Refrescar-Path
}

# Python real: "python.exe" de WindowsApps es un atajo a la Store que no corre nada.
function Buscar-Python {
  $prueba = 'import sys;print(sys.executable);print(sys.version_info[0]*100+sys.version_info[1])'
  $candidatos = @()
  if (Tiene 'py') { $candidatos += ,@('py', '-3') }
  if (Tiene 'python') { $candidatos += ,@('python') }
  foreach ($c in $candidatos) {
    if ($c.Count -gt 1) { $out = & $c[0] $c[1] -c $prueba 2>$null } else { $out = & $c[0] -c $prueba 2>$null }
    if ($LASTEXITCODE -eq 0 -and $out -and $out.Count -ge 2) {
      $exe = "$($out[0])".Trim()
      if ($exe -notmatch 'WindowsApps' -and [int]$out[1] -ge 309) { return $exe }
    }
  }
  return $null
}

function Leer-Json($archivo) {
  if (-not (Test-Path $archivo)) { return $null }
  try { return (Get-Content -Raw -Encoding UTF8 $archivo | ConvertFrom-Json) } catch { return $null }
}

function Plugins-Instalados {
  $j = Leer-Json (Join-Path $ClaudeDir 'plugins\installed_plugins.json')
  $h = @{}
  if ($j -and $j.plugins) {
    foreach ($prop in $j.plugins.PSObject.Properties) {
      $partes = $prop.Name.Split('@')
      if ($partes.Count -eq 2 -and $partes[1] -eq $MARKET) {
        $e = @($prop.Value)[0]
        if ($e -and $e.installPath) { $h[$partes[0]] = $e.installPath }
      }
    }
  }
  return $h
}

function Leer-Clave($pregunta) {
  $s = Read-Host $pregunta -AsSecureString
  $b = [Runtime.InteropServices.Marshal]::SecureStringToBSTR($s)
  try { return [Runtime.InteropServices.Marshal]::PtrToStringBSTR($b) }
  finally { [Runtime.InteropServices.Marshal]::ZeroFreeBSTR($b) }
}

Write-Host ''
Write-Host '  Setup de Claude - Estudio BVA' -ForegroundColor Cyan
Write-Host '  Instala lo necesario para usar los plugins del estudio. Tarda unos minutos.'
Refrescar-Path

# -- 1. Programas -------------------------------------------------------------
Titulo '1/6  Programas'

if (-not (Tiene 'git')) { Instalar-Winget 'Git.Git' 'Git' $null }
if (Tiene 'git') { Ok "Git $((& git --version) -replace 'git version ','')" } else { Cortar 'No se pudo instalar Git. Cerra esta ventana, abri una nueva y volve a intentar.' }

$nodeOk = $false
if (Tiene 'node') { $nodeOk = ([int](((& node -v) -replace '^v','').Split('.')[0]) -ge 20) }
if (-not $nodeOk) {
  if (Tiene 'node') { Info 'Node es anterior a la version 20: se actualiza.' }
  Instalar-Winget 'OpenJS.NodeJS.LTS' 'Node.js LTS' $null
  if (Tiene 'node') { $nodeOk = ([int](((& node -v) -replace '^v','').Split('.')[0]) -ge 20) }
}
if ($nodeOk) { Ok "Node $(& node -v)" } else { Cortar 'No se pudo instalar Node 20 o superior. Cerra esta ventana, abri una nueva y volve a intentar.' }

$python = Buscar-Python
if (-not $python) {
  Instalar-Winget 'Python.Python.3.12' 'Python 3.12' '/quiet InstallAllUsers=0 PrependPath=1 Include_launcher=1'
  $python = Buscar-Python
}
if ($python) { Ok "Python $python" } else { Cortar 'No se pudo instalar Python. Cerra esta ventana, abri una nueva y volve a intentar.' }

if (-not (Tiene 'gh')) { Instalar-Winget 'GitHub.cli' 'GitHub CLI' $null }
if (Tiene 'gh') { Ok 'GitHub CLI' } else { Cortar 'No se pudo instalar GitHub CLI.' }

if (-not (Tiene 'claude')) {
  if ($SinProgramas) { Cortar 'Falta Claude Code y se pidio -SinProgramas.' }
  Info 'Instalando Claude Code (instalador oficial de claude.ai)...'
  & powershell -NoProfile -ExecutionPolicy Bypass -Command 'irm https://claude.ai/install.ps1 | iex'
  Refrescar-Path
}
if (Tiene 'claude') { Ok "Claude Code $((& claude --version) -join ' ')" } else { Cortar 'No se pudo instalar Claude Code. Ver https://code.claude.com/docs/en/setup' }

# La app de escritorio: es donde el equipo usa las skills.
$app = @("$env:LOCALAPPDATA\AnthropicClaude", "$env:LOCALAPPDATA\Programs\Claude", "$env:ProgramFiles\Claude") | Where-Object { Test-Path $_ }
if (-not $app -and (Tiene 'winget')) {
  & winget list --id Anthropic.Claude -e *> $null
  if ($LASTEXITCODE -eq 0) { $app = @('winget') }
}
if ($app) { Ok 'App de Claude' }
elseif (-not $SinProgramas -and (Tiene 'winget')) {
  $r = Read-Host '  No encuentro la app de Claude. Instalarla ahora? (S/n)'
  if ($r -notmatch '^[nN]') { Instalar-Winget 'Anthropic.Claude' 'la app de Claude' $null }
  else { Aviso 'Bajala de https://claude.ai/download y entra con tu cuenta del estudio.' }
} else { Aviso 'No encuentro la app de Claude: bajala de https://claude.ai/download' }

# -- 2. GitHub ----------------------------------------------------------------
Titulo '2/6  GitHub (repo privado del estudio)'

& gh auth status --hostname github.com *> $null
if ($LASTEXITCODE -ne 0) {
  Info 'Vas a entrar a GitHub con el navegador:'
  Info '  - si pregunta, elegi HTTPS y "Y" para autenticar Git;'
  Info '  - copia el codigo de 8 letras que aparece aca, pegalo en la pagina y toca Authorize.'
  & gh auth login --hostname github.com --git-protocol https --web
  & gh auth status --hostname github.com *> $null
  if ($LASTEXITCODE -ne 0) { Cortar 'No se completo el login de GitHub.' }
}
& gh auth setup-git --hostname github.com *> $null
$usuario = (& gh api user --jq .login 2>$null)
Ok "Logueado en GitHub como $usuario"

$env:GIT_TERMINAL_PROMPT = '0'
& git ls-remote "https://github.com/$REPO" HEAD *> $null
$acceso = ($LASTEXITCODE -eq 0)
Remove-Item Env:GIT_TERMINAL_PROMPT -ErrorAction SilentlyContinue
if (-not $acceso) {
  Cortar ("Tu usuario de GitHub ($usuario) no tiene acceso a $REPO. Pedile a Joaquin que te invite a la organizacion " +
          "estudiobva con ese usuario, acepta la invitacion (llega por mail o en github.com/orgs/estudiobva/invitation) y volve a correr esto.")
}
Ok "Acceso a $REPO"

# -- 3. Marketplace y plugins -------------------------------------------------
Titulo '3/6  Marketplace y plugins'

$manifest = Join-Path $ClaudeDir "plugins\marketplaces\$MARKET\.claude-plugin\marketplace.json"
if (Test-Path $manifest) { & claude plugin marketplace update $MARKET }
else { & claude plugin marketplace add $REPO }
$m = Leer-Json $manifest
if (-not $m) { Cortar "No se pudo agregar el marketplace $MARKET." }
Ok "Marketplace $MARKET"

$inst = Plugins-Instalados
foreach ($pl in $m.plugins) {
  $id = "$($pl.name)@$MARKET"
  if ($inst.ContainsKey($pl.name)) { & claude plugin update $id *> $null; Ok "$($pl.name) (ya estaba; actualizado si habia version nueva)" }
  else {
    & claude plugin install $id --scope user
    if ($LASTEXITCODE -eq 0) { Ok "$($pl.name) instalado" } else { Aviso "No se pudo instalar $($pl.name)" }
  }
}
$inst = Plugins-Instalados
if (-not $inst.ContainsKey('setup-bva')) { Cortar 'No quedo instalado el plugin setup-bva (lo necesita el resto del setup).' }
$setup = $inst['setup-bva']

# -- 4. Dependencias ----------------------------------------------------------
Titulo '4/6  Dependencias de los plugins (npm, Chromium, Python)'
$env:BVA_PYTHON = $python
& node (Join-Path $setup 'scripts\dependencias.js')
if ($LASTEXITCODE -ne 0) { Aviso 'Quedaron dependencias con error (ver arriba). El diagnostico lo va a marcar.' }

# -- 5. Configuracion ---------------------------------------------------------
Titulo '5/6  Configuracion (.env, unidad compartida, planilla de claves)'

function Configurar {
  $a = @((Join-Path $setup 'scripts\configurar.js'), '--json', "--python=$python")
  if ($Unidad) { $a += "--unidad=$Unidad" }
  $txt = (& node @a) -join "`n"
  try { return ($txt | ConvertFrom-Json) } catch { Cortar "configurar.js no respondio bien: $txt" }
}

$cfg = Configurar
if (-not $cfg.unidad) {
  Cortar ('No encuentro la unidad compartida "BVA - Sociedades Farmaceuticas". Abri Google Drive para escritorio, entra con tu ' +
          'cuenta @estudiobva.com, espera a que aparezca en el Explorador y volve a correr esto (o pasa -Unidad "<ruta>").')
}
Ok "Unidad compartida: $($cfg.unidad)"

$intentos = 0
while (-not $cfg.claves.ok -and $cfg.claves.motivo -eq 'clave' -and $intentos -lt 3) {
  if ($intentos -gt 0) { Aviso 'La contrasena no abre la planilla. Proba de nuevo.' }
  $env:BVA_SETUP_CLAVE = Leer-Clave '  Contrasena de Claves_Organismos.xlsx (no se ve al escribir)'
  $cfg = Configurar
  Remove-Item Env:BVA_SETUP_CLAVE -ErrorAction SilentlyContinue
  $intentos++
}
if ($cfg.claves.ok) {
  Ok "Claves_Organismos.xlsx abre ($($cfg.claves.titulares) titulares)"
  if ($cfg.claves.cifrado -eq $false) { Aviso 'Claves_Organismos.xlsx no tiene contrasena de apertura: avisale a Joaquin.' }
}
elseif ($cfg.claves.motivo -eq 'clave') { Cortar 'La contrasena de Claves_Organismos.xlsx no es correcta. Pedisela a Joaquin.' }
else { Cortar "Claves_Organismos.xlsx: $($cfg.claves.error)" }

if ($cfg.clavesBancos -and -not $cfg.clavesBancos.ok -and $cfg.clavesBancos.motivo -eq 'clave') {
  $env:BVA_SETUP_CLAVE_BANCOS = Leer-Clave '  Contrasena de Claves_Bancos.xlsx (no se ve al escribir; Enter para saltear si no usas bancos-bva)'
  if ($env:BVA_SETUP_CLAVE_BANCOS) { $cfg = Configurar }
  Remove-Item Env:BVA_SETUP_CLAVE_BANCOS -ErrorAction SilentlyContinue
}
if ($cfg.clavesBancos) {
  if ($cfg.clavesBancos.ok) { Ok 'Claves_Bancos.xlsx abre' } else { Aviso "Claves_Bancos.xlsx: $($cfg.clavesBancos.error) (solo afecta a bancos-bva)" }
}
Ok "Configuracion guardada en $($cfg.envFile)"

# Tambien como variables de usuario: los scripts de Python que Claude corre
# directo (sin pasar por Node) leen BVA_UNIDAD_PATH del entorno, no del .env.
[Environment]::SetEnvironmentVariable('BVA_UNIDAD_PATH', $cfg.unidad, 'User')
[Environment]::SetEnvironmentVariable('BVA_PYTHON', $cfg.python, 'User')
Ok 'Variables de usuario BVA_UNIDAD_PATH y BVA_PYTHON'

# -- 6. Diagnostico -----------------------------------------------------------
Titulo '6/6  Diagnostico'
& node (Join-Path $setup 'scripts\diagnostico.js')
$diag = $LASTEXITCODE

Write-Host ''
if ($diag -eq 0) {
  Write-Host '  Listo. Cerra la app de Claude del todo (tambien desde el icono junto al reloj) y volvela a abrir:' -ForegroundColor Green
  Write-Host '  los plugins del estudio aparecen en la pestana Code. Si algo falla mas adelante, pedile a Claude'
  Write-Host '  "revisa el setup del estudio" o volve a correr este instalador.'
} else {
  Write-Host '  El setup termino con errores (ver el diagnostico). Mandale a Joaquin este archivo:' -ForegroundColor Yellow
  Write-Host "  $log"
}
try { Stop-Transcript | Out-Null } catch {}
exit $diag
