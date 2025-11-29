# --- CONFIGURAÇÃO AUTOMÁTICA DE CAMINHO ---
# $PSScriptRoot pega a pasta onde este script está salvo
$calibreConvert = "$PSScriptRoot\ebook-convert.exe"

# --- TÍTULO E CORES ---
Clear-Host
Write-Host "==================================================" -ForegroundColor Cyan
Write-Host "      CONVERSOR PORTATIL DE MOBI PARA EPUB" -ForegroundColor Cyan
Write-Host "==================================================" -ForegroundColor Cyan
Write-Host ""

# --- VERIFICAÇÃO DE SEGURANÇA ---
if (!(Test-Path $calibreConvert)) {
    Write-Host "[ERRO CRÍTICO]" -ForegroundColor Red
    Write-Host "O arquivo 'ebook-convert.exe' não foi encontrado nesta pasta!" -ForegroundColor White
    Write-Host "Certifique-se de que este script está junto com os arquivos do Calibre."
    Read-Host "Pressione Enter para sair..."
    Exit
}

# --- PERGUNTA AO USUÁRIO ---
Write-Host "Cole o caminho da pasta onde estão os arquivos MOBI:" -ForegroundColor Yellow
Write-Host "(Exemplo: B:\Download\Livros\MOBI)"
$pastaOrigem = Read-Host "Caminho"

# Remove aspas extras se o usuário colou com elas
$pastaOrigem = $pastaOrigem -replace '"', ''

if (!(Test-Path $pastaOrigem)) {
    Write-Host "A pasta não foi encontrada!" -ForegroundColor Red
    Read-Host "Pressione Enter para sair..."
    Exit
}

# --- PREPARAÇÃO ---
$pastaDestino = "$pastaOrigem\EPUB_CONVERTIDOS"
$arquivoLog = "$pastaDestino\Relatorio_Erros.txt"

if (!(Test-Path $pastaDestino)) { New-Item -ItemType Directory -Path $pastaDestino | Out-Null }
if (Test-Path $arquivoLog) { Remove-Item $arquivoLog }

$arquivosMobi = @(Get-ChildItem -Path "$pastaOrigem\*.mobi")
$totalArquivos = $arquivosMobi.Count
$contadorGeral = 0
$cntSucesso = 0
$cntJaExiste = 0
$cntErro = 0
$listaDeErros = @()

Write-Host "`nIniciando processamento de $totalArquivos livros..." -ForegroundColor Cyan
Write-Host "--------------------------------------------------"

# --- LOOP PRINCIPAL ---
foreach ($arquivo in $arquivosMobi) {
    $contadorGeral++
    $restantes = $totalArquivos - $contadorGeral
    $nomeBase = $arquivo.BaseName
    $caminhoEpub = "$pastaDestino\$nomeBase.epub"
    $progresso = "[$contadorGeral/$totalArquivos]"

    if (Test-Path $caminhoEpub) {
        Write-Host "$progresso Pulando: $nomeBase (Já existe)" -ForegroundColor DarkGray
        $cntJaExiste++
    }
    else {
        Write-Host "$progresso Convertendo: $nomeBase..." -ForegroundColor Green -NoNewline
        
        # Executa conversão capturando erro
        $resultado = & $calibreConvert "$($arquivo.FullName)" "$caminhoEpub" 2>&1
        
        if ($LASTEXITCODE -eq 0) {
            Write-Host " [OK] -> Faltam: $restantes" -ForegroundColor Yellow
            $cntSucesso++
        } else {
            Write-Host " [FALHA]" -ForegroundColor Red
            $cntErro++
            
            # Formata o erro para o log
            $msgErro = $resultado | Select-Object -Last 5
            $detalhe = "ARQUIVO: $nomeBase`nMOTIVO: $msgErro`n----------------------------------"
            $listaDeErros += $detalhe
        }
    }
}

# --- RELATÓRIO FINAL ---
Write-Host "`n--------------------------------------------------"
Write-Host "RESUMO FINAL" -ForegroundColor Cyan
Write-Host "Total: $totalArquivos | Sucessos: $cntSucesso | Pulados: $cntJaExiste | Erros: $cntErro"

if ($cntErro -gt 0) {
    Write-Host "`n[ATENÇÃO] Houve erros. Gerando log..." -ForegroundColor Red
    $listaDeErros | Out-File -FilePath $arquivoLog -Encoding UTF8
    Invoke-Item $arquivoLog
} else {
    Write-Host "`nSucesso Total!" -ForegroundColor Green
}

Read-Host "Pressione Enter para fechar..."