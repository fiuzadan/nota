# Assistente NotaJá — configuração e publicação

## 1. Arquivos no GitHub

Coloque `NotaJa-Suporte.ps1` na raiz de um repositório público.

Edite estas variáveis no começo do arquivo:

```powershell
$RemoteScriptUrl = 'https://raw.githubusercontent.com/USUARIO/REPOSITORIO/main/NotaJa-Suporte.ps1'
$InstallerUrl = 'https://github.com/USUARIO/REPOSITORIO/releases/download/v5.10c/NotaJa.exe'
$ExpectedInstallerSha256 = 'HASH_SHA256_DO_INSTALADOR'
```

O script pode ficar normalmente no repositório. Para o instalador EXE, prefira
um ativo de **GitHub Release**, principalmente quando o arquivo for grande.

## 2. Gerar o SHA-256 do instalador

No computador onde está o instalador:

```powershell
Get-FileHash .\NotaJa.exe -Algorithm SHA256
```

Copie somente o valor da coluna `Hash` para `$ExpectedInstallerSha256`.

## 3. Executar pelo PowerShell

Depois de publicar e configurar o endereço RAW:

```powershell
irm 'https://raw.githubusercontent.com/USUARIO/REPOSITORIO/main/NotaJa-Suporte.ps1' | iex
```

O próprio script solicitará elevação do UAC e abrirá novamente como
administrador.

Também é possível baixar e executar localmente:

```powershell
iwr 'https://raw.githubusercontent.com/USUARIO/REPOSITORIO/main/NotaJa-Suporte.ps1' -OutFile "$env:TEMP\NotaJa-Suporte.ps1"
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "$env:TEMP\NotaJa-Suporte.ps1"
```

## 4. Menu disponível

- **Nova instalação**
  - Banco de dados em nuvem: `mysqlnuvem`
  - Estação de trabalho local: `semmysql`
  - Banco de dados local/MySQL 5.6: `instalarmysql`

- **Reinstalação**
  1. Encerra processos conhecidos do NotaJá.
  2. Executa `unins000.exe` silenciosamente.
  3. Faz backup e remove somente a lista configurada de System32/SysWOW64.
  4. Baixa novamente o instalador.
  5. Pergunta o tipo e reinstala silenciosamente.

## 5. Pastas criadas

```text
%TEMP%\NotaJa-Suporte
%ProgramData%\NotaJa-Suporte\Logs
%ProgramData%\NotaJa-Suporte\BackupComponentes
```

O script não remove `C:\NFE` nem apaga intencionalmente os dados do MySQL.

## 6. Teste recomendado

Antes de usar em cliente:

1. Publique o script e o instalador.
2. Configure o SHA-256.
3. Teste cada opção em uma máquina virtual.
4. Confira os logs e o backup dos componentes removidos.
5. Confirme que o banco local existente continua acessível após a reinstalação.
