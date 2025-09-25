# README — Processar Notas por Páginas (Apps Script)

## Visão geral

Este projeto lê arquivos (JPG / PNG / PDF) colocados em uma **pasta do Google Drive**, executa OCR para extrair texto, tenta localizar **data** e **valor total**, registra esses dados em uma **planilha Google (Sheet1)** e **move** os arquivos processados para uma subpasta chamada `Processados`.

**Arquivos principais:**

* `processarNotasPorPaginas()` → função principal que percorre a pasta, faz OCR e grava na planilha.
* Funções auxiliares: `processarBlobOCR`, `esperarOCR`, `extrairValoresDaLinha`, `normalizarValor`, `padronizarData`, `isNomeGenerico`.

---

## Antes de rodar — pré-requisitos

1. Conta Google com acesso ao Drive e ao Apps Script. (Se for conta Workspace, verifique políticas do administrador.)
2. Uma **pasta no Google Drive** com os arquivos a processar.
3. Uma **planilha Google** criada; crie uma aba chamada `Sheet1` (ou altere `ABA_PLANILHA` para o nome da aba que preferir).
4. Copie o **ID da pasta** e o **ID da planilha** e preencha as variáveis no topo do script:

```javascript
var PASTA_ID = "COLE_AQUI_ID_DA_PASTA";
var PLANILHA_ID = "COLE_AQUI_ID_PLANILHA";
var ABA_PLANILHA = "Sheet1";
```

---

## Permissões e configuração necessárias (passo a passo)

> **Importante:** o fluxo de autorização envolve configuração do projeto Google Cloud (GCP) conectado ao Apps Script. Siga na ordem.

### 1) Vincular projeto Apps Script ao projeto GCP (Project ID/Number)

* No editor do Apps Script: clique em **Projeto → Configurações do projeto** (ícone de engrenagem).
* Anote o **Project number / Project ID** mostrado.
* No Google Cloud Console ([https://console.cloud.google.com/](https://console.cloud.google.com/)), selecione exatamente esse projeto ao fazer as etapas abaixo.

### 2) Ativar APIs no Cloud Console (no projeto correto)

No Cloud Console → **APIs e serviços → Biblioteca** ative:

* **Google Drive API**
* **Google Sheets API**
* **Google Docs API** (usado para documentos temporários / OCR)

> Se você optar por usar Cloud Vision (alternativa), ative também **Cloud Vision API** e observe que a Vision exige faturamento ativo.

### 3) Ativar o Advanced Drive Service no Apps Script

No editor do Apps Script:

* Menu antigo: **Recursos → Serviços avançados do Google** → ative **Drive API**.
* Editor novo: clique em **Services** (ícone de + / Add service) e adicione **Drive API**.

Após ativar, recarregue o editor.

### 4) Configurar Tela de Consentimento OAuth (OAuth Consent Screen)

No Cloud Console (projeto correto): **APIs e serviços → Tela de consentimento OAuth**

1. Escolha **Tipo de usuário: Externo** (para contas pessoais).
2. Preencha:

   * **Nome do app**: ex: `OCR-Notas - SeuNome`
   * **E-mail de suporte**: seu e-mail
3. Em **Escopos**, não adicione manualmente aqui — os escopos virão do seu manifest, mas garanta que o app esteja configurado.
4. Em **Usuários de teste**, adicione o e-mail que vai autorizar (o seu).
5. Salve.

> Observação: enquanto o app estiver *em modo de teste*, apenas usuários listados poderão autorizar.

### 5) Manifest (appsscript.json) — manter escopos mínimos

No editor do Apps Script: **Ver → Manifest (appsscript.json)**. Use um manifest mínimo com os escopos necessários. Exemplo sugerido:

```json
{
  "timeZone": "America/Fortaleza",
  "oauthScopes": [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
    "https://www.googleapis.com/auth/documents",
    "https://www.googleapis.com/auth/script.external_request"
  ],
  "exceptionLogging": "STACKDRIVER"
}
```

> Mantenha apenas os escopos que o seu código realmente precisa. Escopos adicionais (sensíveis) exigem verificação adicional pela Google.

### 6) Revogar autorizações antigas e testar (passo final)

1. Em [https://myaccount.google.com/permissions](https://myaccount.google.com/permissions) remova autorizações antigas do app/script (se houver).
2. Abra uma janela **incógnita** no navegador, faça login com a conta do projeto.
3. No Apps Script, execute a função principal (`processarNotasPorPaginas`) ou uma função de teste (ex.: `testeDriveSheets` abaixo). Aceite a tela de autorização quando aparecer.

**Função de teste (cole e rode para validar autorizações):**

```javascript
function testeDriveSheets(){
  Logger.log(Session.getActiveUser().getEmail());
  var p = DriveApp.getRootFolder();
  Logger.log('Root: ' + p.getId());
  var ss = SpreadsheetApp.create('Teste-Auth-' + new Date().getTime());
  Logger.log('Planilha criada: ' + ss.getId());
}
```

---

## Como usar (resumo rápido)

1. Coloque arquivos (JPG/PNG/PDF) na pasta do Drive cujo ID você configurou.
2. No Apps Script, atualize as variáveis `PASTA_ID`, `PLANILHA_ID`, `ABA_PLANILHA`.
3. Execute `processarNotasPorPaginas()` pela primeira vez e aceite as permissões.
4. Verifique a planilha `Sheet1` — cada arquivo processado adiciona uma linha com data, valor e link para o arquivo.
5. Os arquivos processados são movidos para a subpasta `Processados` dentro da pasta original.

---

## Gatilhos (automação)

* Para rodar automaticamente: **Triggers (Relógio)** → Adicionar trigger → função `processarNotasPorPaginas` → Time-driven → defina a frequência desejada.

---

## Erros comuns & soluções

* **`Drive.Files.insert is not a function`**: verifique se o **Advanced Drive Service** está ativado no Apps Script e se o projeto GCP correto está selecionado. Depois recarregue o editor. Se persistir, teste criar um novo projeto Apps Script e migrar o código.

* **`Invalid image data`**: significa que o blob enviado não era um tipo aceito ou o método usado não consegue processar o formato. Verifique se o arquivo é realmente `image/*` ou `application/pdf`. Arquivos que já são Google Docs não devem ser enviados para OCR.

* **`getFolderById`**** com erro**: confirme que o **PASTA_ID** está correto (parte da URL) e que a conta que executa o script tem acesso à pasta (não apenas leitura via link). Teste abrir a pasta com a mesma conta.

* **Erro 500 / "Ocorreu um erro desconhecido" na autorização**: siga a ordem do README (vincular projeto GCP, configurar Tela de Consentimento, reduzir escopos, revogar permissões e tentar em janela incógnita). Consulte os logs do Cloud Console para detalhes.

* **403 ao usar Cloud Vision**: se optar por usar Cloud Vision, certifique-se de **ativar faturamento** no projeto GCP — a Vision API exige faturamento mesmo para usar a cota gratuita.

---

## Debug / logs

* No editor do Apps Script: **Executions (Executar → Exibições/Executions)** (novo editor) ou **Ver → Logs** para ver `Logger.log`.
* Cloud Console → **Logging / Logs Explorer** para ver erros detalhados de APIs e OAuth.

---

## Recomendações e limitações

* OCR via `Drive.Files.insert(..., {ocr:true})` funciona para imagens e PDFs simples, mas pode falhar em documentos muito inclinados, fotos de baixa qualidade ou PDFs com layouts complexos.
* Para melhor qualidade em casos difíceis, considere a **Cloud Vision API**, mesmo que exija faturamento (tem cota gratuita inicial).
* Ajuste as **expressões regulares** de extração conforme os recibos/nota fiscais que você recebe (formatos regionais variam).

---

## Exemplo de workflow de troubleshooting (rápido)

1. Teste mínimo: rode `testeDriveSheets()` — se autorizar e criar planilha, permissões básicas OK.
2. Se `Drive.Files.insert` falhar, verifique Advanced Drive Service + GCP project binding.
3. Se OCR falhar em arquivos específicos, teste converter o PDF para imagens e tentar novamente.

---

**Licença:** MIT (uso pessoal / interno).
