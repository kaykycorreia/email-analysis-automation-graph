# email-analysis-automation-graph
Automação em Python para leitura de e-mails via Microsoft Graph API, geração de relatórios Excel e consolidação automática de chamados.

# Automação de Análise de E-mails com Microsoft Graph API

## 📌 Visão Geral
Este projeto consiste em uma solução de automação desenvolvida em Python para análise de e-mails corporativos utilizando a **Microsoft Graph API**.  
A automação realiza a leitura da caixa de entrada, identifica e-mails contendo uma palavra-chave específica, gera relatórios em Excel e consolida os dados automaticamente para análise de volume de chamados.

O objetivo é transformar a leitura manual de e-mails em **dados estruturados**, facilitando o acompanhamento de demandas, incidentes e padrões recorrentes.

---

## ⚙️ Funcionalidades Principais
- Autenticação segura no Microsoft Azure (OAuth 2.0)
- Leitura automática da caixa de entrada do Outlook
- Filtro de e-mails por:
  - Intervalo de datas
  - Palavra-chave no assunto ou corpo do e-mail
- Geração automática de relatório Excel
- Consolidação dos chamados em uma aba de resumo
- Contagem e ordenação dos chamados mais recorrentes
- Organização automática dos relatórios em pastas específicas
- Geração de logs para auditoria e monitoramento

---

## 🧠 Como a Solução Funciona
1. O script se autentica no Azure utilizando Microsoft Graph API
2. Realiza a leitura paginada dos e-mails da caixa de entrada
3. Filtra mensagens com base em palavra-chave definida pelo usuário
4. Gera um relatório Excel com os e-mails encontrados
5. Processa o relatório:
   - Normaliza os títulos
   - Agrupa chamados semelhantes
   - Cria uma aba de resumo com quantidade de ocorrências
6. Move os arquivos para pastas organizadas
7. Registra toda a execução em logs

---

## 📊 Resultado Final
- Relatórios estruturados em Excel
- Aba de resumo com:
  - Tipo de chamado
  - Quantidade de ocorrências
- Visão clara dos principais motivos de contato por e-mail
- Redução significativa do tempo gasto em análise manual

---

## 🛠️ Tecnologias Utilizadas
- Python
- Microsoft Graph API
- MSAL (Microsoft Authentication Library)
- Pandas
- Requests
- OpenPyXL
- OAuth 2.0 (Azure AD)

---

## 🎯 Casos de Uso
- Suporte de TI (análise de chamados via e-mail)
- Gestão de incidentes
- Monitoramento de demandas recorrentes
- Geração de indicadores operacionais
- Apoio à tomada de decisão

---

## ⚠️ Observações Importantes
- As credenciais do Azure devem ser configuradas via variáveis de ambiente
- Não utilize este script com dados sensíveis em ambientes públicos
- Recomenda-se execução em ambiente controlado

---

## 📄 Licença
Projeto desenvolvido para fins educacionais, automação de processos e demonstração técnica.

