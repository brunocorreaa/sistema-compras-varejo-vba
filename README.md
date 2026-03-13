# 🛒 Sistema Integrado de Planejamento e Governança de Compras

## 📌 Visão Geral
Este sistema representa o **núcleo estratégico de uma operação de varejo de larga escala**. Mais do que uma planilha, trata-se de um ecossistema de gestão desenvolvido em VBA que atua como o "Cérebro Operacional" no planejamento de compras. O sistema centraliza a evolução de verbas, cadastro de produtos, cotações de fornecedores e a emissão de ordens de compra, garantindo que bilhões em potencial de investimento sejam geridos com precisão absoluta.

## 🏗️ Arquitetura e Complexidade Técnica
A complexidade deste projeto é evidenciada pela orquestração de **mais de 32 macros e formulários interconectados**, desenhados para suportar um fluxo de trabalho de missão crítica:

* **Matriz de Dados Robusta:** O sistema gerencia um "Planilhão" com **mais de 100 colunas dinâmicas**, onde cada campo possui regras de negócio específicas e dependências lógicas.
* **Teia de Dependências:** Diferente de macros isoladas, aqui as rotinas são sensíveis e intimamente ligadas. A saída de um módulo de *Cotação* é o gatilho de validação para o módulo de *Pedido*, exigindo uma visão sistêmica para qualquer manutenção ou atualização.
* **Engine de Visibilidade:** Para manter a usabilidade em um ambiente de 100+ colunas, o sistema conta com uma lógica de UI que oculta ou exibe campos em tempo real conforme o perfil do usuário e a fase do planejamento.

## 🛡️ Governança e Rigor de Dados (Compliance)
Devido ao alto valor financeiro envolvido nas decisões tomadas dentro desta ferramenta, a **Governança** é o pilar central:

1.  **Validações em Camadas:** O sistema executa uma bateria de testes de integridade antes de qualquer processamento. Verificações de divergência entre Tamanho e EAN, estouro de verba (budget), valores negativos e inconsistências de calendário são processadas instantaneamente para impedir o erro humano.
2.  **Audit Log (Rastreabilidade Total):** Cada ação dentro do arquivo é logada. O sistema registra quem executou a macro, o timestamp, o tempo de ociosidade e o status final. Isso permite uma auditoria completa de cada ciclo de planejamento.
3.  **Gestão de Erros Profissional:** Estruturas avançadas de *Error Trapping* garantem a continuidade da operação. Erros não previstos são capturados e registrados em um log de depuração, impedindo o corrompimento do arquivo e facilitando a correção técnica.

## 🚀 Módulos Principais e Automação
* **RPA para Fornecedores:** Automação do ciclo de vida da cotação. O sistema gera arquivos técnicos, anexa etiquetas específicas (com regras por marca/setor) e manuais de processo, realizando o envio automático via Outlook API.
* **Inteligência de Mix e Clonagem:** Lógica complexa para replicação de itens que preserva vínculos de ID e integridade de dados, permitindo a escalabilidade do planejamento de coleções.
* **Processamento de Imagens:** Estrutura que manipula fotos de produtos dinamicamente, ajustando escalas e centralizações para a geração de catálogos e relatórios consolidados.

## 🛡️ Privacidade e Nota Técnica
Este repositório contém uma **versão estritamente lógica e anonimizada**. 
* Todas as fórmulas de negócio, conexões de Power Query e dados reais de fornecedores foram removidos.
* O foco aqui é demonstrar a **capacidade de engenharia de software**, arquitetura de sistemas em ambiente Office e a robustez das automações aplicadas em um cenário real e complexo de compras.
