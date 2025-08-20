# <img src="./assets/favicon-96x96.png" alt="Ícone Planilista" width="45"> Planilista

[![Versão](https://img.shields.io/badge/vers%C3%A3o-2.5-brightgreen.svg)](https://github.com/chrispsz/planilista)
[![Status](https://img.shields.io/badge/status-ativo-success.svg)](https://github.com/chrispsz/planilista)
[![Licença: MIT](https://img.shields.io/badge/Licen%C3%A7a-MIT-blue.svg)](https://opensource.org/licenses/MIT)

> Transforme a organização por cores da sua planilha em listas de texto e planilhas CSV prontas para usar. Rápido, intuitivo e offline.

---

## 📋 Índice

*   [🚀 Sobre o Projeto](#-sobre-o-projeto)
*   [✨ Principais Funcionalidades](#-principais-funcionalidades)
*   [🛠️ Como Usar](#️-como-usar)
*   [💻 Tecnologias Utilizadas](#-tecnologias-utilizadas)
*   [📝 Licença](#-licença)

---

## 🚀 Sobre o Projeto

O **Planilista** é uma ferramenta web que automatiza a extração de dados de planilhas com base na cor de preenchimento das células. Na versão 2.5, ele evoluiu para uma solução completa de processamento, gerando não apenas listas de texto, mas também planilhas `.csv` formatadas, ideais para importação em outros sistemas ou para a criação de listas de contato.

Com uma interface limpa, filtros inteligentes e opções de exportação flexíveis, o Planilista elimina tarefas manuais demoradas, convertendo seu sistema de cores em dados acionáveis em segundos. Suporta arquivos `.xlsx` e `.xls`, com otimização especial para este último.

## ✨ Principais Funcionalidades

*   **Extração Inteligente por Cores:** Analisa a cor de fundo das células para agrupar e categorizar dados automaticamente.
*   **Exportação Avançada (TXT e CSV):** Gere listas em `.txt` ou crie planilhas `.csv` prontas para importação. Inclui uma ferramenta para mapear DDDs e formatar números de telefone automaticamente.
*   **Controle Fino do Processamento:** Opções para **remover valores duplicados** e **excluir automaticamente itens de confirmação**, limpando seus resultados com um único clique.
*   **Gerenciamento de Configurações:** **Importe e exporte** suas configurações de mapeamento de cores em formato JSON, perfeito para backup ou para compartilhar seu setup.
*   **Mapeamento de Cores Persistente:** Defina quais cores correspondem a quais descrições (ex: `#FFC7CE` = "Pendente"). As configurações são salvas localmente no navegador.
*   **Experiência de Usuário Refinada:**
    *   **Tema Claro e Escuro** para conforto visual.
    *   **Notificações "Toast"** para feedback instantâneo.
    *   **Painel de estatísticas** para uma visão geral dos dados processados.
    *   Interface moderna com ícones **Lucide** para clareza e legibilidade.

## 🛠️ Como Usar

1.  **Upload do Arquivo:** Arraste e solte sua planilha (`.xls` ou `.xlsx`) na área indicada ou clique para selecionar.
2.  **Configuração de Cores:**
    *   Acesse o menu **"Configurar cores"**.
    *   Mapeie os códigos hexadecimais das cores da sua planilha para as descrições desejadas.
    *   *Dica: O Planilista detecta e exibe as cores presentes no arquivo para facilitar o mapeamento.*
3.  **Ajuste do Processamento:**
    *   Selecione a **aba** correta da sua planilha.
    *   Defina a **coluna-alvo** (a coluna com os dados que você quer extrair).
    *   Marque as opções **"Remover duplicados"** ou **"Excluir confirmações"** conforme necessário.
4.  **Processamento e Exportação:**
    *   Clique em **"Reprocessar"** para aplicar as configurações.
    *   Escolha o formato de saída: **Copiar**, **Baixar como .txt** ou **Gerar .csv**.
    *   *Opcional:* Para o CSV, configure previamente os telefones no menu **"Telefone"** para formatação automática.

## 💻 Tecnologias Utilizadas

Este projeto foi construído com foco em performance e simplicidade, rodando 100% no navegador do cliente sem necessidade de um backend.

*   **HTML5:** Estrutura semântica e acessível.
*   **CSS3:** Design responsivo, temas (light/dark mode) e layout moderno com Flexbox e Grid.
*   **JavaScript (Vanilla ES6+):** Toda a lógica da aplicação, manipulação do DOM e gerenciamento de estado local.
*   **SheetJS (xlsx):** Biblioteca essencial para a leitura e o parsing dos arquivos de planilha diretamente no navegador.
*   **Lucide Icons:** Biblioteca de ícones open-source leve e moderna para uma interface clara.

## 📝 Licença

Distribuído sob a Licença MIT.
