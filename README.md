# 📦 Gerenciador de Ativos (EPI & Fardamento)

Este projeto é um sistema de controle de estoque e movimentação de ativos desenvolvido em **Google Apps Script**. O objetivo é digitalizar o registro de entradas e saídas, eliminando o uso de papel e agilizando a consulta de informações para colaboradores e gestores.

---

## 🚀 1. Guia do Usuário

### A. Introdução e Acesso
* **Objetivo:** Centralizar o controle de estoque, registrar todas as movimentações e facilitar o dia a dia com consultas rápidas e simples.
* **Primeiros Passos:**
    1. Crie uma pasta no seu Google Drive chamada `Gerenciador de Ativos`.
    2. Dentro da pasta, crie um novo arquivo de script do **Google Apps Script**.
    3. Ao implantar ou rodar pela primeira vez, autorize as permissões necessárias na sua conta Google.
* **Criando o Ambiente:** O Apps Script não suporta nativamente arquivos `.js` ou `.css`. Por isso, crie os seguintes arquivos (todos com extensão `.html` onde indicado):
    * `pos.gs`: Onde ficará todo o Backend.
    * `get.gs`: Onde é injetado o HTML.
    * `index.html`: Estrutura principal da página (HTML).
    * `javascript.html`: Funções e lógicas JavaScript do Frontend.
    * `style.html`: Folhas de estilo (CSS) do projeto.
* **Banco de Dados:** Na mesma pasta, crie uma **Google Planilha** chamada `Banco`. Ela deve conter as seguintes abas (escritas exatamente assim):
    * `estoque`: ID, Código, Descrição, Quantidade, Tamanho, Extra.
    * `registroepi`: id, data retirada, código, descrição, quantidade, tamanho, matricula, nome, ca, data devolução, motivo.
    * `colaboradores`: ID, Matricula, Nome, Função.
    * `movimentações`: id, data, código, descrição, quantidade, tamanho, matricula, nome, Data devolução.
    * `epimovimentacoes`: id, matricula, nome, função, epi, data entrega, data devolução, c.a, descrição, código.
* **Integração:** 1. Copie o ID da sua planilha (localizado na URL entre a 5ª barra e a última).
    2. No arquivo `pos.gs`, localize a variável `const ss` e cole o ID:
       `const ss = SpreadsheetApp.openById("SEU_ID_AQUI");`

---

### B. Guia de Operação

#### 📊 Dashboard (Página Inicial)
* **Card 1 (Estoque):** Exibe a quantidade total de itens e a porcentagem de disponibilidade. Mostra a data do dia atual.
* **Card 2 (Estoque Baixo):** Alerta para itens com menos de **7 unidades**. Exibe o item com menor quantidade.
* **Card 3 (Retiradas do Dia):** Mostra a quantidade de itens retirados hoje e o nome do último item entregue.
* **Últimos Itens:** Tabela com os itens que foram retirados no dia atual.

#### ➕ Adicionar e Editar
* **Adicionar Itens:** Cadastro de novos itens informando Código, Descrição, Quantidade, Tamanho e C.A (se for EPI).
* **Editar Item:** Permite alterar Descrição, Quantidade, Tamanho e C.A através do Código (o campo Código fica bloqueado para edição).

#### 📤 Retirar (Saída de Material)
1. Informe o **Código** do item e a **Matrícula** do colaborador (use Enter ou Lupa para buscar).
2. Selecione o **Destino**: Fardamento ou EPI.
   * Se for **EPI**, selecione obrigatoriamente o **Motivo**.
3. Clique em **Adicionar** para colocar o item na lista de conferência temporária.
4. **Finalizar Retirada:** Envia os dados para a planilha e atualiza o estoque automaticamente.

#### 👥 Colaboradores
* **Adicionar:** Registro de Matrícula, Nome e Função.
* **Buscar:** Consulta e edita os dados de colaboradores existentes através da matrícula.

#### 🧹 Higienização de EPIs
* **Retirada/Devolução:** Controle específico de EPIs que saem para limpeza. A devolução exige a validação do **Nº do EPI Registrado**.
* **Relatório:** Gera um PDF com os registros do colaborador. O sistema bloqueia a impressão se houver pendências de devolução.

#### 👕 Fardamentos e EPIs (Geral)
* **Devolução:** Telas específicas para baixar itens pendentes informando a data de devolução.
* **Acompanhamento:** Gera relatórios detalhados por matrícula para impressão.

---

### C. Solução de Problemas

| Problema | O que fazer |
| :--- | :--- |
| **"Item não cadastrado"** | Verifique se o código existe na aba `estoque` da planilha. |
| **"Loader" travado** | Atualize a página (F5) ou clique no ícone de recarregar. |
| **"Colaborador não encontrado"** | Verifique a matrícula ou realize um novo cadastro no menu Colaboradores. |
| **Item sem estoque** | Vá em Adicionar > Editar Item e atualize a quantidade. |

---

## 💡 2. Dicas e Atalhos

* **Agilidade:** A tecla `ENTER` realiza a busca automaticamente, sem necessidade de clicar na lupa.
* **Fechamento:** Clicar fora de qualquer janela modal (pop-up) fechará a mesma.
* **Identificação Visual:**
    * 🟡 **Amarelo:** Identifica um **EPI** na lista de retirada.
    * 🔵 **Azul:** Identifica um **Fardamento** na lista de retirada.
    * 🔴 **Vermelho:** Números de quantidade na aba estoque ficam vermelhos se forem menores que 7.

---

## 📩 3. Contato e Suporte

Dúvidas ou sugestões? Entre em contato:
📧 **Email:** pedrosimaocontato@gmail.com