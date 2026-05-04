# 📊 Projeto: Ecossistema de Gestão Operacional Integrada (ERP-BI)

Este projeto simula o ciclo completo de operações administrativas, financeiras e logísticas de uma organização. Utilizando **Excel avançado** e **integração com Word**, a solução transforma dados brutos em um painel de Business Intelligence (BI) capaz de guiar decisões estratégicas.
 
--- 

## 🏗️ Arquitetura do Sistema

O projeto foi construído sob o princípio da **normalização de dados**, garantindo que informações de Clientes, Fornecedores e Produtos fossem relacionadas através de chaves primárias (IDs), evitando redundância.

### Camada de Simulação e Geração de Dados (DB_Movimentação)

Diferente de planilhas estáticas, este projeto utiliza lógica de programação em Excel para simular 1000 transações baseadas em regras de negócio.

Abaixo, as fórmulas que compõem o motor de dados:

- **Gerenciamento de Linha do Tempo:** Gera uma cronologia realista de eventos em ordem crescente utilizando a função `MATRIZALEATÓRIA` combinada com a função `CLASSIFICAR` em uma coluna fora da tabela oficial e depois atribui os valores para a coluna `Data` localizada dentro da tabela oficial do excel.

    - *Fórmula para a geração da cronologia na `Coluna O`:* 
        `=CLASSIFICAR(MATRIZALEATÓRIA(1000;1;$B$21;HOJE();VERDADEIRO))`
    
    - *Fórmula para a atribuição dos dados:*
        `=SE($O$7="Desbloqueado", O22, [@Data])`

- **ID da Transação:**
    - *Fórmula para geração de id:*
        `="T"&TEXTO(LIN(A22)-1;"0000")`
 
- **Tipo de Fluxo (Simulação de Fluxo com Inteligência de Estoque):** Utiliza a função `LET` para garantir que saídas ocorram apenas se houver saldo positivo. Soma as entradas e subtrai as saídas em uma variável `saldo_atual`. Permite um fluxo de saída apenas com `saldo_atual` positivo. Detalhe para a estrutura condicinal que utiliza lógica `AND` de forma eficiente para determinar se o  valor será somado ou subtraído da variável `saldo_atual`.
 
    - *Fórmula para determinar o tipo de fluxo:*
    
        ```excel
        =SE(
            $O$7="Desbloqueado";
            LET(
                saldo_atual;
                SOMA(
                    ($C$2:$C21="Entrada")*($F$2:$F21=[@[ID Produto]])*($H$2:$H21) 
                    - ($C$2:$C21="Saida")*($F$2:$F21=[@[ID Produto]])*($H$2:$H21)
                );
                SE(
                    saldo_atual > 0;
                        SE(
                            ALEATÓRIOENTRE(0;10)>3;
                            "Saida";
                            "Entrada"
                        );
                    "Entrada"
                )
            );
            [@[Tipo Fluxo]]
        )    
        ```
    
- **Identificação de Entidades e Produtos:**

    - **ID Entidade:** Gera um identificador aleatório de cliente ou fornecedor dentro do limite de cada tabela.
        - *Fórmula:*
        ```excel
         =SE(
            $O$7="Desbloqueado";
            SE(
                [@[Tipo Fluxo]]="Saida";
                "CLI"&TEXTO(ALEATÓRIOENTRE(1;20);"000"); "FORN"&TEXTO(ALEATÓRIOENTRE(1;20);"000")
            );
            [@[ID Entidade]]
        )
        ```

    - **Nome da Entidade (`PROCX`):** Realiza a busca do nome do cliente ou fornecedor nas respectivas tabelas de cadastro utilizando a função `PROCX`. Na ordem são enviados: o identificador a ser localizado, a tabela e a coluna onde será realizada a busca e, for último, a tabela e a coluna onde estará o valor a ser retornado.    
        - *Fórmula:* 
        ```excel
        =SE(
            [@[Tipo Fluxo]]="Saida";
            PROCX(
                [@[ID Entidade]];
                Tabela_Clientes[ID_Cliente];
                Tabela_Clientes[Nome_Cliente];
                "Cliente não encontrado"
            );
            PROCX(
                [@[ID Entidade]];
                Tabela_Fornecedores[ID_Fornecedor];
                Tabela_Fornecedores[Nome_Fantasia];
                "Fornecedor não encontrado"
            )
        )
        ```  

    - **ID do prduto:** Gera um identificador aleatório de produto dentro do limite de cada tabela.    
        - *Fórmula:*
        ```excel
        =SE(
            $O$7="Desbloqueado";
            "PROD"&TEXTO(ALEATÓRIOENTRE(1;200);"000");
            [@[ID Produto]]
        )
         ```

    - **Nome da Entidade (`PROCX`):** Realiza a busca do produto na respectiva tabela de cadastro utilizando a função `PROCX`. Na ordem são enviados: o identificador a ser localizado, a tabela e a coluna onde será realizada a busca e, for último, a tabela e a coluna onde estará o valor a ser retornado.
        - *Fórmula:*
        ```excel
        =PROCX(
            [@[ID Produto]],
            Tabela_Produtos[ID_Produto],
            Tabela_Produtos[Descricao_Produto],
            "Produto não encontrado"
        )    
        ```    

- **Cálculo de Quantidade e Valores:** 

    - **Quantidade Dinâmica:** A simulação da quantidade de uma compra ou venda é determinada com base no histórico de compra e venda desse produto. Esse campo considera que o campo `Tipo Fluxo` já analisou esse histório bloqueando uma simulação de `saída` com estoque negativo. Desta forma, a fórmula é responsável por calcular o valor do `saldo_atual`e simular uma compra ou venda com valores aleatórios de modo que o `saldo_atual` após uma simulação de `saída` seja sempre um valor positivo não nulo e após uma simulação de `entrada` seja sempre um valor acima do valor `estoque_minimo`.

        - *Fórmula:* 
        ```excel
        =SE(
            $O$7="Desbloqueado",
            LET(
                saldo_atual,
                SOMA(
                    ($C$2:$C21="Entrada")*($F$2:$F21=[@[ID Produto]])*($H$2:$H21) 
                    - ($C$2:$C21="Saida")*($F$2:$F21=[@[ID Produto]])*($H$2:$H21)
                ),                
                estoque_minimo,
                PROCX(
                    [@[ID Produto]],
                    Tabela_Produtos[ID_Produto],
                    Tabela_Produtos[Estoque_Minimo],
                    0
                ),
                SE(
                    saldo_atual > estoque_minimo,
                    SE(
                        [@[Tipo Fluxo]]="Saida",
                        ALEATÓRIOENTRE(1, saldo_atual),
                        ALEATÓRIOENTRE(1, 100)
                    ),
                    SE(
                        [@[Tipo Fluxo]]="Saida",
                        ALEATÓRIOENTRE(1, saldo_atual),
                        SE(
                        saldo_atual < 0,
                        ABS(saldo_atual) + estoque_minimo + ALEATÓRIOENTRE(1, 100),
                        estoque_minimo - saldo_atual + ALEATÓRIOENTRE(1, 100)
                        )
                    )
                )
            ),
            [@Quantidade]
        )
        ```

    - **Valor Unitário:** Valor é localizado na tabela de produtos através da função `PROCX`, sendo utilizado os argumentos: índice do produto, nome da tabela de produtos com a coluna onde procurar o índice e o nome da tabela com a coluna onde procurar o valor do produto.        
        - *Fórmula:*            
        ```excel
        =PROCX(
            [@[ID Produto]],
            Tabela_Produtos[ID_Produto],
            Tabela_Produtos[Preco_Unitario_Venda]
        )
        ```

    - **Total:** Utiliza os campos de alor unitário e quantidade de itens presentes na tabela.
        - *Fórmula:*
        ```excel
            =[@Quantidade] * [@[Valor Unitário]]
        ```

- **Status e Previsão:**

    - **Financeiro:**
        - *Fórmula:*
        ```excel
        =SE(
            $O$7="Desbloqueado",
            ESCOLHER(ALEATÓRIOENTRE(1,3),"Pago","Pendente","Atrasado"),
            [@[Status Financeiro]]
        )        
        ```

    - **Logístico:**
        - *Fórmula:*
        ```excel
        =SE(
            $O$7="Desbloqueado",
            ESCOLHER(ALEATÓRIOENTRE(1,3),"Entregue","Em Rota","Aguardando"),
            [@[Status Logistico]]
        )     
        ```        

    - **Previsão de Entrega:**
        - *Fórmula:*
        ```excel
        =SE(
            $O$7="Desbloqueado",
            SE(
                [@[Tipo Fluxo]]="Entrada",
                [@Data] + PROCX([@[ID Entidade]],Tabela_Fornecedores[ID_Fornecedor],Tabela_Fornecedores[Lead_Time]),
                "N/A"
            ),
            [@[Previsão de Entrega]]
        )
        ```
---

## 🛠️ Detalhamento Técnico: Fórmulas de Negócio

### 📈 Controle de Estoque (Lógica de Inventário)

O controle de inventário processa o histórico da `DB_Movimentação` em tempo real.

- **Cálculo de Entradas/Saídas:** Utilização da função `SOMASES` para filtrar movimentações por ID de Produto e Tipo de Operação.
    - **Entradas:** Soma todas as quantidades de produtos que atendam aos critérios de índice de produto e fluxo de entrada.
        - *Fórmula:*
        ```excel
        =SOMASES(
            Tabela_Movimentacao_Geral[Quantidade],
            Tabela_Movimentacao_Geral[ID Produto],
            CONTROLE_ESTOQUE!$A2,
            Tabela_Movimentacao_Geral[Tipo Fluxo],
            "Entrada"
        )
        ```

    - **Saídas:** Soma todas as quantidades de produtos que atendam aos critérios de índice de produto e fluxo de saída.
        - *Fórmula:*
        ```excel
        =SOMASES(
            Tabela_Movimentacao_Geral[Quantidade],
            Tabela_Movimentacao_Geral[ID Produto],
            CONTROLE_ESTOQUE!$A2,
            Tabela_Movimentacao_Geral[Tipo Fluxo],
            "Saida"
        )
        ```

- **Lógica de Reposição:** Compara o valor do estoque atual com o valor de estoque mínimo para gerar um status de controle de estoque, informando a necessidade de reposição do estoque.

    $$Status = \begin{cases} \text{Comprar Urgente}, & \text{se } Saldo < Estoque Minimo \\ \text{Estoque OK}, & \text{caso contrário} \end{cases}$$
    
    - *Fórmula:* 
    ```excel
    =SE(
        CONTROLE_ESTOQUE!$F2 < CONTROLE_ESTOQUE!$C2;
        "COMPRAR URGENTE";
        "Estoque OK"
    )
    ```
    
## 💰 Gestão de Fluxo de Caixa (Financeiro)

Foram desenvolvidas duas visões financeiras: **Entrada de Caixa** (Recebíveis de Clientes) e **Saída de Caixa** (Contas a Pagar a Fornecedores).

- **Análise do Faturamento:** O sistema isola valores com status "Pago" para monitorar a estabilidade financeira.

    - *Fórmula:*
    ```excel
    =SOMASES(
        Tabela_Movimentacao_Geral[Valor Total];
        Tabela_Movimentacao_Geral[Tipo Fluxo];
        "Saida"/"Entrada";
        Tabela_Movimentacao_Geral[Status Financeiro];
        "Pago"
    )
    ```

- **Análise de Inadimplência:** O sistema isola valores com status "Atrasado" para monitorar o risco financeiro.

    - *Fórmula:* 
    ```excel
    =SOMASES(
        Tabela_Movimentacao_Geral[Valor Total];
        Tabela_Movimentacao_Geral[Tipo Fluxo];
        "Saida";
        Tabela_Movimentacao_Geral[Status Financeiro];
        "Atrasado"
    )
    ```
    
- **Contas a Receber/Pagar:** O sistema isola valores com status "Pendente" para monitorar o risco de liquidez no fluxo e entrada e o risco de crédito no fluxo de saída.
    
    - *Fórmula:* 
    ```excel    
    =SOMASES(
        Tabela_Movimentacao_Geral[Valor Total];
        Tabela_Movimentacao_Geral[Tipo Fluxo];
        "Saida"/"Entrada";
        Tabela_Movimentacao_Geral[Status Financeiro];
        "Pendente"
    )
    ```
  
- **Total Consolidado:** Soma das movimentações pagas, atrasadas e pendentes de um fluxo.

    - *Fórmula:* 
    ```excel
    =SOMA(
        Dashboard[@[Faturamento]:[Contas a Receber]]
    )
    ```
---

## 🎨 Interface e Experiência do Usuário (UX)

Para garantir que o Dashboard seja intuitivo, foi implementado **Rótulos Dinâmicos** que se ajustam conforme o filtro de contexto selecionado (valor da cédula `CALCULUS!$A$11`):

- **Título Fluxo de Caixa:**
    - *Fórmula:*
    ```excel
    =SE(
        CALCULUS!$A$11="Saida";
        "Fluxo de Caixa (Saída)";
        SE(
            CALCULUS!$A$11="Entrada";
            "Fluxo de Caixa (Entrada)";
            "Fluxo de Caixa (Total)"
        )
    )
    ```

- **Título Logística:**
    - *Fórmula:*
    ```excel
    =SE(
        CALCULUS!$A$11="Saida"
        ;"Monitoramento de Pedidos e Entregas (Saídas)";
        SE(
            CALCULUS!$A$11="Entrada";
            "Monitoramento de Pedidos e Entregas (Entradas)";
            "Monitoramento de Pedidos e Entregas (Total)"
        )
    )
    ```

- **Título Evolução Mensal:**
    - *Fórmula:*
    ```excel    
    =SE(
        CALCULUS!$A$11="Saida";
        "Evolução do Faturamento Mensal (Saídas)";
        SE(
            CALCULUS!$A$11="Entrada";
            "Evolução do Faturamento Mensal (Entradas)";
            "Evolução do Faturamento Mensal (Total)"
        )
    )
    ```

- **Título Ranking de Produtos:**
    - *Fórmula:*
    ```excel
    =SE(
        CALCULUS!$A$11="Saida";
        "Ranking de Receita por Produto (Saídas)";
        SE(
            CALCULUS!$A$11="Entrada";
            "Ranking de Receita por Produto (Entradas)";
            "Ranking de Receita por Produto (Total)"
        )
    )
    ```

---

## 📂 Integração Multissistemas (Excel + Word)

### (TÓPICO SENDO IMPLEMENTADO)

Para atender aos requisitos de apoio administrativo, o projeto inclui:

1. **Relatórios Executivos:** Vinculação de gráficos dinâmicos do Excel no Word via **Object Linking and Embedding (OLE)**, garantindo que o relatório mensal se atualize com um clique.

2. **Automação de Cotações (Mala Direta):** Sistema de **Mala Direta** vinculado à base de dados para gerar automaticamente cartas de cotação para os fornecedores que atingiram o nível crítico de reposição.

---

## 🚀 Skills Demonstradas

- **Excel:** Funções Matriciais (`LET`, `MATRIZALEATÓRIA`), Funções de Busca Avançada (`PROCX`), Lógica Condicional Complexa e Dashboards Dinâmicos com Títulos Adaptativos.

- **Administrativo/Financeiro:** Controle de estoque, fluxo de caixa, análise de faturamento, gestão de inadimplência e logística de suprimentos.

- **ADS:** Estruturação de dados relacionais, lógica de simulação, tratamento de erros e automação de processos.
 
 --- 

## 🎓 Nota do Desenvolvedor

Este projeto foi concebido para demonstrar como o raciocínio lógico de Análise e Desenvolvimento de Sistemas pode elevar a eficiência de tarefas administrativas tradicionais, transformando planilhas em ferramentas robustas de gestão de dados.
