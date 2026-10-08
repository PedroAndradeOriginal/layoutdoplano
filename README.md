# Exportação de Layout da Day

Aplicação web local para transformar uma única planilha de efetivo em três formatos de exportação: Plano de Saúde, solicitação do cartão Caju e pagamento mensal da Caju.

O processamento acontece integralmente no navegador. Nenhuma planilha ou informação pessoal é enviada, armazenada ou compartilhada.

## Funcionalidades

- Upload de arquivos Excel nos formatos `.xlsx` e `.xls`.
- Suporte a arrastar e soltar a planilha.
- Seleção da data inicial de admissão.
- Inclusão das admissões da data selecionada e das datas posteriores.
- Escolha entre `Plano de saúde`, `Caju cartão` e `Caju pagamento` no mesmo fluxo.
- Prévia dos registros encontrados antes da exportação.
- Geração do layout com base na planilha-modelo original.
- Geração do CSV da Caju com nome completo, CPF, e-mail e celular.
- Preservação do nome obrigatório `colaboradores-exemplo (20).csv`.
- Manutenção das colunas posteriores ao telefone em branco no arquivo da Caju.
- Geração da planilha mensal da Caju com modelo interno, abas, fórmulas e formatação preservadas.
- Seleção da competência mensal e inclusão de quem esteve ativo em qualquer dia do mês.
- Atualização dos dados cadastrais pelo efetivo atual.
- Manutenção dos parâmetros de benefício do mês anterior por matrícula.
- Identificação de novos colaboradores que precisam de revisão dos benefícios.
- Preservação da formatação, larguras, alturas, mesclagens, filtros, validações e abas do modelo.
- Cálculo automático da vigência para o dia seguinte à admissão.
- Processamento totalmente local e sem necessidade de servidor.
- Interface responsiva com modal de privacidade e animação de carregamento.

## Como usar

1. Baixe ou clone este repositório.
2. Abra a pasta `dist`.
3. Abra o arquivo `index.html` em um navegador moderno.
4. Arraste a planilha de efetivo para a área indicada ou clique para selecioná-la.
5. Em `Caju pagamento`, escolha a nova competência.
6. Confira a quantidade de pessoas e a prévia apresentada.
7. Clique no botão de geração correspondente à opção selecionada.

O arquivo será baixado com um nome semelhante a:

```text
Layout_Hapvida_Admissoes_2026-09-21.xlsx
```

Na opção Caju, o nome do arquivo será sempre:

```text
colaboradores-exemplo (20).csv
```

## Exportação para Caju

O arquivo da Caju mantém as 11 colunas do modelo recebido. O preenchimento automático termina em `Telefone Celular`:

| Coluna da Caju | Origem no efetivo |
|---|---|
| Nome completo | `Nome completo` ou `Nome complet` |
| CPF | `CPF`, somente números |
| Email | `Email Princ` |
| Telefone Celular | `DDD Celular` + `Num. Celular` |
| CEP até Estado | Permanecem em branco |

Os nomes são copiados integralmente, sem abreviações. A aplicação também avisa quando algum registro selecionado está sem e-mail.

## Pagamento mensal da Caju

Na opção `Caju pagamento`, a aplicação usa um modelo sanitizado incorporado ao próprio site e preserva suas sete abas, fórmulas, filtros e formatação.

A competência inclui todas as pessoas que estiveram ativas em pelo menos um dia do mês: admissão até o último dia da competência e desligamento vazio ou igual/posterior ao primeiro dia.

Os cálculos básicos começam zerados, evitando que um benefício seja concedido automaticamente sem conferência.

A coluna numérica adicional exigida pelo modelo mensal é preenchida com zero em todos os registros, sem exibição desse dado na interface.

## Correspondência dos dados

| Layout gerado | Coluna do efetivo ou regra |
|---|---|
| Matrícula | `Matricula` |
| Nome do beneficiário | `Nome complet` |
| Data de nascimento | `Data Nasc.` |
| Sexo | `Sexo` |
| CPF | `CPF` |
| Nome da mãe | `Nome Mae` |
| Vigência do plano | Data de admissão + 1 dia |
| Relação de dependência | `TITULAR` |
| Logradouro | `Endereço` |
| Número | `Num.Endereço` ou `NrLogradouro` |
| Complemento | `Compl.Ender.` |
| Bairro | `Bairro` |
| Município | `Municipio` |
| CEP | `Cep` |
| CPF do titular | Repete o CPF do beneficiário |
| Data de admissão | `Data Admis.` |
| Estado civil | `Est. Civil` |

## Privacidade

A aplicação não possui servidor, banco de dados, ferramenta de análise ou integração externa.

Ao selecionar uma planilha:

- o navegador lê o arquivo diretamente no dispositivo;
- os dados permanecem somente na memória durante o processamento;
- nenhuma informação é enviada pela internet;
- nada é salvo automaticamente após o fechamento da página.

Por conter dados pessoais, as planilhas de efetivo e os arquivos exportados não devem ser adicionados ao repositório.

## Estrutura do projeto

```text
Layout do Plano/
├── README.md
├── README.txt
└── dist/
    ├── index.html
    ├── styles.css
    ├── app.js
    └── assets/
        ├── logo-hc4.png
        ├── logo-hapvida.png
        ├── logo-caju.png
        ├── monthly-template-data.js
        ├── template-data.js
        ├── xlsx.full.min.js
        └── jszip.min.js
```

## Tecnologias

- HTML5
- CSS3
- JavaScript
- SheetJS
- JSZip

As bibliotecas necessárias estão incluídas na pasta do projeto para permitir o funcionamento local, inclusive sem conexão com a internet.

## Requisitos

- Navegador atualizado, como Google Chrome, Microsoft Edge ou Firefox.
- Planilha de efetivo com os nomes de colunas esperados pela aplicação.

Não é necessário instalar dependências, executar comandos ou iniciar um servidor.

## Observações

- O modelo incorporado ao projeto é utilizado como base da exportação para preservar sua estrutura visual.
- A aplicação gera registros de titulares. Dependentes que não estejam presentes na planilha de efetivo não são criados automaticamente.
- Matrículas, CPFs e CEPs são tratados como identificadores para preservar zeros à esquerda.
- As duas versões de efetivo analisadas são aceitas, incluindo as variações `Nome complet` e `Nome completo`.

## Autor

Desenvolvido por **Pedro Andrade**.
