# Controle de Ordens de Produção

Sistema em **Excel + VBA** para controlar Ordens de Produção (OPs): cada máquina tem a sua aba, cada OP é uma linha e cada OP tem um checklist de tarefas (PW, impressão, relatório e sistema) marcado conforme o trabalho é feito.

Funciona somente com o Microsoft Excel (arquivo `.xlsm`), sem instalar nada e sem dependências externas.

> Todos os exemplos de OPs, clientes e itens deste repositório são fictícios.

## O problema que resolve

Antes do sistema, o acompanhamento era manual: OPs anotadas à mão em um caderno, tarefas riscadas no papel e dados digitados a partir de planilhas de programação mensal com layouts diferentes. Isso era lento, sujeito a erro de digitação e difícil de consultar. O sistema centraliza tudo e responde rápido a perguntas como "em qual máquina está a OP tal e o que falta fazer nela?".

## Funcionalidades

- OPs organizadas por máquina e por mês, em blocos (uma Tabela do Excel por mês).
- Checklist por OP, marcado com duplo clique, independente entre OPs, com registro de data e hora.
- Cadastro, edição, exclusão e arquivamento de OPs por formulário, com validação.
- Restauração de OPs arquivadas.
- Bloqueio de OPs duplicadas em toda a base (a OP é comparada por uma chave só com dígitos).
- Busca rápida de uma OP em todas as máquinas e no arquivo.
- Importação de OPs a partir das planilhas de programação, escolhendo apenas as desejadas. O arquivo original nunca é aberto: o sistema trabalha em uma cópia temporária.
- Criação e exclusão de máquinas e setores, com o menu redesenhado automaticamente.

<!-- Adicione prints em docs/imagens/ e referencie aqui, por exemplo:
![Menu](docs/imagens/menu.png) -->

## Requisitos

- Windows
- Microsoft Excel 64 bits, com macros habilitadas
- Um usuário por vez (sem edição simultânea)

## Estrutura do repositório

```
controle-ordens-producao/
├── README.md
├── LICENSE
├── src/
│   ├── modulos/        Módulos VBA (.bas)
│   │   ├── mdlUtil.bas        funções auxiliares
│   │   ├── mdlConfig.bas      configuração, máquinas, setores
│   │   ├── mdlBlocos.bas      blocos de mês
│   │   ├── mdlOP.bas          regras das OPs e checklist
│   │   ├── mdlNavegacao.bas   menu, botões, busca, máquinas
│   │   └── mdlImportar.bas    importação da programação
│   ├── formularios/    Formulários (.frm + .frx): frmOP e frmImportar
│   └── eventos/        Eventos da pasta de trabalho (EstaPastaDeTrabalho.cls)
├── docs/
│   └── Controle_de_Ordens_de_Producao.docx   Documentação funcional e técnica
└── planilha/           Planilha modelo (.xlsm), sem dados reais
```

## Como instalar

1. Crie uma pasta de trabalho no Excel e salve como **Pasta de Trabalho Habilitada para Macro (.xlsm)**, em uma pasta local (de preferência fora de pastas sincronizadas).
2. Abra o editor do VBA (`Alt + F11`).
3. Importe os arquivos de `src/`: clique com o botão direito no projeto > **Importar arquivo** e selecione cada `.bas` de `src/modulos/` e cada `.frm` de `src/formularios/` (o `.frx` precisa estar na mesma pasta do `.frm`).
4. Abra o item `EstaPastaDeTrabalho` e cole o conteúdo de `src/eventos/EstaPastaDeTrabalho.cls` (não é um módulo novo).
5. Menu **Depurar > Compilar VBAProject**. Sem mensagem, está correto.
6. Execute a macro `InstalarSistema`. Ela atribui os setores, migra as abas e desenha o menu e os botões.
7. Salve e teste com uma OP fictícia: cadastrar, marcar caixas, editar, arquivar, restaurar, buscar e excluir.

**Atalho com a planilha modelo:** a pasta `planilha/` traz o `ControleOP_modelo.xlsm`, que já tem as abas, as tabelas, o menu e os botões, mas **sem o código VBA** (o código fica em `src/`). Abra o modelo e siga apenas os passos 2 a 7 acima, importando os módulos, os formulários e o código de eventos. Todos os dados dele são fictícios.

O passo a passo completo está na seção 12 da documentação.

## Decisões técnicas

- **Caixas do checklist como caractere dentro da célula** (☐/☑), e não controles do Excel: filtros e ordenações funcionam sem marcar a OP errada, e um único tratador de evento atende todas as abas.
- **OP guardada como texto e comparada por chave** (só dígitos, sem zeros à esquerda): evita que o Excel transforme `62.620` em número ou data, e torna a comparação imune a pontos e espaços.
- **Mês lido do título do bloco**: a planilha é a fonte da verdade, corrigir um mês é editar um texto visível.
- **Cópia temporária da programação**: o arquivo original, grande e em pasta sincronizada, nunca é aberto nem alterado.
- **Colunas da programação descobertas pelo cabeçalho**, porque as planilhas têm layouts diferentes entre si.
- **Controles dos formulários criados em tempo de execução**: a instalação não exige permissão de acesso ao projeto VBA.

Todas as decisões, com as alternativas descartadas e os motivos, estão na seção 10 da documentação.

## Documentação

A documentação funcional e técnica completa (estrutura da planilha, regras de negócio, telas, módulos, fluxos, validações, limitações e perguntas frequentes) está em [`docs/Controle_de_Ordens_de_Producao.docx`](docs/Controle_de_Ordens_de_Producao.docx).

## Autor

Luís Eduardo Sales de Oliveira

## Licença

Distribuído sob a licença MIT. Veja o arquivo [LICENSE](LICENSE).
