# Sistema de Controle de Acesso Escolar

## Sobre o Projeto

Sistema de controle de acesso desenvolvido em VBA para Microsoft Excel entre junho de 2022 e agosto de 2023, **aplicado em contexto real no CIE Miécimo da Silva**.

### Problema Original
A escola enfrentava inconsistências no controle do refeitório, sem forma confiável de rastrear os acessos e prevenir duplicações.

### Solução Desenvolvida
Sistema com leitura de crachá que registrava automaticamente os acessos, implementando validações que impediam registros duplicados e garantiam rastreabilidade completa.

### Evolução do Projeto
Devido ao sucesso da solução, o sistema foi posteriormente **adaptado para controle de entrada e saída geral da escola**, melhorando significativamente a segurança e organização.

### Impacto
- Eliminou inconsistências no controle do refeitório
- Preveniu acessos duplicados
- Melhorou segurança no controle de entrada/saída de alunos no ambiente escolar
- Sistema utilizado diariamente pela equipe administrativa

## Funcionalidades

* Interface central para acessar todas as funcionalidades do sistema.
* Registro de Entrada e Saída: Permite registrar o ponto de alunos.
* Registro por Pesquisa: Buscar registros existentes por nome, matrícula ou outro critério.
* Salvar Planilha: Salva os dados atuais do sistema.
* Backup de Planilha: Cria cópias de segurança das planilhas.
* Limpeza de Registros: Remove registros antigos.


### Pré-requisitos
* Microsoft Excel 2010 ou superior
* Macros habilitadas: Vá em **Arquivo → Opções → Central de Confiabilidade → Configurações de Macro → Habilitar todas as macros**.

### Como usar
1. Baixe a pasta `planilhas/`.
2. Abra a planilha desejada no Excel.
3. Use o botão já criado ou crie um botão e vincule o seguinte código: "TelaPrincipal.Show".


### Configurações
* Diretório de backup: Pasta onde serão salvos os backups automáticos das planilhas.
* Diretório de foto: Pasta onde as fotos serão armazenadas.
* Mensagens do controle de acesso: Personalize as mensagens exibidas ao registrar entradas.
* Senha das planilhas: Defina uma senha para proteger o sistema e os dados das planilhas.


## Módulos e Formulários

### Módulos 

- var.bas: Variáveis globais do sistema
- SalvarLimpar.bas: Funções de backup e limpeza
- NumDeRegistrados.bas: Contagem de registros
- Relogio.bas: Funções de data/hora
- ScrollMouse.bas: Scroll aprimorado nos formulários

### Formulários

- TelaPrincipal: Interface principal do sistema, para acessar as funcionalidades.
- Verificador / VerificadorSaida / Verificador2 / Verificador3: Formulários responsáveis pelo controle de entrada e saída.
- Pesquisa: Formulário para buscar registros existentes por nome, matrícula ou outro critério.

## Código

O código VBA está organizado dentro de src/, dividido por funcionalidades:
- entrada-saida/ → módulos e forms de controle de entrada e saída
- refeitorio/ → módulos e forms de controle do refeitório
