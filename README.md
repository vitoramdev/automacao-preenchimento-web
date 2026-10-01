# Automação de Cadastro Web

Aplicação desenvolvida em Python para automatizar o cadastro de produtos em sistemas web a partir de dados armazenados em arquivos Excel ou CSV.

O objetivo do projeto é reduzir tarefas manuais e repetitivas, realizando automaticamente o preenchimento e envio dos dados no sistema.

## Funcionalidades

* Leitura de arquivos Excel e CSV
* Seleção do arquivo diretamente pela interface
* Preenchimento automatizado de formulários web
* Automação do navegador utilizando Playwright
* Execução em modo normal ou em segundo plano (Headless)
* Barra de progresso da execução
* Registro das operações realizadas
* Interface gráfica desenvolvida com CustomTkinter

## Tecnologias utilizadas

* Python
* Playwright
* Pandas
* CustomTkinter
* Excel / CSV

## Como funciona

O usuário informa a URL do sistema e suas credenciais de acesso, seleciona a base de dados em Excel ou CSV e inicia a automação.

A aplicação lê os dados da planilha e utiliza o Playwright para acessar o sistema, preencher os campos e enviar os registros automaticamente.

## Estrutura do projeto

```text
auto.py       → código principal da aplicação
auto.spec     → configuração utilizada para geração do executável
produtos.csv  → arquivo de exemplo utilizado nos testes
```

## Demonstração

Vídeo demonstrando o funcionamento da automação:

https://youtu.be/S5eSRPbXhK4

## Objetivo do projeto

Projeto desenvolvido como aplicação prática de conhecimentos em Python, automação web, manipulação de dados e desenvolvimento de interfaces gráficas.
