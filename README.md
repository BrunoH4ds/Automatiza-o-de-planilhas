# Automação de Planilhas Excel

Este projeto tem como objetivo automatizar a organização de dados em uma planilha Excel. Através do uso da biblioteca `openpyxl`, o script lê uma aba de dados, cria novas abas com base em valores únicos de uma coluna específica e transfere as informações relevantes para essas abas.

## Funcionalidades

- Criação de novas abas em um arquivo Excel com base nos valores de um bairro.
- Transferência de dados de uma aba base para as novas abas criadas.
- Cabeçalhos pré-definidos nas novas abas para melhor organização dos dados.

## Pré-requisitos

Para executar este projeto, você precisará ter Python e as seguintes bibliotecas instaladas:

- `openpyxl`: Para manipulação de arquivos Excel.
- `copy`: Para copiar estilos de células.

Caso não tenha as bibliotecas instaladas, o script já contém um mecanismo para instalá-las automaticamente. 

## Instalação

Execute o script em um ambiente Python. Ele tentará instalar automaticamente as bibliotecas necessárias. Caso as bibliotecas já estejam instaladas, o script seguirá sua execução normalmente.

## Uso

- **Preparação do Arquivo:** O arquivo Excel deve conter uma aba chamada "Base de Dados" com uma coluna de bairros na coluna C.
- **Executar o Script:** Execute o script Python para criar novas abas com os nomes dos bairros e transferir as informações.

## Salvar o Arquivo

Após a execução do script, as alterações serão salvas no arquivo original Bairros.xlsx.

## Observações

- Estrutura da Planilha: Certifique-se de que a planilha contém os dados corretamente formatados antes de executar o script.
- Cuidado com a Sobrescrição: O script sobrescreve o arquivo original. Mantenha uma cópia de segurança caso necessário.
  
## Contribuições

Sinta-se à vontade para contribuir com melhorias e sugestões para este projeto.

## Licença
Este projeto está licenciado sob a MIT License - veja o arquivo LICENSE para mais detalhes.
