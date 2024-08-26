# Automatizador de Busca de Ofertas

Este projeto consiste em uma automação desenvolvida com Selenium, que busca por ofertas de produtos em sites como Google Shopping e Buscapé. O script filtra os produtos com base em um intervalo de preço e evita termos banidos. Os resultados encontrados são exportados para uma planilha e enviados por e-mail.

## Requisitos

Certifique-se de ter as seguintes bibliotecas instaladas:

- `selenium`
- `pandas`
- `openpyxl`

Além disso, é necessário ter o [ChromeDriver](https://sites.google.com/chromium.org/driver/) instalado e configurado no seu PATH.

### Instalação do ChromeDriver

1. Acesse o site oficial do [ChromeDriver](https://sites.google.com/chromium.org/driver/).
2. Baixe a versão compatível com o seu navegador Google Chrome.
3. Extraia o arquivo baixado e mova o executável para uma pasta do sistema ou adicione ao PATH.

## Configuração

1. Prepare um arquivo `buscas.xlsx` contendo as colunas:
   - **Nome**: Nome do produto a ser pesquisado.
   - **Termos banidos**: Termos que você deseja excluir das pesquisas.
   - **Preço mínimo**: Preço mínimo do produto.
   - **Preço máximo**: Preço máximo do produto.

2. O arquivo `buscas.xlsx` pode ter a seguinte estrutura:

| Nome          | Termos banidos    | Preço mínimo | Preço máximo |
|---------------|-------------------|--------------|--------------|
| Notebook Dell | usado,refurbished  | 2500         | 5000         |

## Funcionalidades

- **Google Shopping**: Realiza buscas no Google Shopping por produtos especificados no arquivo `buscas.xlsx`.
- **Buscapé**: Realiza buscas no Buscapé de acordo com as especificações fornecidas.
- **Envio de E-mail**: Envia um e-mail com os produtos encontrados, em formato HTML, utilizando o Microsoft Outlook.

## Como Utilizar

1. Execute o script `main.py`:

2. O script buscará os produtos em ambas as plataformas e irá gerar um arquivo `Ofertas.xlsx` contendo as ofertas encontradas.

3. Se houver ofertas dentro da faixa de preço desejada, um e-mail será enviado automaticamente com a lista das ofertas em formato HTML.

## Estrutura do Código

- **Funções Auxiliares**:
  - `verificar_termos_banidos()`: Verifica se o nome do produto contém termos banidos.
  - `verificar_tem_todos_termos_produtos()`: Verifica se todos os termos do nome do produto estão presentes.
  
- **Funções Principais**:
  - `busca_google_shopping()`: Faz a busca no Google Shopping, filtra os resultados e retorna as ofertas.
  - `busca_buscape()`: Faz a busca no Buscapé, filtra os resultados e retorna as ofertas.
  
- **Automação de Envio de E-mail**:
  - Utiliza a biblioteca `win32com.client` para enviar e-mails pelo Microsoft Outlook.

## Resultado

O script gera um arquivo Excel chamado `Ofertas.xlsx`, contendo as colunas:

- **produto**: Nome do produto encontrado.
- **preço**: Preço do produto encontrado.
- **link**: Link para a página do produto.

Se houver ofertas encontradas, o e-mail será enviado automaticamente com essas informações.

## Contribuição

Sinta-se à vontade para enviar sugestões ou melhorias para este projeto.

## Licença

Este projeto está sob a licença MIT. Consulte o arquivo LICENSE para mais informações.
