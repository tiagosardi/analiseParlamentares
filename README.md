#Análise de dados da Câmara dos Deputados

Leitura de dados abertos de parlamentares para análise de gastos e proposições de partidos.

Pretendo saber quanto cada partido gastou durante a legislatura 54. Assim como, pretendo saber a quantidade real de proposições feitas pelos partidos.

Para obter dados reais, classifiquei cada parlamentar pelo partido e agrupei os valores que me interessavam no próprio partido.

O programa plota gráficos para auxiliar a compreensão dos dados e retorna um novo arquivo xlsx.
-----------------------------------

### INSTALAÇÃO DAS DEPENDÊNCIAS

Biblioteca para ler o arquivo de formato xlsx

>> sudo pip install xlrd


>> sudo apt-get install python-setuptools

>> sudo easy_install pip

>> sudo pip install virtualenv

//plota graficos melhores (precisa de autenticação)

>> sudo pip install plotly 

>> sudo pip install plotly --upgrade


//instalação do metodo de escrita em arquivos xlsx

>> sudo pip install XlsxWriter

-----------------------------------
### OUTRAS OBSERVAÇÕES

Os dados são abertos e foram extraídos do site da Câmara dos Deputados.

O programa retorna um arquivo xlsx com os resultados em tabela. 

### Resultados que chamam a atenção

. Os parlamentares crescem sua quantidade de proposições no último ano de seus mandatos.

. Os maiores gastos dos parlamentares são com passagens aéreas e divulgações.

. Os menores gastos são com capacitações e pesquisas.