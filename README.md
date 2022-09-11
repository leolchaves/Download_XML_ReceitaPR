# Download_XML_ReceitaPR
Faz o download dos arquivos XML solicitados na ReceitaPR

Print de como é o padrão da planilha que precisa estar na mesma pasta desse código para funcionar: ![image](https://user-images.githubusercontent.com/106564039/189552909-8ef272b7-a2b9-466b-a9fe-5d1051d85cbe.png)

O código utiliza selenium, pandas, time, Path, datetime, pyunpack e OS para fazer o download e salvar na pasta correta o XML baixado da Receita PR.
O selenium automatiza o processo de fazer o download no navegador Google Chrome.
O pandas faz a leitura do arquivo Excel com as informações da empresa e do código da empresa no sistema, assim como se ela possui filial colocando a informação de filial após o ponto em "Nome pasta".
O módulo time é utilizado para que o código espere o download dos arquivos acontecer, através do sleep.
O módulo Path é utilizado para passar o caminho do download.
O módulo datetime é utilizado para criar a pasta com o mês referente ao download.
O pyunpack é utilizado para descompactar os arquivos que vem do download da Receita PR e ficar somente com os arquivos XML.
E por fim o módulo "os" é utilizado para excluir os arquivos baixados na pasta temporária de download.
