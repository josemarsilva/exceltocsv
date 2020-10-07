# exceltocsv

## 1. Introdução ##

Este repositório contém o código fonte do componente **exceltocsv**. Este componente distribuído como um arquivo (.jar) pode ser executado em linha de comando (tanto Windows quanto Linux). O **exceltocsv** exporta o conteúdo de cada uma das abas (sheets) do excel para arquivos no formato (.csv) de mesmo nome da aba .


### 1.1. Help Online linha

* O componente **exceltocsv** funciona com argumentos de linha de comando (tanto Windows quanto Linux). Com o argumento '-h' mostra o help da a aplicação.

```bat
C:\My Git\workspace-github\exceltocsv\dist>java -jar exceltocsv.jar
Missing required options: i, o
usage: exceltocsv [<args-options-list>] - v.2020.10.06.2131
 -h,--help                           shows usage help message. See more
                                     https://github.com/josemarsilva/excel
                                     tocsv
 -i,--input-excel-file <arg>         Nome do arquivo que contem Pasta de
                                     trabalho EXCEL (.xls ou .xlsx) que
                                     sera copiada para CSV. Ex:
                                     exceltocsv-exemplo.xlsx
 -o,--output-folder-path-csv <arg>   Nome do caminho da pasta onde deverao
                                     ser gerados os arquivos (.csv)
                                     correspondentes a cada sheet do
                                     workbook. Ex: . ou C:\TEMP
```




### 2. Documentação ###

### 2.1. Diagrama de Caso de Uso (Use Case Diagram) ###

#### Diagrama de Caso de Uso

![UseCaseDiagram](doc/UseCaseDiagram%20-%20Context%20-%20ExcelToCsv.png)


### 2.5. Requisitos ###

* n/a


## 3. Projeto ##

### 3.1. Pré-requisitos ###

* Linguagem de programação: Java
* IDE: Eclipse (recomendado Oxigen 2)
* JDK: 1.8
* Maven Repository dependence: Apache CLI
* Maven Repository dependence: Apache POI

### 3.2. Guia para Desenvolvimento ###

* Obtenha o código fonte através de um "git clone". Utilize a branch "master" se a branch "develop" não estiver disponível.
* Faça suas alterações, commit e push na branch "develop".


### 3.3. Guia para Configuração ###

* n/a


### 3.4. Guia para Teste ###

* n/a


### 3.5. Guia para Implantação ###

* Obtenha o último pacote (.war) estável gerado disponível na sub-pasta './dist'.



### 3.6. Guia para Demonstração ###

* n/a


### 3.7. Guia para Execução ###

* Exemplo do uso do **exceltocsv**  
    * Suponha um arquivo 'exceltocsv-exemplo.xlsx'.
    * Suponha que você queira exportar todas as abas desta planilha no diretorio corrente

```bat
java -jar exceltocsv.jar 

```


## Referências ##

* n/a
