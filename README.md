#  Automacao de Processo de Vendas  

##  Descrição  

Este script foi desenvolvido para automatizar o processamento de vendas, organizando os dados por loja e gerando backups diários. Ele facilita a gestão dos dados ao consolidar informações de diferentes fontes e organizá-las de forma estruturada.  

##  Funcionalidades  

- 🔄 **Carregamento de Dados**: Importa arquivos de emails, lojas e vendas.  
- 📊 **Organização e Processamento**: Mescla os dados das lojas com as vendas, estruturando-os em um dicionário por loja.  
- 📂 **Geração de Backup**: Cria backups automáticos organizados por loja em pastas separadas.  
- 📧 **Envio de Relatórios**: Possibilidade de envio de emails com os relatórios de vendas.  

##  Requisitos  

- Python **3.x**  
- Bibliotecas necessárias (instale com o comando abaixo):  

  ```bash
  pip install pandas pywin32 pathlib
  ```

##  Como Usar  

1. Certifique-se de que os arquivos de entrada estão na pasta correta (`Bases de Dados`):  
   - 📄 `Emails.xlsx`  
   - 📄 `Lojas.csv`  
   - 📄 `Vendas.xlsx`  
2. Execute o script via terminal ou prompt de comando:  

   ```bash
   python Automacao_de_Processo.py
   ```

3. (Opcional) Para execução interativa, use o Jupyter Notebook:  

   ```bash
   jupyter notebook Automacao_de_Processo.ipynb
   ```

##  Estrutura de Arquivos  

```
📂 Projeto
 ├ 📂 Bases de Dados
 │ ├ 📄 Emails.xlsx
 │ ├ 📄 Lojas.csv
 │ ├ 📄 Vendas.xlsx
 ├ 📂 Backup Arquivos Lojas
 │ └ 📂 [Nome das Lojas] -> Arquivos de backup por loja
 ├ 📄 Automacao_de_Processo.py
 ├ 📄 Automacao_de_Processo.ipynb
 └ 📄 README.md
```

##  Exemplo de Saída  

Após a execução, o código irá reconhecer cada loja dentro dos arquivos e verificar se já existe uma pasta no diretório referente a tal loja. Caso não exista, o próprio código irá criar a pasta. Em seguida, ele verificará o resultado do dia anterior daquela loja e salvará os dados em um arquivo .xlsx. O sistema então encaminhará um email sobre a loja, contendo:

   - Uma tabela com os resultados gerais da loja.

   - Indicação se os resultados estão dentro da meta estabelecida na lógica do programa.

   - O arquivo .xlsx da loja, com todos os detalhes das vendas.

```
📂 Backup Arquivos Lojas
 ├ 📂 Loja_A
 │ ├ 📄 vendas_Loja_A_2025-03-22.xlsx
 │ └ 📄 resumo_vendas_Loja_A.txt
 ├ 📂 Loja_B
 │ ├ 📄 vendas_Loja_B_2025-03-22.xlsx
 │ └ 📄 resumo_vendas_Loja_B.txt
```

##   Envio de Relatórios  

Após o processamento dos dados, um email será enviado automaticamente para cada loja, contendo:

   - Resumo dos resultados em uma tabela.

   - Status da meta, indicando se a loja atingiu ou não os objetivos estabelecidos.

   - Arquivo .xlsx anexo, com todos os detalhes das vendas do dia anterior.

Para o envio de emails, este script utiliza a biblioteca win32com.client para interagir com o Microsoft Outlook. Para que o envio dos emails ocorra corretamente, é necessário:

   - Ter o Microsoft Outlook 2016 (ou versões compatíveis) instalado.

   - Configurar previamente uma conta de email no Outlook.

   - Certificar-se de que o Outlook está aberto e funcionando no momento da execução do script.

##  Autor  

 **Desenvolvido por Luiz da Silva Oliveira (ZeedMcFly)**  
Automatização para otimizar processos de vendas e gestão de dados.  

