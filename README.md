#📦 XML NFe Stock Processor (KG)

A simple Python tool that reads Brazilian NFe XML files, extracts product quantities, and generates Excel spreadsheets for daily movement and cumulative stock in KG.

The script separates XMLs into entry and exit flows, updates stock balances automatically, and keeps track of already processed invoices to avoid double counting. It is designed to run locally with minimal setup.

#🚀 What This Project Does

Reads NFe XML files from folders

Extracts product data from each invoice item

Classifies movements as Entry or Exit

Calculates quantity impact on stock (KG)

Generates a daily movement spreadsheet per run

Maintains an updated cumulative stock spreadsheet

Prevents reprocessing of the same invoice

Moves processed XMLs to archive folders

Logs errors without stopping execution

Displays a completion popup when finished

#🧠 Processing Logic

For each XML file:

Read the NFe key (chNFe)

Skip file if the key was already processed

Iterate through invoice items (det/prod nodes)

Extract:

Product code (cProd)

Product description (xProd)

Commercial quantity (qCom)

Apply movement factor:

Entry = +1

Exit = -1

Record movements

Update cumulative stock totals

Stock formula:

stock = previous_stock + total_entries − total_exits

#📁 Folder Structure

Folders are created automatically when the script runs:

/entrada → Entry XML files
/saida → Exit XML files

/processados/entrada → Processed entry XMLs
/processados/saida → Processed exit XMLs

/resultado → Generated Excel files
/erros → Error logs

#📊 Generated Files

Daily Movement File (created on each execution with timestamp):

resultado/movimento_YYYY-MM-DD_HH-MM-SS.xlsx

Columns:

Date — Processing date

Type — Entry or Exit

Codigo — Product code

Produto — Product name

Quantidade_KG — Quantity impact (signed)

chNFe — Invoice key

Arquivo_XML — Source filename

Cumulative Stock File (continuously updated):

resultado/estoque_geral.xlsx

Columns:

Codigo — Product code

Produto — Product name

Estoque_KG — Current stock balance

Processed Invoice Registry:

resultado/processadas.xlsx

Stores processed NFe keys to ensure idempotent behavior.

#▶️ How to Run

Install dependencies:

pip install pandas openpyxl

Tkinter is typically included with standard Python installations.

Place XML files into the folders:

/entrada
/saida

Run the script:

python main.py

After execution:

Movement spreadsheet is generated

Stock spreadsheet is updated

XMLs are moved to processed folders

Errors (if any) are logged

A completion popup is shown

#⚠️ Error Handling

If an XML cannot be parsed or required fields are missing:

Processing continues for other files

A log file is created in /erros

The log contains filename and exception details

#🔁 Idempotent Behavior

Each invoice is identified by its NFe key (chNFe).

If the key already exists in the processed registry:

The XML is skipped

It is moved to the processed folder

It does not affect stock calculations again

#🛠️ Tech Stack

Python 3

pandas

openpyxl

xml.etree.ElementTree

tkinter