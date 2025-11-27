Step-by-step Process on Document Extraction

1. Extract text from EA form
2. Take the info from the extracted text and generate CoA
3. Take the info from the extracted text and generate tax comp

Intermediate step
1. Using the output-record.xlsx file (which contains all client information), 
   run generate_tax_comp.py.
   
   The script will process each row and, based on the client’s tax residency status (R or NR), 
   automatically generate a file with the relevant information filled in.

   If a client is eligible for tax relief, the script will also generate 
   an additional personal tax relief file for that client.