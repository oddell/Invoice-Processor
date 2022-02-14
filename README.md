# Procurement-Invoice-Processor

<img align="right" width="47%" src="https://github.com/oddell/Invoice-Processor/blob/main/Images/Cost%20Time%20Table.PNG?raw=true" /> 

This program recieves PDF's of invoices and using Procurement Document AI scans and organises the information into a database for a cost value report.

The invoices are automatically archived by date and project number. Their path is linked to their invoice ID in the excel for ease of access.

A global invoice reconcilliation is also generated for contract managers overview.

A graphical interface for manually entering invoices with information that can not be found is included. 


## Motivation

Identified that the system (data entry, filing & error checking) could be automated, and that it would be economically advantageous.

Researched, designed, and developed a solution using Python and Google Cloud Platform (GCP).

Implemented the system handling 2000 invoices per month reducing costs by 75% and data entry workload by 95%.

## Key Features

- GCP Procurement [Doc AI](https://cloud.google.com/document-ai)
- Troubleshooting Data
- Manual GUI for Missing Data
- Automatic Global and Project Specific Invoice Reconcilliation Excel
- Automatic Archiving

