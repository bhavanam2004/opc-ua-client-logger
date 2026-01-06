Matrikon OPC Data Logging Project

This project shows how data from an OPC simulation server can be read and saved.
The application connects to the Matrikon OPC Simulation Server and reads simulated tag values.
The data is stored in CSV files so it can be opened later using Excel.
Tools used
* Python
* Matrikon OPC Simulation Server
* OPC DA
* CSV file format
How it works
The OPC client connects to the server and selects simulated tags.
Tag values are read at regular intervals and written to a CSV file.
When the hour changes, a new CSV file is created.
Output
Each CSV file contains one hour of data.
The file name includes the date and hour for easy identification.
Example:
OPC_Log_2026-01-06_16.csv
