Matrikon OPC Data Logging Project

This project shows how data from an OPC simulation server can be observed and recorded.
The application works with the Matrikon OPC Simulation Server and collects simulated tag values.
The collected data is stored in CSV files so that it can be viewed later using Excel or any text editor.

The main purpose of this project is to understand how OPC data is accessed and how industrial data logging works in real systems.

Project objectives

The objectives of this project are:

To connect to an OPC simulation server

To read multiple simulated OPC tags

To capture values at fixed time intervals

To store the collected values in CSV files

To create a new CSV file automatically every hour

Tools and technologies used

The following tools are used in this project:

Python programming language

Matrikon OPC Simulation Server

OPC DA communication

CSV file format

Pandas library for writing files

How the project works

First, the OPC client connects to the Matrikon OPC Simulation Server.
Then a set of simulated tags is selected from the server.
The client reads the tag values at regular intervals.
Each reading is written into a CSV file.
When the hour changes, a new CSV file is created automatically.

Output details

The output of the program is stored in CSV files.
Each CSV file contains data for one hour only.
The file name includes the date and hour so that files are easy to identify.

Example file name:
OPC_Log_2026-01-06_16.csv

Each file contains timestamps and tag values recorded during that hour.

Use of this project

This project can be used for:

Academic assignments

Learning OPC basics

Understanding industrial data logging.