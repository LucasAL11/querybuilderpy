# SQL Query Generator for Excel Data

## Overview

This project is a Python script that generates SQL queries to update, rollback, and verify records in an Oracle database table based on data provided in an Excel file. Users can dynamically select which columns to update and specify a reference column to match records in the database.

## Features

- Dynamic SQL query generation for updates using `CASE` statements.
- Rollback functionality using `MERGE` statements.
- Select queries to verify the updated records.
- User-friendly interface for selecting Excel files and columns using PyQt5.

## Requirements

- Python 3.x
- Pandas library for reading Excel files
- PyQt5 library for the GUI components

### Installation

To install the required packages, you can use pip:

```bash
pip install pandas PyQt5
