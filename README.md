# RFID Inventory Processor

A Java 17 application for processing RFID warehouse inventory records exported from Excel. The project transforms raw inventory logs into structured daily reports, automatically tracks RFID lifecycle events, and generates validation logs for inconsistent data.

## Features

* Parse RFID warehouse Excel files using **Apache POI**
* Generate **one worksheet per inventory date**
* Automatically count daily **STOCK / STOCKIN / STOCKOUT**
* Pair **Stock-In** and **Stock-Out** records by RFID ID
* Track **Closed RFID** with both original Stock-In and Stock-Out timestamps
* Generate **Bug_Logs** for invalid teams, categories, stages, status values, and inconsistent inventory records

## Tech Stack

* **Java 17**
* **Apache POI**
* **TreeMap / HashMap**
* **Object-Oriented Design**

## Workflow

```text
RFID Excel
     │
     ▼
 Data Validation
     │
     ▼
 Record Parsing
     │
     ├── inHouse (Open Inventory)
     ├── closed (Completed Lifecycle)
     ├── datCount (Daily Statistics)
     └── Bug Logs
     │
     ▼
 Daily Excel Report
```

## Output

The generated workbook contains:

* **Bug_Logs** worksheet for data validation issues
* **One worksheet per date**
* Daily inventory records
* Daily Status Count summary
* Closed RFID report with Stock-In date and Stock-Out time

## Project Status

**Version:** v2.0

### What's New in v2.0

* Introduced RFID lifecycle tracking (`inHouse` → `closed`)
* Added Stock-In / Stock-Out pairing by RFID ID
* Added Closed RFID reporting with dual timestamps
* Added Bug_Logs worksheet for data quality inspection
* Refactored report generation into multi-sheet daily reports
