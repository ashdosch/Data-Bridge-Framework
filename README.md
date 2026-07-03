# Data Bridge Framework

A configuration-driven data integration framework for importing, transforming, validating, and synchronizing structured data from multiple external sources into a canonical SQL database.

This project was designed to eliminate vendor-specific import scripts by providing a reusable ingestion engine driven entirely through configuration files.

---

## Overview

Many organizations receive data from dozens of suppliers, each with different spreadsheet formats, naming conventions, and data quality.

Instead of writing a custom parser for every vendor, the Data Bridge Framework allows new data sources to be onboarded by simply creating a YAML mapping file.

The framework handles:

- Excel ingestion
- Schema normalization
- Data type conversion
- Category classification
- SQL Server synchronization
- Logging
- Validation
- Error handling

without requiring modifications to the core ingestion engine.

---

## Key Features

### Configuration-Driven Imports

Vendor-specific mappings are stored in YAML.

Instead of changing Python code, new vendors are added through configuration.

```text
Vendor Excel File
        │
        ▼
Vendor Mapping (YAML)
        │
        ▼
Canonical Product Model
        │
        ▼
SQL Server
```

---

### Automatic Column Mapping

Different vendors may refer to the same field differently.

Example:

```
Fixture Width
Width
Overall Width
Fixture_Width
```

All become

```
width_in
```

using configurable mapping rules.

---

### Layered Category Detection

Products can automatically be classified using multiple strategies:

- Vendor supplied categories
- Regex matching
- Feature detection
- Default fallback rules

Example:

```
"52 Inch Smart Ceiling Fan"
```

↓

```
Category = Fans
```

---

### Intelligent Data Conversion

The framework automatically converts values into canonical types.

Examples include:

| Input | Output |
|--------|---------|
| `$1,249.99` | Decimal |
| `22W` | Integer |
| `Yes` | Boolean |
| `120 - 277` | Numeric |
| blank | NULL |

---

### SQL Server Integration

Supports:

- staging
- inserts
- updates
- merge operations
- dynamic schema creation
- header normalization

---

### Error Handling

Rather than terminating the import, the framework:

- logs row-level conversion issues
- logs file-level failures
- archives processed files
- moves invalid files into an error directory

allowing large imports to continue processing.

---

## Repository Structure

```
Data-Bridge-Framework/

├── ingest_vendors.py
│     Main ingestion engine
│
├── excel-auto-import.py
│     Dynamic Excel import utility
│
├── vendor_mappings/
│     Vendor configuration files
│
├── vendor_mappings_all.yaml
│     Master vendor configuration
│
├── ingest_data.bat
├── ingest_data_verbose.bat
│
└── README.md
```

---

## Technologies

- Python
- Pandas
- PyYAML
- SQL Server
- pyodbc
- Excel
- YAML

---

## Architecture

```
Excel Files
      │
      ▼

Vendor Detection

      │
      ▼

Load YAML Mapping

      │
      ▼

Normalize Headers

      │
      ▼

Category Detection

      │
      ▼

Data Type Conversion

      │
      ▼

Validation

      │
      ▼

SQL Server Upsert

      │
      ▼

Logging + Archive
```

---

## Why YAML?

The goal was to separate business rules from application logic. This transfers the management of python scripts/code to the management of YAML files. 
In my opinion, YAML is easy to read and amend. Not that python is difficult to interpret, but there is a case to be made against editing code vs. configuration files.

Adding support for a new supplier typically requires:

- creating a new YAML mapping
- defining column aliases
- defining conversions
- defining category rules

without modifying the ingestion engine itself.

This significantly reduces maintenance costs and improves long-term scalability.

You have the option to also test an initial import using the default rules.

---

## Typical Use Cases

Although originally designed for ecommerce product catalogs, the framework can be adapted for:

- Product Information Management (PIM)
- ERP integrations
- CRM synchronization
- Supplier catalog imports
- Legacy system migrations
- Master Data Management
- Data warehouse loading

---

## Current Capabilities

- Excel workbook ingestion (simply due to how the vendors push data)
- Multi-sheet imports
- Vendor auto-detection
- Dynamic column mapping
- Canonical schema transformation
- Automatic category inference
- Data validation
- SQL Server synchronization
- Row-level error logging
- Batch processing

---

## Planned Enhancements

- CSV support
- REST API connectors
- PostgreSQL connector
- MySQL connector
- Incremental synchronization
- Docker deployment
- Unit testing
- Web dashboard
- Pipeline scheduling
- Email notifications

---

## About

This repository demonstrates the architecture behind a reusable enterprise data integration framework developed for managing large-scale vendor product imports.

Sensitive business logic, connection strings, and proprietary mappings have been intentionally removed or generalized.