# ReconFlow

Validation MVP for reconciling invoice and payment spreadsheets.

## Features
- Upload .xlsx or .csv invoice/payment files
- Exact matching by reference + amount
- Fallback matching by amount + date
- Duplicate and unmatched detection
- Net variance summary
- Excel export

## Required columns
Reference-like: Reference, Ref, OrderId, InvoiceId, TransactionId, Id
Amount-like: Amount, TransactionAmount, Total, Value
Date is optional.

## Deployment
Build from this folder with the included Dockerfile. The app listens on Railway's PORT variable.
