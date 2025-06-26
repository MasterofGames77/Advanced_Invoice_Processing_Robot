# 🧾 Advanced Invoice Processing Bot (Blue Prism)

This Blue Prism RPA project automates the extraction, validation, and queuing of invoice data from an Excel spreadsheet using exception handling and work queue management. It’s designed to simulate a scalable and reliable back-office invoice processing system.

## ✅ Key Features

- **Excel Data Import**  
  Automatically reads invoice data from a structured Excel sheet using `MS Excel VBO :: Get Worksheet As Collection`.

- **Dynamic Row Counting**  
  Safely counts and verifies data rows using a loop-based method compatible with all Blue Prism versions.

- **Field Extraction per Row**  
  For each invoice, the bot extracts key fields such as:
  - `Invoice Number`
  - `Amount`

- **Work Queue Integration**  
  Each invoice record is added to a Blue Prism work queue for scalable background processing and later use in downstream systems.

- **Exception Handling**  
  The bot gracefully handles:
  - Empty Excel files
  - Missing fields
  - Invalid or zero-dollar invoices

- **Reusable Modular Design**  
  Built using sub-processes and utility pages (`Read`, `Add to Queue`) for maintainability and testing.

## 🖼️ Screenshots

| Data Flow Overview | Looping & Row Handling |
|--------------------|------------------------|
| ![Main Flow](./screenshots/main-flow.png) | ![Loop](./screenshots/loop-invoices.png) |

> 🔍 Screenshots show key stages: reading Excel, looping through invoices, row validation, and queue submission.

## ⚙ Technologies Used

- **RPA Platform**: Blue Prism
- **Language**: Visual Logic (Blue Prism internal scripting)
- **Input Source**: Excel (.xlsx)
- **Output Target**: Blue Prism Work Queue

## 💡 How It Works

1. **Launches Excel**, opens the specified file, and reads invoice data into a collection.
2. **Loops through each row**, validates the data, and extracts invoice fields.
3. **Adds each invoice** to a work queue for asynchronous processing.
4. **Tracks row count** and handles empty collections without crashing.
5. **Logs or skips errors** using decision and exception stages.

## 📌 Limitations

- Requires Microsoft Excel **desktop** installed (not compatible with Excel Online).
- Blue Prism `.bprelease` files are not included due to GitHub limitations. Only screenshots and documentation are provided.
