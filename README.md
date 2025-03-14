# Excel Data Transformer

A Python-based web application for transforming Excel data from one format to another specialized format.

## Overview

This application automates the process of transforming Excel files from Dati_imp format into the Gensoft format required by your business systems. It eliminates the need for manual processing through intermediary Excel sheets with complex formulas, reducing errors and saving time.

## Features

- **Simple Web Interface**: Upload files through an intuitive drag-and-drop interface
- **Automated Transformation**: Handles all data extraction, mapping, and calculations
- **Multi-step Processing**: Manages all the complex business rules for data conversion
- **Transformation History**: Keeps track of recent file transformations
- **Error Handling**: Provides clear error messages when issues arise

## How It Works

The transformation process follows these steps:

1. **File Upload**: System accepts Excel files (.xlsx or .xls)
2. **Data Extraction**: Reads data from the source sheet
3. **Code Processing**: Extracts product codes and colors from complex codes
4. **Price Calculations**: Applies VAT (20%) and retail pricing formulas
5. **Category Mapping**: Maps product divisions to proper categories
6. **Formatting**: Structures data according to Gensoft requirements
7. **Header Creation**: Creates exact two-row header format required by target system
8. **Output**: Generates a ready-to-use Excel file in the required format

## Installation

### Prerequisites

- Python 3.7+
- pip

### Setup

1. Clone the repository:
   ```
   git clone https://github.com/yourusername/excel-transformer.git
   cd excel-transformer
   ```

2. Install the required packages:
   ```
   pip install -r requirements.txt
   ```

3. Run the application:
   ```
   python app.py
   ```

4. Open your browser and navigate to:
   ```
   http://localhost:5000
   ```

## Usage

1. Click the "Browse Files" button or drag and drop your Excel file onto the designated area
2. Click "Transform File" to start the transformation process
3. Once complete, the application will provide a download link for the transformed file
4. Previous transformations will be listed in the "Recent Transformations" section

## Technical Details

The transformer performs several specific transformations:

- Extracting product codes using text_before and text_after functions
- Mapping season codes to their quarter/year format (e.g., 251 → Q1-25)
- Calculating prices with proper VAT (20%) and retail pricing formulas
- Rounding retail prices according to business rules (ceil to nearest 10, then subtract 1)
- Setting proper product descriptions based on product categories
- Formatting headers on two rows as required for the Gensoft system

## Requirements

The application relies on the following Python packages:

- Flask (web framework)
- pandas (data manipulation)
- openpyxl (Excel file handling)
- werkzeug (file upload handling)

A complete list is available in the `requirements.txt` file.

## Contributing

Contributions are welcome! Please feel free to submit a Pull Request.

## License

This project is licensed under the MIT License - see the LICENSE file for details.
