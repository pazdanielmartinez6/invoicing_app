# Invoice Generator

An automated PDF invoice generator that combines invoice data with backup information to create professional, multi-page invoices.

## Features

- 🚀 Automated invoice PDF generation
- 📊 Excel data integration
- 🖼️ GUI for easy file selection
- 📄 Multi-page invoice support with backup data
- 🔧 Cross-platform compatibility
- ⚙️ Configurable via JSON

## Project Structure

```
invoice_generator/
├── invoice_generator.py    # Main application
├── config.json             # Configuration file
├── requirements.txt        # Python dependencies
├── templates/              # PDF and image templates
│   ├── front_pager.pdf
│   ├── blank_template.pdf
│   └── applogo.png
├── output/                 # Generated invoices (created automatically)
│   ├── one_pager/         # Temporary front pages
│   ├── back_up/           # Temporary backup pages
│   └── [Invoice PDFs]     # Final merged invoices
└── README.md              # This file
```

## Installation

### Prerequisites

- Python 3.8 or higher
- pip (Python package installer)

### Setup Steps

1. **Clone or download this repository**
   ```bash
   git clone <your-repo-url>
   cd invoice_generator
   ```

2. **Install required packages**
   ```bash
   pip install -r requirements.txt
   ```

3. **Prepare template files**
   - Place your PDF templates in the `templates/` folder:
     - `front_pager.pdf` - Invoice front page template
     - `blank_template.pdf` - Backup page template
     - `applogo.png` - Application logo (optional)

4. **Verify configuration**
   - Review `config.json` and adjust text positions if needed

## Usage

### Running the Application

```bash
python invoice_generator.py
```

### Step-by-Step Process

1. **Launch the application**
   - The GUI window will open

2. **Load Invoice Input File**
   - Click "Load Invoice Input File"
   - Select your Excel file with invoice data
   - Required columns:
     - Invoice Number
     - Invoice Date
     - Due Date
     - PO
     - Line Description
     - Invoice Amount
     - VAT Amount
     - Total

3. **Load Backup File**
   - Click "Load Backup File"
   - Select your Excel file with backup data
   - Required columns:
     - Date Quote Sent to E///
     - Supplier Quote ref.
     - Client Ref
     - Financial Month
     - Site Name 
     - Reviewed Quote/Estimate (£)
     - PO Order No.

4. **Generate Invoices**
   - Click "Generate Invoices"
   - Confirm the operation
   - Wait for processing to complete
   - Find your invoices in the `output/` folder

### Output Files

The application generates:
- **Individual invoice PDFs** in the `output/` folder
- **QT_Fillable_data.xlsx** - Reference data for QT updates

## Configuration

Edit `config.json` to customize:

### Paths
```json
"paths": {
  "templates": "templates",
  "output": "output"
}
```

### Text Positions
Adjust X,Y coordinates for text placement on PDFs:
```json
"text_positions": {
  "invoice_reference": [430, 126],
  "invoice_date": [430, 137],
  ...
}
```

### PDF Settings
```json
"pdf_settings": {
  "rows_per_backup_page": 58,
  "max_site_name_chars": 30,
  "max_client_ref_chars": 20
}
```

## Excel File Format

### Invoice Input File

| Invoice Number | Invoice Date | Due Date | PO | Line Description | Invoice Amount | VAT Amount | Total |
|---------------|--------------|----------|-----|------------------|----------------|------------|-------|
| INV001 | 2024-01-15 | 2024-02-15 | PO123 | Jan-24 | 10000.00 | 2000.00 | 12000.00 |

### Backup File

| Date Quote Sent | Supplier Quote ref. | Client Ref | Financial Month | Site Name | Reviewed Quote/Estimate (£) | PO Order No. |
|----------------|---------------------|------------|-----------------|-----------|----------------------------|--------------|
| 2024-01-10 | QT001 | CLT001 | Jan-24 | Site A | 5000.00 | PO123 |

## Troubleshooting

### Common Issues

**Issue: "Config file not found"**
- Ensure `config.json` is in the same directory as `invoice_generator.py`

**Issue: "Template file not found"**
- Check that all template files are in the `templates/` folder
- Verify file names match configuration

**Issue: "Failed to load Excel file"**
- Ensure the Excel file is not open in another program
- Verify column names match expected format
- Check for correct date formats in Excel

**Issue: Text appears in wrong position**
- Adjust coordinates in `config.json` under `text_positions`
- Use a PDF editing tool to find correct X,Y coordinates

### Logging

The application logs all operations. Check the console output for detailed information about:
- File loading status
- Processing progress
- Error messages

## Advanced Usage

### Customizing Text Positions

To adjust where text appears on your PDFs:

1. Open your PDF template in a PDF editor that shows coordinates
2. Note the X,Y position where you want text to appear
3. Update the corresponding position in `config.json`
4. Test with a single invoice

### Modifying Page Layout

To change how many rows appear per backup page:
```json
"pdf_settings": {
  "rows_per_backup_page": 58  // Adjust this number
}
```

## Development

### Code Structure

- **InvoiceGeneratorConfig**: Handles configuration and paths
- **DataFrameManager**: Manages Excel file loading
- **PDFGenerator**: Creates and manipulates PDFs
- **InvoiceProcessor**: Main business logic
- **InvoiceGeneratorGUI**: User interface

### Contributing

1. Fork the repository
2. Create a feature branch
3. Make your changes
4. Test thoroughly
5. Submit a pull request

## Requirements

- Python 3.8+
- PyMuPDF (fitz) - PDF manipulation
- pandas - Excel data handling
- Pillow - Image processing
- openpyxl - Excel file reading
- tkinter - GUI (usually included with Python)

## License

[Add your license here]

## Support

For issues or questions:
- Check the troubleshooting section
- Review log output in the console
- Create an issue in the repository

## Version History

### v2.0
- Refactored for cross-platform compatibility
- Added configuration file
- Improved error handling
- Enhanced logging
- Better code organization

### v1.0
- Initial release
- Basic invoice generation functionality