import fitz  # PyMuPDF
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox
from PIL import Image, ImageTk
import math
import json
from pathlib import Path
from typing import Optional, Tuple
import logging

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)


class InvoiceGeneratorConfig:
    """Handles configuration loading and path management"""
    
    def __init__(self, config_path: Path = None):
        if config_path is None:
            config_path = Path(__file__).parent / "config.json"
        
        self.config_path = config_path
        self.config = self._load_config()
        self.base_dir = Path(__file__).parent
        
    def _load_config(self) -> dict:
        """Load configuration from JSON file"""
        try:
            with open(self.config_path, 'r') as f:
                return json.load(f)
        except FileNotFoundError:
            logger.error(f"Config file not found: {self.config_path}")
            raise
        except json.JSONDecodeError as e:
            logger.error(f"Invalid JSON in config file: {e}")
            raise
    
    def get_template_path(self, template_name: str) -> Path:
        """Get path to a template file"""
        return self.base_dir / self.config['paths']['templates'] / template_name
    
    def get_output_path(self, subfolder: str = '') -> Path:
        """Get output path, creating directory if needed"""
        output_path = self.base_dir / self.config['paths']['output'] / subfolder
        output_path.mkdir(parents=True, exist_ok=True)
        return output_path
    
    def get_position(self, key: str) -> Tuple[int, int]:
        """Get text position from config"""
        return tuple(self.config['text_positions'][key])


class DataFrameManager:
    """Manages Excel file loading and processing"""
    
    def __init__(self):
        self.invoice_data: Optional[pd.DataFrame] = None
        self.backup_data: Optional[pd.DataFrame] = None
    
    def load_invoice_data(self, file_path: str) -> bool:
        """Load invoice input Excel file"""
        try:
            self.invoice_data = pd.read_excel(file_path)
            logger.info(f"Loaded invoice data: {len(self.invoice_data)} rows")
            return True
        except Exception as e:
            logger.error(f"Error loading invoice data: {e}")
            messagebox.showerror("Error", f"Failed to load invoice data:\n{e}")
            return False
    
    def load_backup_data(self, file_path: str) -> bool:
        """Load backup Excel file"""
        try:
            self.backup_data = pd.read_excel(file_path)
            logger.info(f"Loaded backup data: {len(self.backup_data)} rows")
            return True
        except Exception as e:
            logger.error(f"Error loading backup data: {e}")
            messagebox.showerror("Error", f"Failed to load backup data:\n{e}")
            return False
    
    def is_ready(self) -> bool:
        """Check if both datasets are loaded"""
        return self.invoice_data is not None and self.backup_data is not None


class PDFGenerator:
    """Handles PDF generation and manipulation"""
    
    def __init__(self, config: InvoiceGeneratorConfig):
        self.config = config
        self.positions = {key: config.get_position(key) for key in config.config['text_positions']}
    
    def merge_pdfs_in_folders(self, folder_path1: Path, folder_path2: Path, output_pdf: Path):
        """Merge all PDFs from two folders into a single PDF"""
        def load_pdf_files(folder_path: Path):
            return sorted([f for f in folder_path.glob("*.pdf")])
        
        try:
            pdf_files1 = load_pdf_files(folder_path1)
            pdf_files2 = load_pdf_files(folder_path2)
            
            output_doc = fitz.open()
            
            for pdf_files in [pdf_files1, pdf_files2]:
                for pdf_file in pdf_files:
                    input_doc = fitz.open(pdf_file)
                    output_doc.insert_pdf(input_doc)
                    input_doc.close()
            
            output_doc.save(output_pdf)
            output_doc.close()
            logger.info(f"Merged PDF created: {output_pdf}")
        except Exception as e:
            logger.error(f"Error merging PDFs: {e}")
            raise
    
    def delete_pdf_files_in_folders(self, *folder_paths: Path):
        """Delete all PDF files in specified folders"""
        for folder_path in folder_paths:
            try:
                for pdf_file in folder_path.glob("*.pdf"):
                    pdf_file.unlink()
                logger.info(f"Cleaned up PDFs in: {folder_path}")
            except Exception as e:
                logger.error(f"Error deleting PDFs in {folder_path}: {e}")
    
    def format_accounting_month(self, date) -> str:
        """Format date as accounting month (e.g., 'Jan-25')"""
        return date.strftime('%b-%y')
    
    def calculate_vat_position(self, vat_amount: float) -> str:
        """Determine which VAT position to use based on amount"""
        vat_middle = round(vat_amount, 2)
        vat_string = str(vat_middle)
        parts = vat_string.split(".")
        integ_part = parts[0]
        integ_num = len(integ_part)
        
        if integ_num == 5:
            return "vat_dos"
        elif 1000 < vat_middle < 2000:
            return "vat_dos"
        else:
            return "vat"
    
    def create_front_page(self, row: pd.Series, output_path: Path):
        """Create the front page of the invoice"""
        template_path = self.config.get_template_path("front_pager.pdf")
        
        try:
            pdf_document = fitz.open(template_path)
            page = pdf_document[0]
            
            # Format data
            invoice_date_str = row['Invoice Date'].strftime('%d/%m/%Y')
            due_date_str = row['Due Date'].strftime('%d/%m/%Y')
            acc_month_str = self.format_accounting_month(row['Line Description'])
            
            quantity_middle = round(row['Invoice Amount'] / 1000, 5)
            quantity_final = format(quantity_middle, '.5f').rstrip('0')
            
            net_amount_str = f"£{row['Invoice Amount']:.2f}"
            sub_amount_str = f"{row['Invoice Amount']:.2f}"
            vat_amount_str = f"{row['VAT Amount']:.2f}"
            total_amount_str = f"{row['Total']:.2f}"
            
            # Insert text at positions
            page.insert_text(self.positions["invoice_reference"], str(row['Invoice Number']), fontname="helv", fontsize=8)
            page.insert_text(self.positions["invoice_date"], invoice_date_str, fontname="helv", fontsize=8)
            page.insert_text(self.positions["due_date"], due_date_str, fontname="helv", fontsize=8)
            page.insert_text(self.positions["po"], str(row['PO']), fontname="helv", fontsize=8)
            page.insert_text(self.positions["accounting_month_uno"], acc_month_str, fontname="helv", fontsize=7)
            page.insert_text(self.positions["accounting_month_dos"], acc_month_str, fontname="helv", fontsize=7)
            page.insert_text(self.positions["accounting_month_tres"], acc_month_str, fontname="helv", fontsize=7)
            
            # Conditional quantity formatting
            quantity_fontsize = 6.4 if len(quantity_final) > 7 else 7
            page.insert_text(self.positions["quantity"], quantity_final, fontname="helv", fontsize=quantity_fontsize)
            
            page.insert_text(self.positions["net_amount"], net_amount_str, fontname="helv", fontsize=7)
            page.insert_text(self.positions["sub_total"], sub_amount_str, fontname="helv", fontsize=8)
            
            # Conditional VAT position
            vat_position_key = self.calculate_vat_position(row['VAT Amount'])
            page.insert_text(self.positions[vat_position_key], vat_amount_str, fontname="helv", fontsize=8)
            
            page.insert_text(self.positions["total"], total_amount_str, fontname="helv", fontsize=8)
            
            pdf_document.save(output_path)
            pdf_document.close()
            logger.info(f"Created front page: {output_path}")
            
            return net_amount_str
        except Exception as e:
            logger.error(f"Error creating front page: {e}")
            raise
    
    def create_backup_pages(self, dataframes: list, invoice_number: str, net_amount_str: str):
        """Create backup pages for the invoice"""
        template_path = self.config.get_template_path("blank_template.pdf")
        backup_folder = self.config.get_output_path("back_up")
        
        max_char_limit = 30
        max_char_limit_two = 20
        
        try:
            for i, df_page in enumerate(dataframes):
                pdf_document = fitz.open(template_path)
                page = pdf_document[0]
                
                # Select and format columns
                quote_ref = df_page[['Supplier Quote ref.']].copy()
                client_ref = df_page[['Client Ref']].copy()
                site_name = df_page[['Site Name ']].copy()
                reviewed_quote = df_page[['Reviewed Quote/Estimate (£)']].copy()
                
                # Format currency
                reviewed_quote['Reviewed Quote/Estimate (£)'] = reviewed_quote['Reviewed Quote/Estimate (£)'].apply(
                    lambda x: f'£{x:,.2f}'
                )
                
                # Truncate and pad strings
                site_name['Site Name '] = site_name['Site Name '].str.slice(0, max_char_limit).str.ljust(max_char_limit)
                client_ref['Client Ref'] = client_ref['Client Ref'].str.slice(0, max_char_limit_two).str.ljust(max_char_limit_two)
                
                # Convert to string format
                text_quote_ref = quote_ref.to_string(index=False, header=False)
                text_client_ref = client_ref.to_string(index=False, header=False)
                text_site_name = site_name.to_string(index=False, header=False)
                text_reviewed_quote = reviewed_quote.to_string(index=False, header=False)
                
                # Insert text
                page.insert_text(self.positions["bloque_uno"], text_quote_ref, fontname="helv", fontsize=8)
                page.insert_text(self.positions["bloque_two"], text_client_ref, fontname="helv", fontsize=8)
                page.insert_text(self.positions["bloque_three"], text_site_name, fontname="helv", fontsize=8)
                page.insert_text(self.positions["bloque_four"], text_reviewed_quote, fontname="helv", fontsize=8)
                
                # Add total to last page
                if i == len(dataframes) - 1:
                    page.insert_text(self.positions["total_two"], net_amount_str, fontname="helv", fontsize=8)
                    output_path = backup_folder / f"{invoice_number} - 999.pdf"
                else:
                    output_path = backup_folder / f"{invoice_number} - {i}.pdf"
                
                pdf_document.save(output_path)
                pdf_document.close()
                logger.info(f"Created backup page: {output_path}")
        except Exception as e:
            logger.error(f"Error creating backup pages: {e}")
            raise


class InvoiceProcessor:
    """Main processor for invoice generation"""
    
    def __init__(self, config: InvoiceGeneratorConfig, data_manager: DataFrameManager):
        self.config = config
        self.data_manager = data_manager
        self.pdf_generator = PDFGenerator(config)
        self.qt_data = pd.DataFrame(columns=['Supplier Quote ref.', 'Invoice Number'])
    
    def split_dataframe(self, df: pd.DataFrame, rows_per_page: int = 58) -> list:
        """Split dataframe into pages"""
        dataframes = []
        current_row = 0
        
        while current_row < len(df):
            page_df = df.iloc[current_row:current_row + rows_per_page]
            dataframes.append(page_df)
            current_row += rows_per_page
        
        return dataframes
    
    def process_single_invoice(self, row: pd.Series):
        """Process a single invoice"""
        try:
            # Prepare backup data
            backup_df = self.data_manager.backup_data.copy()
            backup_df['Financial Month'] = pd.to_datetime(backup_df['Financial Month'], format='%b-%y')
            backup_df['Financial Month'] = backup_df['Financial Month'].apply(self.pdf_generator.format_accounting_month)
            
            acc_month_str = self.pdf_generator.format_accounting_month(row['Line Description'])
            
            # Filter backup data
            condition1 = backup_df['Financial Month'] == acc_month_str
            condition2 = backup_df['PO Order No.'] == row['PO']
            filtered_df = backup_df[condition1 & condition2]
            
            selected_columns = ['Supplier Quote ref.', 'Client Ref', 'Site Name ', 'Reviewed Quote/Estimate (£)']
            filtered_data = filtered_df[selected_columns]
            
            # Update QT data
            qt_temp = pd.DataFrame({
                'Supplier Quote ref.': filtered_data['Supplier Quote ref.'],
                'Invoice Number': row['Invoice Number']
            })
            self.qt_data = pd.concat([self.qt_data, qt_temp], ignore_index=True)
            
            # Split into pages
            page_dataframes = self.split_dataframe(filtered_data)
            
            # Create front page
            one_pager_folder = self.config.get_output_path("one_pager")
            front_page_path = one_pager_folder / f"{row['Invoice Number']}.pdf"
            net_amount_str = self.pdf_generator.create_front_page(row, front_page_path)
            
            # Create backup pages
            self.pdf_generator.create_backup_pages(page_dataframes, row['Invoice Number'], net_amount_str)
            
            # Merge PDFs
            final_output = self.config.get_output_path() / f"{row['Invoice Number']}.pdf"
            backup_folder = self.config.get_output_path("back_up")
            
            self.pdf_generator.merge_pdfs_in_folders(one_pager_folder, backup_folder, final_output)
            
            # Clean up temporary files
            self.pdf_generator.delete_pdf_files_in_folders(one_pager_folder, backup_folder)
            
            logger.info(f"Successfully processed invoice: {row['Invoice Number']}")
        except Exception as e:
            logger.error(f"Error processing invoice {row['Invoice Number']}: {e}")
            raise
    
    def process_all_invoices(self):
        """Process all invoices in the dataset"""
        if not self.data_manager.is_ready():
            messagebox.showerror("Error", "Please load both input and backup Excel files first.")
            return
        
        try:
            total_invoices = len(self.data_manager.invoice_data)
            logger.info(f"Starting to process {total_invoices} invoices")
            
            for index, row in self.data_manager.invoice_data.iterrows():
                logger.info(f"Processing invoice {index + 1}/{total_invoices}")
                self.process_single_invoice(row)
            
            # Save QT fillable data
            qt_output_path = self.config.get_output_path() / "QT_Fillable_data.xlsx"
            self.qt_data.to_excel(qt_output_path, index=False)
            logger.info(f"QT fillable data saved to: {qt_output_path}")
            
            return True
        except Exception as e:
            logger.error(f"Error during invoice processing: {e}")
            messagebox.showerror("Error", f"Processing failed:\n{e}")
            return False


class InvoiceGeneratorGUI:
    """GUI for the invoice generator"""
    
    def __init__(self):
        self.config = InvoiceGeneratorConfig()
        self.data_manager = DataFrameManager()
        self.processor = InvoiceProcessor(self.config, self.data_manager)
        
        self.root = tk.Tk()
        self.root.title("Invoice Generator v2.0")
        self.root.geometry("500x400")
        
        self.setup_gui()
    
    def setup_gui(self):
        """Setup the GUI elements"""
        # Logo
        try:
            logo_path = self.config.get_template_path("applogo.png")
            logo_image = Image.open(logo_path)
            logo_photo = ImageTk.PhotoImage(logo_image)
            logo_label = tk.Label(self.root, image=logo_photo)
            logo_label.image = logo_photo  # Keep a reference
            logo_label.pack(pady=10)
        except Exception as e:
            logger.warning(f"Could not load logo: {e}")
        
        # Instructions
        instruction_label = tk.Label(
            self.root, 
            text="Invoice Generator",
            font=("Arial", 14, "bold")
        )
        instruction_label.pack(pady=10)
        
        # Status labels
        self.invoice_status = tk.Label(self.root, text="Invoice Data: Not loaded", fg="red")
        self.invoice_status.pack(pady=5)
        
        self.backup_status = tk.Label(self.root, text="Backup Data: Not loaded", fg="red")
        self.backup_status.pack(pady=5)
        
        # Buttons
        btn_frame = tk.Frame(self.root)
        btn_frame.pack(pady=20)
        
        load_invoice_btn = tk.Button(
            btn_frame,
            text="Load Invoice Input File",
            command=self.load_invoice_data,
            width=25,
            height=2
        )
        load_invoice_btn.pack(pady=5)
        
        load_backup_btn = tk.Button(
            btn_frame,
            text="Load Backup File",
            command=self.load_backup_data,
            width=25,
            height=2
        )
        load_backup_btn.pack(pady=5)
        
        process_btn = tk.Button(
            btn_frame,
            text="Generate Invoices",
            command=self.generate_invoices,
            width=25,
            height=2,
            bg="green",
            fg="white",
            font=("Arial", 10, "bold")
        )
        process_btn.pack(pady=15)
    
    def load_invoice_data(self):
        """Load invoice input Excel file"""
        file_path = filedialog.askopenfilename(
            title="Select Invoice Input File",
            filetypes=[("Excel Files", "*.xlsx"), ("All Files", "*.*")]
        )
        if file_path:
            if self.data_manager.load_invoice_data(file_path):
                self.invoice_status.config(text="Invoice Data: Loaded ✓", fg="green")
    
    def load_backup_data(self):
        """Load backup Excel file"""
        file_path = filedialog.askopenfilename(
            title="Select Backup File",
            filetypes=[("Excel Files", "*.xlsx"), ("All Files", "*.*")]
        )
        if file_path:
            if self.data_manager.load_backup_data(file_path):
                self.backup_status.config(text="Backup Data: Loaded ✓", fg="green")
    
    def generate_invoices(self):
        """Generate all invoices"""
        if not self.data_manager.is_ready():
            messagebox.showwarning(
                "Missing Data",
                "Please load both the Invoice Input and Backup files before generating invoices."
            )
            return
        
        # Confirm before processing
        result = messagebox.askyesno(
            "Confirm",
            f"Ready to generate {len(self.data_manager.invoice_data)} invoices. Continue?"
        )
        
        if result:
            # Disable buttons during processing
            for widget in self.root.winfo_children():
                if isinstance(widget, tk.Button):
                    widget.config(state='disabled')
            
            self.root.update()
            
            # Process invoices
            success = self.processor.process_all_invoices()
            
            # Re-enable buttons
            for widget in self.root.winfo_children():
                if isinstance(widget, tk.Button):
                    widget.config(state='normal')
            
            if success:
                messagebox.showinfo(
                    "Success",
                    "All invoices have been generated successfully!\n\n"
                    f"Output location: {self.config.get_output_path()}"
                )
    
    def run(self):
        """Start the GUI"""
        self.root.mainloop()


def main():
    """Main entry point"""
    try:
        app = InvoiceGeneratorGUI()
        app.run()
    except Exception as e:
        logger.error(f"Application error: {e}")
        messagebox.showerror("Fatal Error", f"Application failed to start:\n{e}")


if __name__ == "__main__":
    main()