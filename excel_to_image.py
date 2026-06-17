import win32com.client
import fitz  # PyMuPDF
import os
import sys

def convert_excel_to_image(excel_path, output_png_path):
    print(f"[excel_to_image] Converting {excel_path} to image...")
    
    excel_path = os.path.abspath(excel_path)
    output_png_path = os.path.abspath(output_png_path)
    
    if not os.path.exists(excel_path):
        raise FileNotFoundError(f"Excel file not found: {excel_path}")
        
    temp_pdf_path = output_png_path.replace(".png", "_temp.pdf")
    if os.path.exists(temp_pdf_path):
        os.remove(temp_pdf_path)
        
    # Start Excel in the background
    excel = None
    wb = None
    try:
        # Initialize COM
        import pythoncom
        pythoncom.CoInitialize()
        
        excel = win32com.client.Dispatch("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        
        # Open the workbook
        wb = excel.Workbooks.Open(excel_path)
        ws = wb.ActiveSheet
        
        # Print Area and Page Fit settings
        ws.PageSetup.PrintArea = "$B$2:$R$45"
        ws.PageSetup.Zoom = False
        ws.PageSetup.FitToPagesWide = 1
        ws.PageSetup.FitToPagesTall = 1
        
        # Export as PDF (0 represents xlTypePDF)
        print(f"[excel_to_image] Exporting temporary PDF to: {temp_pdf_path}")
        ws.ExportAsFixedFormat(0, temp_pdf_path)
        
        # Save and close. This is crucial because Excel will save the evaluated cache
        # of all formulas, so libraries like openpyxl/pandas can read calculated values later.
        wb.Close(SaveChanges=True)
        wb = None
        print(f"[excel_to_image] Excel workbook closed and saved with formula cache.")
    except Exception as e:
        print(f"[excel_to_image] Excel COM automation failed: {e}")
        if wb is not None:
            try:
                wb.Close(SaveChanges=False)
            except:
                pass
        raise e
    finally:
        if excel is not None:
            try:
                excel.Quit()
            except:
                pass
                
    # Convert PDF to PNG via PyMuPDF (fitz) at 300 DPI
    try:
        if not os.path.exists(temp_pdf_path):
            raise FileNotFoundError(f"Temporary PDF file was not created: {temp_pdf_path}")
            
        print(f"[excel_to_image] Rendering first page of PDF at 300 DPI...")
        doc = fitz.open(temp_pdf_path)
        page = doc[0]  # First page
        
        # 300 DPI zoom factor is 300 / 72 = 4.16666667
        zoom = 300 / 72
        mat = fitz.Matrix(zoom, zoom)
        pix = page.get_pixmap(matrix=mat)
        
        # Save image
        pix.save(output_png_path)
        doc.close()
        print(f"[excel_to_image] Image successfully generated: {output_png_path}")
    except Exception as e:
        print(f"[excel_to_image] PDF to PNG conversion failed: {e}")
        raise e
    finally:
        # Clean up temporary PDF
        if os.path.exists(temp_pdf_path):
            try:
                os.remove(temp_pdf_path)
            except Exception as e:
                print(f"[excel_to_image] Warning: Failed to delete temporary PDF file: {e}")

if __name__ == "__main__":
    if len(sys.argv) < 3:
        print("Usage: python excel_to_image.py <excel_path> <output_png_path>")
        sys.exit(1)
    try:
        excel_p = sys.argv[1]
        png_p = sys.argv[2]
        convert_excel_to_image(excel_p, png_p)
    except Exception as e:
        print(f"Error: {e}")
        sys.exit(1)
