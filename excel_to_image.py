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
        
        # Set minimal margins (0.1 inches) to reduce initial blank margins
        ws.PageSetup.LeftMargin = excel.InchesToPoints(0.1)
        ws.PageSetup.RightMargin = excel.InchesToPoints(0.1)
        ws.PageSetup.TopMargin = excel.InchesToPoints(0.1)
        ws.PageSetup.BottomMargin = excel.InchesToPoints(0.1)
        
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
                
    # Convert PDF to PNG via PyMuPDF (fitz) at 300 DPI, cropping empty margin space
    try:
        if not os.path.exists(temp_pdf_path):
            raise FileNotFoundError(f"Temporary PDF file was not created: {temp_pdf_path}")
            
        print(f"[excel_to_image] Reading PDF and calculating content bounding box for tight cropping...")
        doc = fitz.open(temp_pdf_path)
        page = doc[0]  # First page
        
        # Compute bounding box of content (text and drawing paths like table borders)
        x0, y0, x1, y1 = float('inf'), float('inf'), float('-inf'), float('-inf')
        
        # 1. Bounding box from text blocks
        blocks = page.get_text("blocks")
        for b in blocks:
            x0 = min(x0, b[0])
            y0 = min(y0, b[1])
            x1 = max(x1, b[2])
            y1 = max(y1, b[3])
            
        # 2. Bounding box from drawings (lines, cells, borders)
        drawings = page.get_drawings()
        for d in drawings:
            rect = d.get("rect")
            if rect:
                # Exclude shapes that cover the entire page (like background color fills)
                if rect.width < page.rect.width * 0.99 or rect.height < page.rect.height * 0.99:
                    x0 = min(x0, rect[0])
                    y0 = min(y0, rect[1])
                    x1 = max(x1, rect[2])
                    y1 = max(y1, rect[3])
                    
        # Apply crop box if content is found, else render whole page
        if x0 != float('inf'):
            # Add a small padding (5 points) to avoid clipping text borders
            padding = 5
            cropped_rect = fitz.Rect(
                max(0, x0 - padding),
                max(0, y0 - padding),
                min(page.rect.width, x1 + padding),
                min(page.rect.height, y1 + padding)
            )
            print(f"[excel_to_image] Cropping page to content bounding box: {cropped_rect}")
            page.set_cropbox(cropped_rect)
        else:
            print("[excel_to_image] Warning: No content found, rendering full page.")
            
        # Render page at 300 DPI (zoom factor is 300 / 72)
        zoom = 300 / 72
        mat = fitz.Matrix(zoom, zoom)
        pix = page.get_pixmap(matrix=mat)
        
        # Save image
        pix.save(output_png_path)
        doc.close()
        print(f"[excel_to_image] Cropped image successfully generated: {output_png_path}")
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
