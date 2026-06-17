import os
import re
import sys
import subprocess
from report_generator import generate_report
from excel_to_image import convert_excel_to_image

def find_latest_day():
    base_dir = os.path.dirname(os.path.abspath(__file__))
    input_dir = os.path.join(base_dir, "input-this_years_support_status")
    
    if not os.path.exists(input_dir):
        print(f"[run_pipeline] Error: Input directory '{input_dir}' does not exist.")
        return None
        
    pattern = re.compile(r"2027_apply3-(\d+)day-input\.xlsx")
    days = []
    
    for filename in os.listdir(input_dir):
        match = pattern.match(filename)
        if match:
            days.append(int(match.group(1)))
            
    if not days:
        print("[run_pipeline] Warning: No input files matching pattern '2027_apply3-Nday-input.xlsx' found.")
        return None
        
    latest_day = max(days)
    print(f"[run_pipeline] Detected latest day input file: Day {latest_day}")
    return latest_day

def run_git_command(args, base_dir):
    try:
        print(f"[run_pipeline] Executing: {' '.join(args)} in {base_dir}")
        result = subprocess.run(args, capture_output=True, text=True, errors="replace", check=True, cwd=base_dir)
        print(result.stdout)
        if result.stderr:
            print(result.stderr)
        return True
    except subprocess.CalledProcessError as e:
        print(f"[run_pipeline] Git command failed: {' '.join(args)}")
        print(f"Stdout: {e.stdout}")
        print(f"Stderr: {e.stderr}")
        return False

def main():
    base_dir = os.path.dirname(os.path.abspath(__file__))
    
    # 1. Detect latest day
    day = find_latest_day()
    if day is None:
        print("[run_pipeline] Error: Could not detect current day number. Aborting.")
        sys.exit(1)
        
    xlsx_filename = f"2027_report_{day}day.xlsx"
    png_filename = f"2027_report_{day}day.png"
    
    excel_path = os.path.join(base_dir, "output-report", xlsx_filename)
    png_path = os.path.join(base_dir, "output-report", png_filename)
    
    # 2. Run report generator
    try:
        generate_report(day)
    except Exception as e:
        print(f"[run_pipeline] Error during report generation: {e}")
        sys.exit(1)
        
    # 3. Run excel to image converter (also saves the evaluated formulas cache)
    try:
        convert_excel_to_image(excel_path, png_path)
    except Exception as e:
        print(f"[run_pipeline] Error during excel to image conversion: {e}")
        sys.exit(1)
        
    print(f"[run_pipeline] Day {day} report and image generated successfully.")
    
    # 4. Git Stage, Commit & Push
    # We use relative paths from the workspace root for Git operations
    rel_xlsx_path = os.path.relpath(excel_path, base_dir)
    rel_png_path = os.path.relpath(png_path, base_dir)
    
    print("[run_pipeline] Starting GitHub synchronization...")
    
    # Stage files
    if run_git_command(["git", "add", rel_xlsx_path, rel_png_path], base_dir):
        # Commit files
        commit_message = f"Auto-update: Day {day} Support Status Report"
        if run_git_command(["git", "commit", "-m", commit_message], base_dir):
            # Push to main branch
            run_git_command(["git", "push", "origin", "main"], base_dir)
        else:
            print("[run_pipeline] Git commit failed (might be no changes to commit).")
    else:
        print("[run_pipeline] Git add failed.")

if __name__ == "__main__":
    main()
