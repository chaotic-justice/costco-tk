import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from openpyxl import Workbook
import pandas as pd
import os
import sys
from pathlib import Path

from utils.tree import CostcoTree, pencil

import sys
import os
import platform

def check_dependencies():
    """Check and install required packages"""
    required_packages = ['pandas', 'numpy', 'tkinter']
    missing_packages = []

    for package in required_packages[:2]:  # Check pandas and numpy
        try:
            __import__(package)
        except ImportError:
            missing_packages.append(package)

    if missing_packages:
        print(f"Missing packages: {', '.join(missing_packages)}")
        response = input("Install automatically? (y/n): ")
        if response.lower() == 'y':
            import subprocess
            for package in missing_packages:
                print(f"Installing {package}...")
                subprocess.check_call([sys.executable, "-m", "pip", "install", package])

        # Re-check
        for package in missing_packages:
            try:
                __import__(package)
            except ImportError:
                print(f"Failed to install {package}. Please install manually.")
                print(f"Run: pip install {package}")
                return False

    return True

def create_launcher_scripts():
    """Create platform-specific launcher scripts"""
    current_platform = platform.system()

    if current_platform == "Windows":
        # Create Windows batch file
        batch_content = """@echo off
        echo Starting CSV Calculator...
        python "%~dp0pdf_counter.py"
        pause"""

        with open("run_csv_calculator.bat", "w") as f:
            f.write(batch_content)

        print("Created 'run_csv_calculator.bat' - Double-click this to run on Windows")

class PDFPageCounter:
    def __init__(self, root):
        self.root = root
        self.root.title("Costco PDFs Analyzer")
        self.root.geometry("700x550")

        # Variables
        date = pd.Timestamp.now()
        self.current_month_str = date.strftime("%B %Y")
        self.pdf_files = []
        self.output_filename = tk.StringVar(value=f"{self.current_month_str}_costco_output.xlsx")

        # Configure style
        self.setup_styles()

        # Create GUI
        self.create_widgets()

    def setup_styles(self):
        self.style = ttk.Style()
        self.style.theme_use('clam')

    def create_widgets(self):
        # Title
        emo = pencil()
        title_label = tk.Label(
            self.root,
            text=f"Costco PDFs Analyzer {emo}",
            font=("Arial", 16, "bold"),
            fg="#4dda52"
        )
        title_label.pack(pady=15)

        # Description
        desc_label = tk.Label(
            self.root,
            text="Upload up to 25 PDF files to generate costco report",
            font=("Arial", 13),
            fg="#7f8c8d"
        )
        desc_label.pack(pady=5)

        # File selection frame
        file_frame = tk.Frame(self.root)
        file_frame.pack(pady=10, padx=20, fill="both", expand=True)

        # File listbox with scrollbar
        listbox_frame = tk.Frame(file_frame)
        listbox_frame.pack(fill="both", expand=True)

        self.file_listbox = tk.Listbox(
            listbox_frame,
            height=8,
            selectmode=tk.EXTENDED,
            font=("Arial", 13),
            relief=tk.SOLID,
            borderwidth=1
        )
        self.file_listbox.pack(side=tk.LEFT, fill="both", expand=True, padx=(0, 5))

        scrollbar = tk.Scrollbar(listbox_frame, command=self.file_listbox.yview)
        scrollbar.pack(side=tk.RIGHT, fill="y")
        self.file_listbox.config(yscrollcommand=scrollbar.set)

        # Button frame for file operations
        file_btn_frame = tk.Frame(file_frame)
        file_btn_frame.pack(fill="x", pady=10)

        # File operation buttons
        ttk.Button(
            file_btn_frame,
            text="Add PDF Files",
            command=self.add_pdf_files,
            width=15
        ).pack(side=tk.LEFT, padx=(0, 10))

        ttk.Button(
            file_btn_frame,
            text="Clear All",
            command=self.clear_files,
            width=15
        ).pack(side=tk.LEFT, padx=(0, 10))

        ttk.Button(
            file_btn_frame,
            text="Remove Selected",
            command=self.remove_selected,
            width=15
        ).pack(side=tk.LEFT)

        # Output configuration frame
        output_frame = tk.Frame(self.root)
        output_frame.pack(pady=15, padx=20, fill="x")

        tk.Label(
            output_frame,
            text="Output Excel Configuration:",
            font=("Arial", 13, "bold")
        ).pack(anchor="w", pady=(0, 10))

        # Filename row
        filename_frame = tk.Frame(output_frame)
        filename_frame.pack(fill="x", pady=8)

        tk.Label(
            filename_frame,
            text="Filename:",
            font=("Arial", 13),
            width=10
        ).pack(side=tk.LEFT)

        self.output_entry = ttk.Entry(
            filename_frame,
            textvariable=self.output_filename,
            font=("Arial", 13),
            width=30
        )
        self.output_entry.pack(side=tk.LEFT, padx=(5, 10))

        # Save location row
        location_frame = tk.Frame(output_frame)
        location_frame.pack(fill="x", pady=5)

        tk.Label(
            location_frame,
            text="Save to:",
            font=("Arial", 13),
            width=10
        ).pack(side=tk.LEFT)

        self.save_location = tk.StringVar(value="Current Directory")
        self.location_label = tk.Label(
            location_frame,
            textvariable=self.save_location,
            font=("Arial", 13),
            relief=tk.SUNKEN,
            width=40,
            anchor="w",
            padx=5
        )
        self.location_label.pack(side=tk.LEFT, padx=(5, 10))

        ttk.Button(
            location_frame,
            text="Browse...",
            command=self.choose_save_location,
            width=10
        ).pack(side=tk.LEFT)

        # Default save location (Documents folder or home directory)
        self.default_save_dir = self.get_default_save_dir()

        # File count label
        self.file_count_label = tk.Label(
            self.root,
            text="Files: 0/25",
            font=("Arial", 9),
            fg="#7f8c8d"
        )
        self.file_count_label.pack()

        # Generate button
        generate_btn = tk.Button(
            self.root,
            text="Generate Excel Report",
            command=self.generate_report,
            fg="#1c8046",
            # fg="white",
            font=("Arial", 12, "bold"),
            padx=20,
            pady=10,
            relief=tk.RAISED,
            cursor="hand2"
        )
        generate_btn.pack(pady=20)

        # Status label
        self.status_label = tk.Label(
            self.root,
            text="Ready",
            font=("Arial", 13),
            fg="#7f8c8d"
        )
        self.status_label.pack()

    def get_default_save_dir(self):
        """Get a safe default directory for saving files"""
        try:
            # Try to get the Documents folder
            if sys.platform == "win32":
                import ctypes.wintypes
                CSIDL_PERSONAL = 5  # My Documents
                SHGFP_TYPE_CURRENT = 0  # Get current, not default value

                buf = ctypes.create_unicode_buffer(ctypes.wintypes.MAX_PATH)
                ctypes.windll.shell32.SHGetFolderPathW(None, CSIDL_PERSONAL, None, SHGFP_TYPE_CURRENT, buf)
                documents_path = buf.value
                if os.path.exists(documents_path):
                    return documents_path
            elif sys.platform == "darwin":  # macOS
                home = os.path.expanduser("~")
                documents = os.path.join(home, "Documents")
                if os.path.exists(documents):
                    return documents

            # Fallback to home directory
            home = os.path.expanduser("~")
            if os.path.exists(home):
                return home

            # Fallback to current directory
            return os.getcwd()

        except:
            # If all else fails, use current directory
            return os.getcwd()

    def choose_save_location(self):
        """Let user choose where to save the Excel file"""
        directory = filedialog.askdirectory(
            title="Select folder to save Excel file",
            initialdir=self.default_save_dir
        )
        if directory:
            self.save_location.set(directory)

    def add_pdf_files(self):
        # Open file dialog for PDF files
        files = filedialog.askopenfilenames(
            title="Select PDF Files",
            filetypes=[("PDF files", "*.pdf"), ("All files", "*.*")]
        )

        # Check if adding these files would exceed limit
        if len(self.pdf_files) + len(files) > 25:
            messagebox.showwarning(
                "Limit Exceeded",
                f"You can only select up to 25 files. You already have {len(self.pdf_files)} files selected."
            )
            return

        # Add new files
        new_files_added = 0
        for file in files:
            self.pdf_files.append(file)
            filename = os.path.basename(file)
            self.file_listbox.insert(tk.END, filename)
            new_files_added += 1

        # Update file count
        self.update_file_count()

        if new_files_added > 0:
            self.status_label.config(text=f"Added {new_files_added} PDF file(s)")

    def clear_files(self):
        self.pdf_files.clear()
        self.file_listbox.delete(0, tk.END)
        self.update_file_count()
        self.status_label.config(text="All files cleared")

    def remove_selected(self):
        # Get selected items in reverse order to maintain correct indices
        selected_indices = self.file_listbox.curselection()
        removed_count = len(selected_indices)

        for index in reversed(selected_indices):
            self.file_listbox.delete(index)
            del self.pdf_files[index]

        self.update_file_count()
        if removed_count > 0:
            self.status_label.config(text=f"Removed {removed_count} file(s)")

    def update_file_count(self):
        count = len(self.pdf_files)
        self.file_count_label.config(text=f"Files: {count}/25")

    def get_safe_save_path(self):
        """Get a safe path where we can save the file"""
        # Get the save directory from user selection or use default
        save_dir = self.save_location.get()
        if save_dir == "Current Directory":
            # Try multiple locations
            locations_to_try = [
                self.default_save_dir,
                os.path.expanduser("~"),
                os.getcwd()
            ]

            for location in locations_to_try:
                try:
                    # Test if we can write to this location
                    test_file = os.path.join(location, "test_write.tmp")
                    with open(test_file, 'w') as f:
                        f.write("test")
                    os.remove(test_file)
                    return location
                except:
                    continue

            # If all fail, ask user
            directory = filedialog.askdirectory(
                title="Select a writable folder for saving Excel file"
            )
            if directory:
                return directory
            else:
                return None
        else:
            return save_dir

    def generate_report(self):
        if not self.pdf_files:
            messagebox.showinfo("No Files", "Please select at least one PDF file.")
            return

        # Get output filename
        output_filename = self.output_filename.get().strip()
        if not output_filename:
            messagebox.showwarning("Invalid Filename", "Please enter a valid output filename.")
            return

        # Ensure .xlsx extension
        if not output_filename.lower().endswith('.xlsx'):
            output_filename += '.xlsx'

        # Get safe save location
        save_dir = self.get_safe_save_path()
        if not save_dir:
            self.status_label.config(text="No valid save location selected")
            return

        # Create full path
        output_path = os.path.join(save_dir, output_filename)

        # Check if file exists and ask for confirmation
        if os.path.exists(output_path):
            response = messagebox.askyesno(
                "File Exists",
                f"'{output_filename}' already exists in\n{save_dir}\n\nOverwrite?"
            )
            if not response:
                self.status_label.config(text="Operation cancelled")
                return

        # Update status
        self.status_label.config(text="Processing PDF files...")
        self.root.update()  # Update GUI to show status change

        try:
            # Try to import required libraries
            try:
                import pandas as pd
            except ImportError:
                messagebox.showerror(
                    "Missing Dependency",
                    "Pandas is required. Please install it with:\n\n"
                    "pip install pandas"
                )
                self.status_label.config(text="Pandas not installed")
                return

            cct = CostcoTree(
                dir_path="costco",
                pdf_files=self.pdf_files,
                output_path=output_path
            )
            wb = Workbook()
            wb.remove(wb.active)
            for idx, pdf_path in enumerate(self.pdf_files):
                df1, df2, tab_name = cct.get_table_from_pdf(pdf_path=pdf_path)
                cct.draw(df1, df2, tab_name=tab_name, wb=wb)

            # Try to save the file
            try:
                # df.to_excel(output_path, index=False)

                # Update status
                self.status_label.config(text=f"Report saved to: {output_path}")

                # Show success message with option to open the file
                response = messagebox.askyesno(
                    "Success",
                    f"Excel report generated successfully!\n\n"
                    f"Saved to: {output_path}\n"
                    f"Total PDFs: {len(self.pdf_files)}\n"
                    f"Would you like to open the file?"
                )

                if response:
                    # Open the file with default application
                    try:
                        if sys.platform == "win32":
                            os.startfile(output_path)
                        elif sys.platform == "darwin":  # macOS
                            subprocess.run(["open", output_path])
                        else:  # Linux
                            subprocess.run(["xdg-open", output_path])
                    except:
                        pass

            except PermissionError:
                self.status_label.config(text="Permission denied - try different location")
                messagebox.showerror(
                    "Permission Error",
                    f"Cannot save to:\n{output_path}\n\n"
                    f"The location may be read-only or you don't have permission.\n"
                    f"Try saving to a different folder (use Browse button)."
                )
                # Try to save to a temp location as fallback
                # self.save_to_temp_fallback(df, output_filename)

            except Exception as e:
                self.status_label.config(text=f"Error saving file: {str(e)[:50]}")
                messagebox.showerror("Save Error", f"Failed to save file:\n{str(e)}")

        except Exception as e:
            self.status_label.config(text="Error generating report")
            messagebox.showerror("Error", f"Failed to generate report:\n{str(e)}")

    def save_to_temp_fallback(self, df, original_filename):
        """Save to temp directory as fallback"""
        try:
            import tempfile

            # Create a temp directory
            temp_dir = tempfile.gettempdir()
            temp_path = os.path.join(temp_dir, original_filename)

            # Save to temp location
            df.to_excel(temp_path, index=False)

            # Update status
            self.status_label.config(text=f"Saved to temp: {temp_path}")

            messagebox.showinfo(
                "Saved to Temporary Location",
                f"File saved to temporary location:\n\n"
                f"{temp_path}\n\n"
                f"Please move it to your desired location."
            )

        except Exception as e:
            self.status_label.config(text="Failed to save anywhere")
            messagebox.showerror(
                "Critical Error",
                f"Could not save file anywhere:\n{str(e)}"
            )



def main():
    root = tk.Tk()
    app = PDFPageCounter(root)

    # Center the window on screen
    root.update_idletasks()
    width = root.winfo_width()
    height = root.winfo_height()
    x = (root.winfo_screenwidth() // 2) - (width // 2)
    y = (root.winfo_screenheight() // 2) - (height // 2)
    root.geometry(f'{width}x{height}+{x}+{y}')

    root.mainloop()


if __name__ == "__main__":
    # Check for required imports
    print("="*60)
    print("PDF Counter")
    print("="*60)
    print(f"Platform: {platform.system()} {platform.release()}")
    print(f"Python: {sys.version}")
    print()

    # Check dependencies
    if not check_dependencies():
        print("\nDependencies missing. Please install required packages.")
        print("Run: pip install pandas numpy")
        input("Press Enter to exit...")
        sys.exit(1)

    # Create launcher scripts
    create_launcher_scripts()
    try:
        import pandas as pd
    except ImportError:
        print("Pandas is required. Install with: pip install pandas")
        print("Trying to install automatically...")
        try:
            import subprocess
            import sys
            subprocess.check_call([sys.executable, "-m", "pip", "install", "pandas", "openpyxl"])
            print("Installation successful. Please restart the application.")
        except:
            print("Failed to install automatically. Please run:")
            print("pip install pandas openpyxl")
        input("Press Enter to exit...")
        sys.exit(1)

    main()