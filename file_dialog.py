import tkinter as tk
from tkinter import filedialog
import os

def show_enhanced_file_dialog():
    """Shows an enhanced unified file dialog for selecting PDF and PPT files."""
    try:
        # Create custom dialog window
        dialog = tk.Tk()
        dialog.title("Select Your Presentation")
        dialog.geometry("800x600")
        dialog.configure(bg="#d4eac8")

        # Center the window
        screen_width = dialog.winfo_screenwidth()
        screen_height = dialog.winfo_screenheight()
        x = (screen_width - 800) // 2
        y = (screen_height - 600) // 2
        dialog.geometry(f"800x600+{x}+{y}")

        # Variable to store selected file
        selected_file = None

        # Main container
        main_frame = tk.Frame(dialog, bg="#d4eac8")
        main_frame.pack(fill="both", expand=True, padx=40, pady=30)

        # Title with enhanced styling
        title_label = tk.Label(
            main_frame,
            text="Choose Your Presentation File",
            font=("Arial", 24, "bold"),
            bg="#d4eac8",
            fg="#2e5b1e"
        )
        title_label.pack(pady=(0, 30))

        # File type frame with modern design
        file_frame = tk.Frame(main_frame, bg="#d4eac8")
        file_frame.pack(fill="x", pady=20)

        def select_file():
            nonlocal selected_file
            filetypes = [
                ("Presentation Files", "*.pdf;*.pptx"),
                ("PDF Files", "*.pdf"),
                ("PowerPoint Files", "*.pptx")
            ]
            
            file_path = filedialog.askopenfilename(
                parent=dialog,
                title="Select Presentation File",
                filetypes=filetypes,
                initialdir=os.path.expanduser("~\\Documents")
            )
            
            if file_path:
                selected_file = file_path
                dialog.quit()
                dialog.destroy()

        # Modern styled button
        select_button = tk.Button(
            file_frame,
            text="Browse Files",
            command=select_file,
            font=("Arial", 12, "bold"),
            bg="#4a934a",
            fg="white",
            relief="flat",
            padx=40,
            pady=15,
            cursor="hand2"
        )
        select_button.pack(pady=20)

        # Add hover effect
        def on_enter(e):
            select_button['bg'] = '#3a733a'
        def on_leave(e):
            select_button['bg'] = '#4a934a'

        select_button.bind("<Enter>", on_enter)
        select_button.bind("<Leave>", on_leave)

        # Supported formats section
        formats_frame = tk.Frame(main_frame, bg="#d4eac8")
        formats_frame.pack(fill="x", pady=20)

        supported_label = tk.Label(
            formats_frame,
            text="Supported Formats:",
            font=("Arial", 12, "bold"),
            bg="#d4eac8",
            fg="#2e5b1e"
        )
        supported_label.pack(pady=(0, 10))

        formats = [
            "PDF Documents (.pdf)",
            "PowerPoint Presentations (.pptx)"
        ]

        for fmt in formats:
            fmt_label = tk.Label(
                formats_frame,
                text=f"• {fmt}",
                font=("Arial", 11),
                bg="#d4eac8",
                fg="#333333"
            )
            fmt_label.pack(pady=2)

        # Instructions
        instruction_text = """
        1. Click 'Browse Files' to select your presentation
        2. Choose either a PDF or PowerPoint file
        3. Your selected file will open in presentation mode
        """
        
        instructions = tk.Label(
            main_frame,
            text=instruction_text,
            font=("Arial", 10),
            bg="#d4eac8",
            fg="#555555",
            justify="left"
        )
        instructions.pack(pady=30)

        # Make dialog modal
        dialog.transient()
        dialog.grab_set()
        dialog.focus_force()

        dialog.mainloop()
        return selected_file

    except Exception as e:
        print(f"Error in file dialog: {e}")
        return None

if __name__ == "__main__":
    # Test the dialog
    file_path = show_enhanced_file_dialog()
    print(f"Selected file: {file_path}")