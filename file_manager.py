import tkinter as tk
from tkinter import filedialog, Label, messagebox, Frame, Button, PhotoImage
import time
import os
from tkinter import ttk

class FileManager:
    """
    Handles file operations for the presentation application.
    """

    def __init__(self):
        """Initialize file manager with no selected file."""
        self.selected_file = None
        self.file_name = None

    def show_file_dialog(self):
        """Shows a custom-styled file dialog to select a PDF or PPT file."""
        try:
            # Create custom dialog window
            dialog = tk.Tk()
            dialog.title("Select Presentation File")
            dialog.geometry("800x500")
            dialog.configure(bg="#d4eac8")  # Match splash screen color

            # Center the window
            screen_width = dialog.winfo_screenwidth()
            screen_height = dialog.winfo_screenheight()
            x = (screen_width - 800) // 2
            y = (screen_height - 500) // 2
            dialog.geometry(f"800x500+{x}+{y}")

            # Configure styles
            style = ttk.Style()
            style.configure('Custom.TFrame', background='#d4eac8')
            style.configure('Custom.TButton', 
                           background='#4a934a',
                           foreground='white',
                           padding=10,
                           font=('Arial', 10))
            style.configure('Custom.TLabel', 
                           background='#d4eac8',
                           font=('Arial', 12))

            # Create main frame
            main_frame = ttk.Frame(dialog, style='Custom.TFrame')
            main_frame.pack(fill='both', expand=True, padx=20, pady=20)

            # Title label
            title_label = ttk.Label(main_frame, 
                                   text="Select Your Presentation File",
                                   font=('Arial', 16, 'bold'),
                                   style='Custom.TLabel')
            title_label.pack(pady=(0, 20))

            # File type buttons frame
            button_frame = ttk.Frame(main_frame, style='Custom.TFrame')
            button_frame.pack(fill='x', pady=20)

            def select_file(file_type):
                filetypes = []
                if file_type == "pdf":
                    filetypes = [("PDF Files", "*.pdf")]
                elif file_type == "ppt":
                    filetypes = [("PowerPoint Files", "*.pptx")]
                else:
                    filetypes = [
                        ("All Presentations", "*.pdf;*.pptx"),
                        ("PDF Files", "*.pdf"),
                        ("PowerPoint Files", "*.pptx")
                    ]

                file_path = filedialog.askopenfilename(
                    parent=dialog,
                    title="Select File",
                    filetypes=filetypes,
                    initialdir=os.path.expanduser("~\\Documents")
                )
                
                if file_path:
                    self.selected_file = file_path
                    dialog.quit()
                    dialog.destroy()
                    return self.selected_file

            # Styled buttons
            pdf_btn = ttk.Button(button_frame, 
                                text="Select PDF",
                                style='Custom.TButton',
                                command=lambda: select_file("pdf"))
            pdf_btn.pack(side='left', padx=10, expand=True)

            ppt_btn = ttk.Button(button_frame, 
                                text="Select PowerPoint",
                                style='Custom.TButton',
                                command=lambda: select_file("ppt"))
            ppt_btn.pack(side='right', padx=10, expand=True)

            # Instructions
            instructions = ttk.Label(main_frame,
                                   text="Choose your presentation file type\n"
                                        "Supported formats: PDF (.pdf) and PowerPoint (.pptx)",
                                   style='Custom.TLabel')
            instructions.pack(pady=20)

            # Make dialog modal
            dialog.lift()
            dialog.focus_force()
            dialog.grab_set()

            # Start the dialog
            dialog.mainloop()

            return self.selected_file

        except Exception as e:
            print(f"Error in file dialog: {e}")
            return None

    def show_help_dialog(self):
        """
        Display help dialog with gesture instructions.
        """
        root = tk.Tk()
        root.withdraw()
        
        help_window = tk.Toplevel(root)
        help_window.title("Gesture Controls Help")
        help_window.geometry("500x400")
        help_window.configure(bg="#f0f0f0")
        
        # Center the window
        window_width = 500
        window_height = 400
        screen_width = help_window.winfo_screenwidth()
        screen_height = help_window.winfo_screenheight()
        x = (screen_width / 2) - (window_width / 2)
        y = (screen_height / 2) - (window_height / 2)
        help_window.geometry(f"{window_width}x{window_height}+{int(x)}+{int(y)}")
        
        # Title
        title_label = Label(
            help_window, 
            text="Hand Gesture Controls", 
            font=("Arial", 16, "bold"),
            bg="#f0f0f0",
            fg="#333333"
        )
        title_label.pack(pady=(20, 15))
        
        # Instructions frame
        instructions_frame = Frame(help_window, bg="#f0f0f0")
        instructions_frame.pack(fill="both", expand=True, padx=25, pady=10)
        
        # Gesture instructions
        gestures = [
            ("Next Slide", "Show all fingers except thumb [0,1,1,1,1]"),
            ("Previous Slide", "Show only thumb [1,0,0,0,0]"),
            ("Pointer", "Show index and middle fingers [0,1,1,0,0]"),
            ("Draw", "Show only index finger [0,1,0,0,0]"),
            ("Erase", "Show index, middle, and ring fingers [0,1,1,1,0]"),
            ("Clear Screen", "Show all five fingers [1,1,1,1,1]"),
        ]
        
        for i, (action, gesture) in enumerate(gestures):
            row_frame = Frame(instructions_frame, bg="#f0f0f0")
            row_frame.pack(fill="x", pady=5, anchor="w")
            
            action_label = Label(
                row_frame, 
                text=action, 
                font=("Arial", 11, "bold"),
                width=15,
                anchor="w",
                bg="#f0f0f0",
                fg="#333333"
            )
            action_label.pack(side=tk.LEFT)
            
            gesture_label = Label(
                row_frame, 
                text=gesture, 
                font=("Arial", 10),
                bg="#f0f0f0",
                fg="#555555"
            )
            gesture_label.pack(side=tk.LEFT, padx=(10, 0))
        
        # Note about gesture area
        note_frame = Frame(help_window, bg="#f0f0f0")
        note_frame.pack(fill="x", padx=25, pady=(20, 5))
        
        note_label = Label(
            note_frame, 
            text="Note: Gestures are only detected in the upper right corner of the screen (above the green dotted line).",
            font=("Arial", 9),
            bg="#f0f0f0",
            fg="#666666",
            wraplength=450,
            justify=tk.LEFT
        )
        note_label.pack(anchor="w")
        
        # Close button
        button_frame = Frame(help_window, bg="#f0f0f0")
        button_frame.pack(pady=20)
        
        close_btn = Button(
            button_frame,
            text="Close",
            command=lambda: [help_window.destroy(), root.destroy()],
            bg="#4285f4",
            fg="white",
            font=("Arial", 10),
            relief=tk.FLAT,
            padx=20,
            pady=8,
            cursor="hand2"
        )
        close_btn.pack()
        
        help_window.protocol("WM_DELETE_WINDOW", lambda: [help_window.destroy(), root.destroy()])
        help_window.lift()
        help_window.focus_force()
        
        root.mainloop()

    def save_file_dialog(self, original_extension):
        """
        Show a dialog to save the file with annotations.
        
        Args:
            original_extension: The extension of the original file (.pdf or .pptx)
            
        Returns:
            The file path to save to, or None if canceled
        """
        root = tk.Tk()
        root.withdraw()
        
        filetypes = []
        if original_extension.lower() == ".pdf":
            filetypes = [("PDF Files", "*.pdf")]
            default_ext = ".pdf"
        elif original_extension.lower() == ".pptx":
            filetypes = [("PowerPoint Files", "*.pptx")]
            default_ext = ".pptx"
        
        filepath = filedialog.asksaveasfilename(
            defaultextension=default_ext,
            filetypes=filetypes,
            title="Save Presentation"
        )
        
        root.destroy()
        return filepath

if __name__ == "__main__":
    # Test the file manager
    fm = FileManager()
    selected_file = fm.show_file_dialog()
    print(f"Selected file: {selected_file}")