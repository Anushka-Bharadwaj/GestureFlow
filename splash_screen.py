
import tkinter as tk
from PIL import Image, ImageTk
from file_manager import FileManager

def show_splash_screen(start_callback):
    splash = tk.Tk()
    splash.title("Splash Screen")
    splash.geometry("600x400")
    splash.configure(bg="#d4eac8")
    splash.overrideredirect(True)

    # Center the splash screen
    screen_width = splash.winfo_screenwidth()
    screen_height = splash.winfo_screenheight()
    x = (screen_width - 600) // 2
    y = (screen_height - 400) // 2
    splash.geometry(f"600x400+{x}+{y}")

    # Load and display GIF
    gif_path = r"D:\INHOUSE PROJECT-- GestureFlow\GestureFlow\Thank you!.gif"
    gif = Image.open(gif_path)

    gif_label = tk.Label(splash, bg="#d4eac8")
    gif_label.pack(expand=True)

    def animate_gif(frame_index):
        try:
            gif.seek(frame_index)
            frame = ImageTk.PhotoImage(gif)
            gif_label.config(image=frame)
            gif_label.image = frame
            next_frame = frame_index + 1 if frame_index + 1 < gif.n_frames else 0
            splash.after(100, animate_gif, next_frame)
        except Exception:
            pass

    def transition_to_main():
        try:
            file_manager = FileManager()
            file_path = file_manager.show_file_dialog()
            if file_path:
                splash.destroy()
                start_callback(file_path)
        except Exception as e:
            print(f"Error during transition: {e}")
            splash.destroy()

    # Start GIF animation
    animate_gif(0)
    
    # Show splash for 10 seconds then transition
    splash.after(10000, transition_to_main)
    
    splash.mainloop()
