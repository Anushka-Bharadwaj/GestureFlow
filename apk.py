#!/usr/bin/env python3
import cv2
import os
import numpy as np
import time
import fitz
import tkinter as tk
from PIL import Image, ImageTk, ImageSequence
from HandTracker import HandDetector
from file_manager import FileManager
from splash_screen import show_splash_screen
from ppt_handler import load_powerpoint
import comtypes.client

def run_main_application(initial_file_path=None):
    """Main function to run the hand gesture presentation control application."""
    print("Starting Hand Gesture Presentation Controller...")
    
    # Initialize components and variables
    width, height = 1280, 720
    detector = HandDetector(detectionCon=0.8, maxHands=1)
    file_manager = FileManager()
    slide_images = []
    
    # Initialize camera first and ensure it's working
    cap = None
    for i in range(5):  # Try multiple times to initialize camera
        try:
            cap = cv2.VideoCapture(i, cv2.CAP_DSHOW)
            if cap.isOpened():
                # Set camera properties
                cap.set(cv2.CAP_PROP_FRAME_WIDTH, width)
                cap.set(cv2.CAP_PROP_FRAME_HEIGHT, height)
                cap.set(cv2.CAP_PROP_AUTOFOCUS, 1)
                # Test camera read
                ret, frame = cap.read()
                if ret:
                    break
            cap.release()
        except Exception as e:
            print(f"Camera init attempt {i} failed: {e}")
    
    if cap is None or not cap.isOpened():
        print("ERROR: Could not initialize camera")
        return
    
    # Create window before loading file
    cv2.namedWindow("Slides", cv2.WINDOW_NORMAL)
    cv2.setWindowProperty("Slides", cv2.WND_PROP_FULLSCREEN, cv2.WINDOW_FULLSCREEN)
    
    # Load initial file if provided
    if initial_file_path:
        file_path = initial_file_path
        if file_path.lower().endswith(".pdf"):
            try:
                doc = fitz.open(file_path)
                print(f"PDF loaded with {len(doc)} pages")
                
                for page in doc:
                    zoom = 2
                    mat = fitz.Matrix(zoom, zoom)
                    pix = page.get_pixmap(matrix=mat, alpha=False)
                    img = np.frombuffer(pix.samples, dtype=np.uint8).reshape(pix.h, pix.w, 3)
                    img = cv2.cvtColor(img, cv2.COLOR_RGB2BGR)
                    img = cv2.resize(img, (width, height))
                    slide_images.append(img)
                
                doc.close()
            except Exception as e:
                print(f"Error loading PDF: {e}")
                blank = np.ones((height, width, 3), dtype=np.uint8) * 255
                cv2.putText(blank, "Error loading PDF", (width//2 - 150, height//2),
                           cv2.FONT_HERSHEY_SIMPLEX, 1, (0, 0, 255), 2)
                slide_images.append(blank)
        elif file_path.lower().endswith(".pptx"):
            slide_images = load_powerpoint(file_path, width, height)
    
    # Initialize other variables
    gest_done = False
    gest_counter = 0
    ge_thresh_y = 450  # Increased threshold for easier gesture detection
    ge_thresh_x = 650  # Reduced threshold to give more space for gestures
    delay = 15
    annotations = [[]]
    annot_num = 0
    annot_start = False
    slide_num = 0
    hs, ws = int(120 * 1.2), int(213 * 1.2)
    zoom_level = 1.0
    zoom_change = 0.1
    min_zoom = 0.5
    max_zoom = 2.0
    last_gesture_time = time.time()
    gesture_cooldown = 0.3
    is_fullscreen = True
    
    if initial_file_path:
        file_path = initial_file_path
    else:
        file_path = file_manager.show_file_dialog()
        
    if file_path:
        if file_path.lower().endswith(".pdf"):
            try:
                doc = fitz.open(file_path)
                print(f"PDF loaded with {len(doc)} pages")
                
                for page in doc:
                    # Increase the resolution of rendered pages
                    zoom = 2  # Increase zoom factor for better quality
                    mat = fitz.Matrix(zoom, zoom)
                    pix = page.get_pixmap(matrix=mat, alpha=False)
                    
                    # Convert to numpy array
                    img = np.frombuffer(pix.samples, dtype=np.uint8).reshape(pix.h, pix.w, 3)
                    
                    # Convert from RGB to BGR (OpenCV format)
                    img = cv2.cvtColor(img, cv2.COLOR_RGB2BGR)
                    
                    # Resize to match window dimensions
                    img = cv2.resize(img, (width, height))
                    
                    slide_images.append(img)
                
                print(f"Successfully loaded {len(slide_images)} slides from PDF")
                doc.close()  # Close the PDF file
            except Exception as e:
                print(f"Error loading PDF file: {e}")
                # Create a blank slide with error message
                blank = np.ones((height, width, 3), dtype=np.uint8) * 255
                cv2.putText(blank, "Error loading PDF", (width//2 - 150, height//2), 
                            cv2.FONT_HERSHEY_SIMPLEX, 1, (0, 0, 255), 2)
                slide_images.append(blank)
        
        elif file_path.lower().endswith(".pptx"):
            slide_images = load_powerpoint(file_path, width, height)
    else:
        # Create a blank image that says "No presentation loaded"
        blank = np.ones((height, width, 3), dtype=np.uint8) * 255
        
        # Add text to the blank image
        text = "No presentation loaded"
        font = cv2.FONT_HERSHEY_SIMPLEX
        textsize = cv2.getTextSize(text, font, 1, 2)[0]
        
        # Center the text
        textX = (blank.shape[1] - textsize[0]) // 2
        textY = (blank.shape[0] + textsize[1]) // 2
        
        cv2.putText(blank, text, (textX, textY), font, 1, (0, 0, 0), 2)
        
        # Add instruction
        instruction = "Press 'O' to open a file or 'Q' to quit"
        instructionsize = cv2.getTextSize(instruction, font, 0.7, 1)[0]
        instructionX = (blank.shape[1] - instructionsize[0]) // 2
        
        cv2.putText(blank, instruction, (instructionX, textY + 50), font, 0.7, (100, 100, 100), 1)
        
        slide_images.append(blank)
    
    print("Application started. Press 'H' for help, 'O' to open a file, 'Q' to quit.")
    
    while True:
        success, frame = cap.read()
        if not success:
            print("Failed to capture frame from webcam")
            break
            
        frame = cv2.flip(frame, 1)
        
        if slide_images:
            # Apply zoom level to the slide
            if zoom_level != 1.0:
                # Get dimensions for zoomed image
                h, w = slide_images[slide_num].shape[:2]
                new_h, new_w = int(h * zoom_level), int(w * zoom_level)
                
                # Resize with zoom factor
                zoomed_img = cv2.resize(slide_images[slide_num], (new_w, new_h))
                
                # Create a blank slide of the original size
                slide_current = np.ones((height, width, 3), dtype=np.uint8) * 255
                
                # Calculate centering offsets
                offset_x = max(0, (width - new_w) // 2)
                offset_y = max(0, (height - new_h) // 2)
                
                # If zoomed image is larger than the window, take the center portion
                if new_w > width:
                    start_x = (new_w - width) // 2
                    end_x = start_x + width
                    offset_x = 0
                else:
                    start_x = 0
                    end_x = new_w
                
                if new_h > height:
                    start_y = (new_h - height) // 2
                    end_y = start_y + height
                    offset_y = 0
                else:
                    start_y = 0
                    end_y = new_h
                
                # Place the zoomed portion into the slide
                try:
                    slide_current[offset_y:offset_y+(end_y-start_y), offset_x:offset_x+(end_x-start_x)] = zoomed_img[start_y:end_y, start_x:end_x]
                except ValueError:
                    # If any issues with dimensions, fallback to standard resize
                    slide_current = cv2.resize(slide_images[slide_num], (width, height))
            else:
                # Normal resize if zoom_level is 1.0
                slide_current = cv2.resize(slide_images[slide_num], (width, height))
        else:
            slide_current = np.ones((height, width, 3), dtype=np.uint8) * 255
            cv2.putText(slide_current, "No slides available", (width//2 - 150, height//2), 
                       cv2.FONT_HERSHEY_SIMPLEX, 1, (0, 0, 255), 2)

        # Find hand and landmarks
        hands, frame = detector.findHands(frame)
        
        # Improved gesture detection
        if hands and not gest_done:
            hand = hands[0]
            cx, cy = hand["center"]  
            lm_list = hand["lmList"]
            fingers = detector.fingersUp(hand)
            
            # More precise coordinate mapping for drawing
            x_val = int(np.interp(lm_list[8][0], [0, frame.shape[1]], [0, width]))
            y_val = int(np.interp(lm_list[8][1], [0, frame.shape[0]], [0, height]))
            index_fing = (x_val, y_val)
                
            current_time = time.time()
            time_since_last = current_time - last_gesture_time
            can_perform_gesture = time_since_last > gesture_cooldown

            if cy < ge_thresh_y and cx > ge_thresh_x:
                annot_start = False

                # Previous Slide - Thumb only
                if fingers == [1, 0, 0, 0, 0]:
                    if slide_num > 0:
                        slide_num -= 1
                        annotations = [[]]
                        annot_num = 0
                        gest_done = True
                        last_gesture_time = current_time

                # Next Slide
                if fingers == [0, 1, 1, 1, 1] and can_perform_gesture:
                    annot_start = False
                    if slide_num < len(slide_images) - 1:
                        gest_done = True
                        slide_num += 1
                        annotations = [[]]
                        annot_num = 0
                        last_gesture_time = current_time

                # Clear Annotations
                if fingers == [1, 1, 1, 1, 1] and can_perform_gesture:
                    if annotations:
                        annot_start = False
                        annotations.clear()
                        annot_num = 0
                        gest_done = True
                        annotations = [[]]
                        last_gesture_time = current_time

                # Zoom In - Pinch gesture (thumb + index)
                if fingers == [1, 1, 0, 0, 0] and can_perform_gesture:
                    # Check distance between thumb and index finger
                    if lm_list[4] and lm_list[8]:  # Thumb tip and index tip
                        # Calculate distance
                        thumb_tip = lm_list[4]
                        index_tip = lm_list[8]
                        dx = thumb_tip[0] - index_tip[0]
                        dy = thumb_tip[1] - index_tip[1]
                        distance = np.sqrt(dx*dx + dy*dy)
                        
                        # Wide spread indicates zoom in
                        if distance > 100:  # Large distance = zoom in
                            zoom_level = min(zoom_level + zoom_change, max_zoom)
                            last_gesture_time = current_time
                            print(f"Zoom in: {zoom_level:.1f}x")
                
                
                # Zoom Out - Spread gesture (thumb + index + spread)
                if fingers == [0, 0, 1, 1, 1] and can_perform_gesture:
                    # Check if thumb and index are spread apart
                    if lm_list[4] and lm_list[8]:  # Thumb tip and index tip
                        # Calculate distance
                        thumb_tip = lm_list[4]
                        index_tip = lm_list[8]
                        dx = thumb_tip[0] - index_tip[0]
                        dy = thumb_tip[1] - index_tip[1]
                        distance = np.sqrt(dx*dx + dy*dy)
                        
                        # Close pinch indicates zoom out
                        zoom_level = max(zoom_level - zoom_change, min_zoom)
                        last_gesture_time = current_time
                        print(f"Zoom out: {zoom_level:.1f}x")

                # Reset Zoom - Three finger pinch
                if fingers == [1, 1, 1, 0, 0] and can_perform_gesture:
                    zoom_level = 1.0
                    last_gesture_time = current_time
                    print("Zoom reset to 1.0x")

                # Jump to first slide
                if fingers == [0,1,0,0,1]:
                    slide_num = 0
                    annotations = [[]]
                    annot_num = 0
                    gest_done = True

                # Jump to last slide
                if fingers == [0,0,0,0,1]:
                    slide_num = len(slide_images) - 1
                    annotations = [[]]
                    annot_num = 0
                    gest_done = True

            # Show Pointer
            if fingers == [0, 1, 1, 0, 0]:
                # Make pointer more visible with a larger circle and contrasting colors
                cv2.circle(slide_current, index_fing, 5, (0, 0, 255), cv2.FILLED)  # Red fill
                cv2.circle(slide_current, index_fing, 6, (0, 0, 0), 2)  # Black outline
                
                
                annot_start = False

            # Draw
            if fingers == [0, 1, 0, 0, 0]:
                # Process with cooldown for smoother performance
                if not annot_start:
                    annot_start = True
                    annot_num += 1
                    annotations.append([])

                # Add point to annotation
                annotations[annot_num].append(index_fing)
                
                # Draw more visible drawing point
                cv2.circle(slide_current, index_fing, 4, (0, 0, 255), cv2.FILLED, cv2.LINE_AA)

            else:
                annot_start = False

            # Erase
            if fingers == [0, 1, 1, 1, 0] and can_perform_gesture:
                if annotations and annot_num >= 0:
                    if len(annotations) > 1:  # Make sure we have at least one annotation besides the initial empty one
                        annotations.pop(-1)
                        annot_num -= 1
                        gest_done = True
                        last_gesture_time = current_time

        else:
            annot_start = False

        # Gesture Performed Iterations:
        if gest_done:
            gest_counter += 1
            if gest_counter > delay:
                gest_counter = 0
                gest_done = False

        # Draw Annotations with improved visual style
        for annotation in annotations:
            for j in range(1, len(annotation)):
                # Draw smoother lines with anti-aliasing
                cv2.line(slide_current, annotation[j - 1], annotation[j], (0, 0, 255), 4, cv2.LINE_AA)
                
                # Add shadow effect for depth
                shadow_p1 = (annotation[j - 1][0] + 3, annotation[j - 1][1] + 3)
                shadow_p2 = (annotation[j][0] + 3, annotation[j][1] + 3)
                cv2.line(slide_current, shadow_p1, shadow_p2, (100, 100, 100), 4, cv2.LINE_AA)
               

        # Adding Camera Image on Slide
        img_small = cv2.resize(frame, (ws, hs))
        h, w, _ = slide_current.shape
        slide_current[h - hs:h, w - ws:w] = img_small

        # Add slide counter
        slide_text = f"Slide {slide_num + 1}/{len(slide_images)}"
        cv2.putText(slide_current, slide_text, (10, 30), cv2.FONT_HERSHEY_SIMPLEX, 0.7, (0, 0, 0), 2)

        cv2.imshow("Slides", slide_current)

        key = cv2.waitKey(1)
        if key == ord('q'):
            break
        elif key == 27:  # ESC key to toggle fullscreen
            is_fullscreen = not is_fullscreen
            cv2.setWindowProperty("Slides", cv2.WND_PROP_FULLSCREEN, 
                                 cv2.WINDOW_FULLSCREEN if is_fullscreen else cv2.WINDOW_NORMAL)
        elif key == ord('o'):  # o for open
            cap.release()
            cv2.destroyAllWindows()
            file_path = file_manager.show_file_dialog()
            
            # Clear existing slides and annotations
            slide_images = []
            annotations = [[]]
            annot_num = 0
            slide_num = 0
            
            if file_path:
                if file_path.lower().endswith(".pdf"):
                    try:
                        doc = fitz.open(file_path)
                        for page in doc:
                            # Increase the resolution of rendered pages
                            zoom = 2  # Increase zoom factor for better quality
                            mat = fitz.Matrix(zoom, zoom)
                            pix = page.get_pixmap(matrix=mat, alpha=False)
                            
                            # Convert to numpy array
                            img = np.frombuffer(pix.samples, dtype=np.uint8).reshape(pix.h, pix.w, 3)
                            
                            # Convert from RGB to BGR (OpenCV format)
                            img = cv2.cvtColor(img, cv2.COLOR_RGB2BGR)
                            
                            # Resize to match window dimensions
                            img = cv2.resize(img, (width, height))
                            
                            slide_images.append(img)
                        doc.close()  # Close the PDF file
                    except Exception as e:
                        print(f"Error loading PDF file: {e}")
                        blank = np.ones((height, width, 3), dtype=np.uint8) * 255
                        cv2.putText(blank, "Error loading PDF", (width//2 - 150, height//2), 
                                   cv2.FONT_HERSHEY_SIMPLEX, 1, (0, 0, 255), 2)
                        slide_images.append(blank)
                
                elif file_path.lower().endswith(".pptx"):
                    try:
                        print("Rendering PowerPoint slides as images...")
                        # Create a temporary directory to store slide images
                        temp_dir = os.path.join(os.getcwd(), "temp_slides")
                        if not os.path.exists(temp_dir):
                            os.makedirs(temp_dir)

                        # Initialize PowerPoint application
                        powerpoint = comtypes.client.CreateObject("PowerPoint.Application")
                        powerpoint.Visible = 1

                        # Open the presentation
                        file_path = os.path.abspath(file_path)  # Ensure the file path is absolute
                        presentation = powerpoint.Presentations.Open(file_path, WithWindow=False)

                        # Export each slide as an image
                        for slide_index in range(1, presentation.Slides.Count + 1):
                            slide_path = os.path.join(temp_dir, f"slide_{slide_index}.jpg")
                            presentation.Slides[slide_index].Export(slide_path, "JPG", width, height)
                            # Load the exported image into OpenCV
                            slide_img = cv2.imread(slide_path)
                            if slide_img is not None:
                                slide_images.append(slide_img)
                            else:
                                print(f"Warning: Slide {slide_index} could not be loaded as an image.")

                        print(f"Successfully rendered {len(slide_images)} slides from PowerPoint")
                    except Exception as e:
                        print(f"Error rendering PowerPoint slides: {e}")
                        # Create a blank slide with error message
                        blank = np.ones((height, width, 3), dtype=np.uint8) * 255
                        cv2.putText(blank, "Error rendering PowerPoint", (width // 2 - 150, height //2),
                                    cv2.FONT_HERSHEY_SIMPLEX, 1, (0, 0, 255), 2)
                        slide_images.append(blank)
                    finally:
                        # Ensure PowerPoint application is closed
                        try:
                            presentation.Close()
                            powerpoint.Quit()
                        except Exception as cleanup_error:
                            print(f"Error during PowerPoint cleanup: {cleanup_error}")

                        # Clean up temporary files
                        if os.path.exists(temp_dir):
                            for file in os.listdir(temp_dir):
                                os.remove(os.path.join(temp_dir, file))
                            os.rmdir(temp_dir)
            else:
                blank = np.ones((height, width, 3), dtype=np.uint8) * 255
                cv2.putText(blank, "No presentation loaded", (width//2 - 150, height//2), 
                           cv2.FONT_HERSHEY_SIMPLEX, 1, (0, 0, 0), 2)
                slide_images.append(blank)
                
            # Restart camera
            cap = cv2.VideoCapture(0)
            cap.set(3, width)
            cap.set(4, height)
        
        elif key == ord('h'):  # h for help
            # Display help
            help_img = np.ones((height, width, 3), dtype=np.uint8) * 255
            cv2.putText(help_img, "Hand Gesture Controls", (width//2 - 150, 50), 
                       cv2.FONT_HERSHEY_SIMPLEX, 1, (0, 0, 0), 2)
            
            # Add gesture instructions
            instructions = [
                "Thumb only [1,0,0,0,0] - Previous Slide",
                "Four fingers (no thumb) [0,1,1,1,1] - Next Slide",
                "All five fingers [1,1,1,1,1] - Clear Annotations",
                "Index finger only [0,1,0,0,0] - Draw",
                "Index + Middle fingers [0,1,1,0,0] - Pointer",
                "Index + Middle + Ring [0,1,1,1,0] - Erase Last Annotation",
                "Thumb + Index far apart - Zoom In",
                "Middle+ Ring + Pinky finger [0,0,1,1,1] - Zoom Out",
                "Thumb + Index + Middle - Reset Zoom",
                "Index + Pinky finger - First Slide",
                "Pinky finger - Last Slide"
                "",
                "KEYBOARD CONTROLS:",
                "ESC - Toggle fullscreen mode",
                "Q - Quit application",
                "O - Open file",
                "H - Help dialog"
            ]
            
            for i, inst in enumerate(instructions):
                cv2.putText(help_img, inst, (100, 120 + i*40), 
                           cv2.FONT_HERSHEY_SIMPLEX, 0.7, (0, 0, 0), 1)
            
            cv2.putText(help_img, "Press any key to return", (width//2 - 150, height - 50), 
                       cv2.FONT_HERSHEY_SIMPLEX, 0.7, (100, 100, 100), 1)
            
            cv2.imshow("Help", help_img)
            cv2.waitKey(0)
            cv2.destroyWindow("Help")

    cap.release()
    cv2.destroyAllWindows()
    print("Application closed")

if __name__ == "__main__":
    show_splash_screen(run_main_application)