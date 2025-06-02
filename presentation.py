import cv2
import numpy as np
import fitz  # PyMuPDF for handling PDFs
import os
import time
from utils import draw_dotted_rect, put_text_with_background

class PresentationController:
    """
    Controls the presentation display and interactions.
    """
    
    def __init__(self, width=1280, height=720):
        """
        Initialize the presentation controller.
        
        Args:
            width: Width of the presentation window
            height: Height of the presentation window
        """
        self.width = width
        self.height = height
        self.slide_images = []
        self.slide_num = 0
        self.annotations = [[]]
        self.annot_num = 0
        self.annot_start = False
        self.webcam_width = int(213 * 1.2)  # Width of the webcam preview
        self.webcam_height = int(120 * 1.2)  # Height of the webcam preview
        self.file_path = None
        self.file_name = None
        
    def load_file(self, file_path):
        """
        Load a file for presentation (PDF or PPT).
        
        Args:
            file_path: Path to the file to load
            
        Returns:
            True if the file was loaded successfully, False otherwise
        """
        print(f"Attempting to load file: {file_path}")
        self.file_path = file_path
        self.file_name = os.path.basename(file_path)
        
        # Clear any existing slides
        self.slide_images = []
        self.slide_num = 0
        self.annotations = [[]]
        self.annot_num = 0
        
        try:
            if file_path.lower().endswith(".pdf"):
                print("Loading PDF file...")
                # Load PDF slides
                doc = fitz.open(file_path)
                print(f"PDF loaded with {len(doc)} pages")
                
                for page_num, page in enumerate(doc):
                    print(f"Processing page {page_num+1}/{len(doc)}")
                    pix = page.get_pixmap(alpha=False)
                    img = np.frombuffer(pix.samples, dtype=np.uint8).reshape(pix.h, pix.w, 3)
                    # Convert from RGB to BGR (OpenCV format)
                    img = cv2.cvtColor(img, cv2.COLOR_RGB2BGR)
                    self.slide_images.append(img)
                
                print(f"Successfully loaded {len(self.slide_images)} slides from PDF")
                return True
                
            elif file_path.lower().endswith(".pptx"):
                print("Loading PowerPoint file...")
                # Create multiple placeholder slides for PPT
                for i in range(5):  # Create 5 sample slides
                    placeholder = np.ones((self.height, self.width, 3), dtype=np.uint8) * 255
                    
                    # Add slide number and title
                    title = f"PowerPoint Slide {i+1}"
                    font = cv2.FONT_HERSHEY_SIMPLEX
                    
                    # Add the title
                    cv2.putText(placeholder, title, (self.width//2 - 150, 100), 
                                font, 1.5, (0, 0, 0), 2)
                    
                    # Add the file name
                    cv2.putText(placeholder, self.file_name, (50, 50), 
                                font, 0.7, (100, 100, 100), 1)
                    
                    # Add some sample content based on slide number
                    if i == 0:
                        cv2.putText(placeholder, "Introduction", (self.width//2 - 100, 200), 
                                   font, 1, (0, 0, 100), 2)
                    elif i == 1:
                        cv2.putText(placeholder, "Key Features", (self.width//2 - 100, 200), 
                                   font, 1, (0, 0, 100), 2)
                    elif i == 2:
                        cv2.putText(placeholder, "Implementation", (self.width//2 - 100, 200), 
                                   font, 1, (0, 0, 100), 2)
                    elif i == 3:
                        cv2.putText(placeholder, "Results", (self.width//2 - 100, 200), 
                                   font, 1, (0, 0, 100), 2)
                    else:
                        cv2.putText(placeholder, "Conclusion", (self.width//2 - 100, 200), 
                                   font, 1, (0, 0, 100), 2)
                    
                    # Add some bullet points
                    for j in range(3):
                        y_pos = 300 + j * 60
                        cv2.putText(placeholder, f"• Bullet point {j+1}", (200, y_pos), 
                                   font, 0.8, (50, 50, 50), 1)
                    
                    self.slide_images.append(placeholder)
                
                print(f"Successfully created {len(self.slide_images)} placeholder slides for PowerPoint")
                return True
            
            else:
                print(f"Unsupported file format: {file_path}")
                return False
                
        except Exception as e:
            print(f"Error loading file: {e}")
            print(f"Exception details: {str(e)}")
            import traceback
            traceback.print_exc()
            return False
    
    def next_slide(self):
        """Move to the next slide if available."""
        if self.slide_num < len(self.slide_images) - 1:
            self.slide_num += 1
            self.clear_annotations()
            return True
        return False
    
    def prev_slide(self):
        """Move to the previous slide if available."""
        if self.slide_num > 0:
            self.slide_num -= 1
            self.clear_annotations()
            return True
        return False
    
    def clear_annotations(self):
        """Clear all annotations on the current slide."""
        self.annotations = [[]]
        self.annot_num = 0
        self.annot_start = False
        
    def erase_last_annotation(self):
        """Erase the last annotation on the current slide."""
        if self.annotations and len(self.annotations) > 1:
            self.annotations.pop()
            self.annot_num = max(0, self.annot_num - 1)
            return True
        return False
    
    def start_annotation(self):
        """Start a new annotation line."""
        if not self.annot_start:
            self.annot_start = True
            self.annot_num += 1
            self.annotations.append([])
    
    def add_annotation_point(self, point):
        """Add a point to the current annotation line."""
        if self.annot_start and point:
            self.annotations[self.annot_num].append(point)
    
    def end_annotation(self):
        """End the current annotation line."""
        self.annot_start = False
    
    def render_frame(self, webcam_frame, hand_data=None, show_debug=True, gesture_action=None):
        """
        Render a frame of the presentation with annotations and webcam.
        
        Args:
            webcam_frame: The webcam feed frame
            hand_data: Data about the detected hand for visualizing pointers
            show_debug: Whether to show debug information
            gesture_action: The current gesture action being performed
            
        Returns:
            The rendered frame
        """
        # Ensure hand_data has default values if it's None
        if hand_data is None:
            hand_data = {
                'gesture_thresh_x': 650,
                'gesture_thresh_y': 300,
                'pointer_pos': None
            }
        # Check if we have any slides loaded
        if not self.slide_images:
            # Create a blank image that says "No presentation loaded"
            blank = np.ones((self.height, self.width, 3), dtype=np.uint8) * 255
            
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
            
            # Resize and place webcam feed
            webcam_small = cv2.resize(webcam_frame, (self.webcam_width, self.webcam_height))
            h, w = blank.shape[:2]
            blank[h - self.webcam_height:h, w - self.webcam_width:w] = webcam_small
            
            return blank
        
        # Get the current slide and resize it to fit the window
        slide_current = cv2.resize(self.slide_images[self.slide_num], (self.width, self.height))
        
        # Draw all annotations
        for annotation in self.annotations:
            for j in range(1, len(annotation)):
                cv2.line(slide_current, annotation[j - 1], annotation[j], (0, 0, 255), 6)
        
        # If we have hand data and it includes a pointer position, draw it
        if hand_data and 'pointer_pos' in hand_data and hand_data['pointer_pos']:
            pointer_pos = hand_data['pointer_pos']
            cv2.circle(slide_current, pointer_pos, 8, (0, 0, 255), cv2.FILLED)
            cv2.circle(slide_current, pointer_pos, 12, (0, 0, 0), 2)
        
        # Add webcam feed
        webcam_small = cv2.resize(webcam_frame, (self.webcam_width, self.webcam_height))
        h, w = slide_current.shape[:2]
        slide_current[h - self.webcam_height:h, w - self.webcam_width:w] = webcam_small
        
        # Add slide counter
        slide_text = f"Slide {self.slide_num + 1}/{len(self.slide_images)}"
        cv2.putText(slide_current, slide_text, (10, 30), cv2.FONT_HERSHEY_SIMPLEX, 0.7, (0, 0, 0), 2)
        
        # Add filename
        if self.file_name:
            cv2.putText(slide_current, self.file_name, (10, 60), cv2.FONT_HERSHEY_SIMPLEX, 0.5, (100, 100, 100), 1)
        # Show current action if any
        if gesture_action:
            # Map action code to readable text
            action_text = gesture_action
            if gesture_action == "NEXT_SLIDE":
                action_text = "Next Slide"
            elif gesture_action == "PREV_SLIDE":
                action_text = "Previous Slide"
            elif gesture_action == "CLEAR_ANNOTATIONS":
                action_text = "Clear Annotations"
            elif gesture_action == "POINTER":
                action_text = "Pointer Mode"
            elif gesture_action == "DRAW":
                action_text = "Drawing Mode"
            elif gesture_action == "ERASE":
                action_text = "Erase Last"
            
            # Show the action text
            put_text_with_background(slide_current, action_text, (10, h - 40), 
                                   cv2.FONT_HERSHEY_SIMPLEX, 1, (255, 255, 255), 
                                   (0, 0, 200), 2, 10)
        
        return slide_current