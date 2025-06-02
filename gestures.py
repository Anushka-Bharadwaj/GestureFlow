import numpy as np

class GestureRecognizer:
    """
    Recognizes hand gestures for controlling the presentation.
    """
    
    def __init__(self, gesture_threshold_x=650, gesture_threshold_y=300, 
                gesture_delay=15):
        """
        Initialize the gesture recognizer with thresholds.
        
        Args:
            gesture_threshold_x: X-coordinate threshold for gesture detection area
            gesture_threshold_y: Y-coordinate threshold for gesture detection area
            gesture_delay: Delay frames between gesture detection to prevent repeated triggers
        """
        self.gesture_threshold_x = gesture_threshold_x
        self.gesture_threshold_y = gesture_threshold_y
        self.gesture_delay = gesture_delay
        self.gesture_cooldown = 0
        self.gesture_active = False
        
    def is_in_gesture_area(self, hand):
        """
        Check if a hand is in the gesture detection area (upper right corner).
        
        Args:
            hand: Hand dictionary from HandDetector
        
        Returns:
            True if the hand is in the gesture area, False otherwise
        """
        if not hand:
            return False
            
        cx, cy = hand["center"]
        return (cy < self.gesture_threshold_y and cx > self.gesture_threshold_x)
    
    def recognize_gesture(self, hand, fingers):
        """
        Recognize a gesture based on finger positions.
        
        Args:
            hand: Hand dictionary from HandDetector
            fingers: List of finger states (0 or 1) from HandDetector.fingersUp()
            
        Returns:
            Tuple of (action, should_trigger)
            action: String describing the recognized gesture, or None
            should_trigger: Boolean indicating if the action should be triggered
        """
        # Check if we're in cooldown period
        if self.gesture_cooldown > 0:
            self.gesture_cooldown -= 1
            return None, False
            
        # Check if hand is in gesture area
        if not self.is_in_gesture_area(hand):
            return None, False
            
        # Recognize gestures and return the action
        action = None
        should_trigger = False
        
        # Next Slide (0,1,1,1,1) - All fingers except thumb
        if fingers == [0, 1, 1, 1, 1]:
            action = "NEXT_SLIDE"
            should_trigger = True
            
        # Previous Slide (1,0,0,0,0) - Only thumb
        elif fingers == [1, 0, 0, 0, 0]:
            action = "PREV_SLIDE"
            should_trigger = True
            
        # Clear Annotations (1,1,1,1,1) - All five fingers
        elif fingers == [1, 1, 1, 1, 1]:
            action = "CLEAR_ANNOTATIONS"
            should_trigger = True
        
        # Pointer Mode (0,1,1,0,0) - Index and middle finger
        elif fingers == [0, 1, 1, 0, 0]:
            action = "POINTER"
            should_trigger = False  # Pointer doesn't need cooldown
            
        # Draw Mode (0,1,0,0,0) - Only index finger
        elif fingers == [0, 1, 0, 0, 0]:
            action = "DRAW"
            should_trigger = False  # Drawing doesn't need cooldown
            
        # Erase Last (0,1,1,1,0) - Index, middle, and ring fingers
        elif fingers == [0, 1, 1, 1, 0]:
            action = "ERASE"
            should_trigger = True
            
        # If we should trigger an action, start the cooldown
        if should_trigger:
            self.gesture_cooldown = self.gesture_delay
            
        return action, should_trigger
        
    def get_pointer_position(self, hand, width, height):
        """
        Get the position of the pointer (index finger tip).
        
        Args:
            hand: Hand dictionary from HandDetector
            width: Width of the frame
            height: Height of the frame
            
        Returns:
            The (x, y) position of the pointer on the screen
        """
        if not hand:
            return None
            
        lm_list = hand["lmList"]
        if not lm_list or len(lm_list) < 9:
            return None
            
        # Get index finger tip position (landmark 8)
        x_val = int(np.interp(lm_list[8][0], [width // 2, width], [0, width]))
        y_val = int(np.interp(lm_list[8][1], [150, height - 150], [0, height]))
        
        return (x_val, y_val)