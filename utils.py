import cv2
import numpy as np

def draw_dotted_rect(img, pt1, pt2, color, thickness=1, style='dotted', gap=10):
    """
    Draw a dotted rectangle on an image.
    
    Args:
        img: The image to draw on
        pt1: First point (top-left corner) as a tuple (x, y)
        pt2: Second point (bottom-right corner) as a tuple (x, y)
        color: Line color as a BGR tuple (Blue, Green, Red)
        thickness: Line thickness
        style: Line style ('dotted' or 'dashed')
        gap: Gap between dots or dashes
    
    Returns:
        The image with the dotted rectangle
    """
    x1, y1 = pt1
    x2, y2 = pt2
    
    # Draw the horizontal lines
    for x in range(x1, x2, gap):
        cv2.line(img, (x, y1), (min(x + gap//2, x2), y1), color, thickness)
        cv2.line(img, (x, y2), (min(x + gap//2, x2), y2), color, thickness)
    
    # Draw the vertical lines
    for y in range(y1, y2, gap):
        cv2.line(img, (x1, y), (x1, min(y + gap//2, y2)), color, thickness)
        cv2.line(img, (x2, y), (x2, min(y + gap//2, y2)), color, thickness)
    
    return img

def resize_image_aspect_ratio(image, target_width=None, target_height=None):
    """
    Resize an image while maintaining its aspect ratio.
    
    Args:
        image: The image to resize
        target_width: The target width (if None, will be calculated from target_height)
        target_height: The target height (if None, will be calculated from target_width)
    
    Returns:
        The resized image
    """
    height, width = image.shape[:2]
    
    # If both dimensions are None, return the original image
    if target_width is None and target_height is None:
        return image
    
    # Calculate the ratio of the original image
    ratio = width / height
    
    # If target_width is None, calculate it from target_height
    if target_width is None:
        target_width = int(target_height * ratio)
    
    # If target_height is None, calculate it from target_width
    if target_height is None:
        target_height = int(target_width / ratio)
    
    # Resize the image
    resized = cv2.resize(image, (target_width, target_height), interpolation=cv2.INTER_AREA)
    
    return resized

def put_text_with_background(img, text, position, font=cv2.FONT_HERSHEY_SIMPLEX, 
                            font_scale=1, color=(255, 255, 255), bg_color=(0, 0, 0),
                            thickness=2, padding=10):
    """
    Put text with background on an image.
    
    Args:
        img: The image to draw on
        text: The text to display
        position: Position tuple (x, y) where x, y is the bottom-left corner of the text
        font: OpenCV font
        font_scale: Font scale
        color: Text color as BGR tuple
        bg_color: Background color as BGR tuple
        thickness: Text thickness
        padding: Padding around the text
        
    Returns:
        The image with text and background
    """
    # Get text size
    text_size = cv2.getTextSize(text, font, font_scale, thickness)[0]
    
    # Calculate text box dimensions
    text_width, text_height = text_size
    x, y = position
    
    # Draw background rectangle
    cv2.rectangle(img, 
                 (x - padding, y - text_height - padding), 
                 (x + text_width + padding, y + padding), 
                 bg_color, 
                 -1)  # -1 fills the rectangle
    
    # Draw text
    cv2.putText(img, text, (x, y), font, font_scale, color, thickness)
    
    return img

def create_blank_image(width, height, color=(255, 255, 255)):
    """
    Create a blank image with the specified dimensions and color.
    
    Args:
        width: Width of the image
        height: Height of the image
        color: Color of the image as a BGR tuple
        
    Returns:
        A blank image
    """
    img = np.ones((height, width, 3), dtype=np.uint8)
    img[:] = color
    return img
import cv2
import numpy as np

def draw_dotted_rect(img, pt1, pt2, color, thickness=1, style='dotted', gap=10):
    """
    Draw a dotted rectangle on an image.
    
    Args:
        img: The image to draw on
        pt1: First point (top-left corner) as a tuple (x, y)
        pt2: Second point (bottom-right corner) as a tuple (x, y)
        color: Line color as a BGR tuple (Blue, Green, Red)
        thickness: Line thickness
        style: Line style ('dotted' or 'dashed')
        gap: Gap between dots or dashes
    
    Returns:
        The image with the dotted rectangle
    """
    x1, y1 = pt1
    x2, y2 = pt2
    
    # Draw the horizontal lines
    for x in range(x1, x2, gap):
        cv2.line(img, (x, y1), (min(x + gap//2, x2), y1), color, thickness)
        cv2.line(img, (x, y2), (min(x + gap//2, x2), y2), color, thickness)
    
    # Draw the vertical lines
    for y in range(y1, y2, gap):
        cv2.line(img, (x1, y), (x1, min(y + gap//2, y2)), color, thickness)
        cv2.line(img, (x2, y), (x2, min(y + gap//2, y2)), color, thickness)
    
    return img

def resize_image_aspect_ratio(image, target_width=None, target_height=None):
    """
    Resize an image while maintaining its aspect ratio.
    
    Args:
        image: The image to resize
        target_width: The target width (if None, will be calculated from target_height)
        target_height: The target height (if None, will be calculated from target_width)
    
    Returns:
        The resized image
    """
    height, width = image.shape[:2]
    
    # If both dimensions are None, return the original image
    if target_width is None and target_height is None:
        return image
    
    # Calculate the ratio of the original image
    ratio = width / height
    
    # If target_width is None, calculate it from target_height
    if target_width is None:
        target_width = int(target_height * ratio)
    
    # If target_height is None, calculate it from target_width
    if target_height is None:
        target_height = int(target_width / ratio)
    
    # Resize the image
    resized = cv2.resize(image, (target_width, target_height), interpolation=cv2.INTER_AREA)
    
    return resized

def put_text_with_background(img, text, position, font=cv2.FONT_HERSHEY_SIMPLEX, 
                            font_scale=1, color=(255, 255, 255), bg_color=(0, 0, 0),
                            thickness=2, padding=10):
    """
    Put text with background on an image.
    
    Args:
        img: The image to draw on
        text: The text to display
        position: Position tuple (x, y) where x, y is the bottom-left corner of the text
        font: OpenCV font
        font_scale: Font scale
        color: Text color as BGR tuple
        bg_color: Background color as BGR tuple
        thickness: Text thickness
        padding: Padding around the text
        
    Returns:
        The image with text and background
    """
    # Get text size
    text_size = cv2.getTextSize(text, font, font_scale, thickness)[0]
    
    # Calculate text box dimensions
    text_width, text_height = text_size
    x, y = position
    
    # Draw background rectangle
    cv2.rectangle(img, 
                 (x - padding, y - text_height - padding), 
                 (x + text_width + padding, y + padding), 
                 bg_color, 
                 -1)  # -1 fills the rectangle
    
    # Draw text
    cv2.putText(img, text, (x, y), font, font_scale, color, thickness)
    
    return img

def create_blank_image(width, height, color=(255, 255, 255)):
    """
    Create a blank image with the specified dimensions and color.
    
    Args:
        width: Width of the image
        height: Height of the image
        color: Color of the image as a BGR tuple
        
    Returns:
        A blank image
    """
    img = np.ones((height, width, 3), dtype=np.uint8)
    img[:] = color
    return img