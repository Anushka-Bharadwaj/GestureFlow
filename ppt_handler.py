import os
import cv2
import numpy as np
import comtypes.client

def load_powerpoint(file_path, width, height):
    slide_images = []
    try:
        print("Rendering PowerPoint slides as images...")
        temp_dir = os.path.join(os.getcwd(), "temp_slides")
        if not os.path.exists(temp_dir):
            os.makedirs(temp_dir)

        powerpoint = comtypes.client.CreateObject("PowerPoint.Application")
        powerpoint.Visible = 1

        file_path = os.path.abspath(file_path)
        presentation = powerpoint.Presentations.Open(file_path, WithWindow=False)

        for slide_index in range(1, presentation.Slides.Count + 1):
            slide_path = os.path.join(temp_dir, f"slide_{slide_index}.jpg")
            presentation.Slides[slide_index].Export(slide_path, "JPG", width, height)
            slide_img = cv2.imread(slide_path)
            if slide_img is not None:
                slide_images.append(slide_img)

        print(f"Successfully rendered {len(slide_images)} slides from PowerPoint")
        
    except Exception as e:
        print(f"Error rendering PowerPoint slides: {e}")
        blank = np.ones((height, width, 3), dtype=np.uint8) * 255
        cv2.putText(blank, "Error rendering PowerPoint", (width // 2 - 150, height // 2),
                    cv2.FONT_HERSHEY_SIMPLEX, 1, (0, 0, 255), 2)
        slide_images.append(blank)
        
    finally:
        try:
            presentation.Close()
            powerpoint.Quit()
        except:
            pass

        if os.path.exists(temp_dir):
            for file in os.listdir(temp_dir):
                os.remove(os.path.join(temp_dir, file))
            os.rmdir(temp_dir)
            
    return slide_images