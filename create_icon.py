"""
Create a modern icon from the DecoPress logo.
This will generate an .ico file with multiple sizes.
"""
import os
from PIL import Image, ImageDraw, ImageOps, ImageFilter

# Ensure PIL is installed
try:
    from PIL import Image
except ImportError:
    print("Pillow not installed. Run: pip install pillow")
    exit(1)

def create_circular_icon(input_image_path, output_ico_path, sizes=[16, 32, 48, 64, 128, 256]):
    """
    Create a circular icon with multiple sizes from an input image.
    Adds a subtle shadow and modern styling.
    """
    print(f"Creating icon from {input_image_path}...")
    
    # Check if input image exists
    if not os.path.exists(input_image_path):
        print(f"Error: Input image {input_image_path} not found")
        return False
    
    try:
        # Open the input image
        image = Image.open(input_image_path)
        
        # Create multiple sized icons
        icons = []
        for size in sizes:
            # Create a new image with transparent background
            icon = Image.new('RGBA', (size, size), (0, 0, 0, 0))
            
            # Resize the input image proportionally
            img_resized = image.copy()
            img_resized.thumbnail((size * 0.8, size * 0.8), Image.LANCZOS)
            
            # Determine position to paste the resized image (center)
            paste_x = (size - img_resized.width) // 2
            paste_y = (size - img_resized.height) // 2
            
            # Create a mask for rounded corners
            mask = Image.new('L', img_resized.size, 0)
            draw = ImageDraw.Draw(mask)
            
            # Calculate the radius for rounded corners (adjust as needed)
            radius = min(img_resized.width, img_resized.height) // 4
            
            # Draw a rounded rectangle
            draw.rounded_rectangle([(0, 0), (img_resized.width, img_resized.height)], 
                                radius=radius, fill=255)
            
            # Apply mask to create rounded corners
            img_rounded = Image.new('RGBA', img_resized.size, (0, 0, 0, 0))
            img_rounded.paste(img_resized, (0, 0), mask)
            
            # Add a subtle shadow
            shadow = Image.new('RGBA', img_rounded.size, (0, 0, 0, 0))
            shadow_draw = ImageDraw.Draw(shadow)
            shadow_draw.rounded_rectangle([(2, 2), (img_rounded.width+2, img_rounded.height+2)], 
                                      radius=radius, fill=(0, 0, 0, 80))
            shadow = shadow.filter(ImageFilter.GaussianBlur(2))
            
            # Paste shadow and then image onto the icon
            icon.paste(shadow, (paste_x-1, paste_y-1), shadow)
            icon.paste(img_rounded, (paste_x, paste_y), img_rounded)
            
            icons.append(icon)
        
        # Save as .ico file with multiple sizes
        icons[0].save(output_ico_path, format='ICO', sizes=[(size, size) for size in sizes], 
                     append_images=icons[1:])
        
        print(f"Icon created successfully: {output_ico_path}")
        return True
    
    except Exception as e:
        print(f"Error creating icon: {str(e)}")
        return False

if __name__ == "__main__":
    # Get the current directory
    script_dir = os.path.dirname(os.path.abspath(__file__))
    
    # Input and output paths
    logo_path = os.path.join(script_dir, "DecoPressLogo.jpg")
    icon_path = os.path.join(script_dir, "DecoPressLogo.ico")
    
    # Create the icon
    create_circular_icon(logo_path, icon_path) 