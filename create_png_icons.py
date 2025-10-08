#!/usr/bin/env python3
"""
Direct PNG Icon Creator for Outlook Add-in
Creates PNG icons directly using PIL
"""

import os
import sys
from pathlib import Path

def create_png_icons():
    """Create PNG icons directly"""
    
    try:
        from PIL import Image, ImageDraw
    except ImportError:
        print("PIL (Pillow) not installed. Installing...")
        os.system("pip install Pillow")
        try:
            from PIL import Image, ImageDraw
        except ImportError:
            print("Failed to install PIL. Please install manually: pip install Pillow")
            return False
    
    # Check if we're in the right directory
    if not os.path.exists('outlook_addin'):
        print("Error: outlook_addin directory not found!")
        print("Please run this script from the project root directory.")
        return False
    
    assets_dir = Path('outlook_addin/assets')
    assets_dir.mkdir(exist_ok=True)
    
    print("Creating PNG icons directly...")
    
    def create_brain_icon(size, filename):
        """Create a brain icon with AI elements"""
        img = Image.new('RGBA', (size, size), (0, 0, 0, 0))  # Transparent background
        draw = ImageDraw.Draw(img)
        
        # Colors
        bg_color = (76, 175, 80, 255)  # Green background
        brain_color = (129, 199, 132, 255)  # Light green brain
        outline_color = (46, 125, 50, 255)  # Dark green outline
        sparkle_color = (255, 215, 0, 255)  # Gold sparkle
        
        # Background circle
        margin = size // 16
        draw.ellipse([margin, margin, size - margin, size - margin], 
                    fill=bg_color, outline=outline_color, width=size//16)
        
        # Brain shape (simplified)
        brain_margin = size // 4
        brain_points = [
            (size//2, brain_margin),
            (brain_margin, size//3),
            (brain_margin, 2*size//3),
            (size//2, size - brain_margin),
            (size - brain_margin, 2*size//3),
            (size - brain_margin, size//3)
        ]
        
        # Draw brain outline
        draw.polygon(brain_points, fill=brain_color, outline=outline_color, width=size//32)
        
        # Central dividing line
        draw.line([(size//2, brain_margin), (size//2, size - brain_margin)], 
                 fill=outline_color, width=size//32)
        
        # Brain texture lines
        for i in range(3):
            y = brain_margin + (i + 1) * (size - 2*brain_margin) // 4
            draw.line([(size//3, y), (2*size//3, y)], 
                     fill=outline_color, width=size//64)
        
        # Neuron dots
        dot_size = size // 20
        positions = [
            (size//2 - size//8, size//2 - size//8),
            (size//2 + size//8, size//2 - size//8),
            (size//2 - size//8, size//2 + size//8),
            (size//2 + size//8, size//2 + size//8)
        ]
        
        for pos in positions:
            draw.ellipse([pos[0] - dot_size, pos[1] - dot_size, 
                         pos[0] + dot_size, pos[1] + dot_size], 
                        fill=outline_color)
        
        # AI sparkles
        sparkle_size = size // 16
        sparkle_positions = [
            (3*size//4, size//4),
            (size//4, size//4),
            (3*size//4, 3*size//4),
            (size//4, 3*size//4)
        ]
        
        for pos in sparkle_positions:
            # Draw diamond sparkle
            sparkle_points = [
                (pos[0], pos[1] - sparkle_size),
                (pos[0] + sparkle_size, pos[1]),
                (pos[0], pos[1] + sparkle_size),
                (pos[0] - sparkle_size, pos[1])
            ]
            draw.polygon(sparkle_points, fill=sparkle_color, outline=outline_color, width=1)
        
        # Save the image
        img.save(filename, 'PNG')
        print(f"Created: {filename}")
    
    # Create both icon sizes
    create_brain_icon(32, assets_dir / 'icon-32.png')
    create_brain_icon(64, assets_dir / 'icon-64.png')
    
    print("PNG icons created successfully!")
    return True

if __name__ == "__main__":
    success = create_png_icons()
    if success:
        print("\nAll icons created successfully!")
    else:
        print("\nFailed to create icons!")
        sys.exit(1)






