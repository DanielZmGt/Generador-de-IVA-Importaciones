from PIL import Image

try:
    img = Image.open("logo 3zg IVA IMP.jpg")
    img.save("logo.ico", format="ICO", sizes=[(256, 256), (128, 128), (64, 64), (48, 48), (32, 32), (16, 16)])
    print("Successfully created logo.ico")
except Exception as e:
    print(f"Error creating icon: {e}")
