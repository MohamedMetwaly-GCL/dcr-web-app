from PIL import Image

def remove_white_background(img_path):
    img = Image.open(img_path)
    img = img.convert("RGBA")
    datas = img.getdata()

    new_data = []
    # Using a threshold to catch near-white compression artifacts
    for item in datas:
        if item[0] > 230 and item[1] > 230 and item[2] > 230:
            # White-ish pixel -> fully transparent
            new_data.append((255, 255, 255, 0))
        else:
            # Keep original
            new_data.append(item)

    img.putdata(new_data)
    img.save(img_path, "PNG")
    print(f"Removed white background from {img_path}")

remove_white_background(r"D:\DCR\static\logo-login.png")
