from selenium import webdriver
import pandas as pd


from PIL import Image

def crop_image(image_path, crop_length):
    # 檢查圖片格式
    if not image_path.lower().endswith(".png"):
        raise ValueError("Image format must be PNG")

    # 讀取圖片
    img = Image.open(image_path)

    # 計算圖片的寬度和高度
    width, height = img.size

    # 計算切片的數量
    num_crops = (height // crop_length) + 1

    filepaths = []
    # 進行切片
    for i in range(num_crops):
        # 計算切片的左上角座標
        x = 0
        y = i * crop_length

        # 計算切片的右下角座標
        w = width
        h = y + crop_length

        # 進行切片
        cropped_img = img.crop((x, y, w, h))

        # 儲存切片
        filepath = f"{image_path[:-4]}_{i + 1}.png"
        cropped_img.save(filepath)
        filepaths.append(filepath)
        
    return filepaths


# 使用範例
x = crop_image(r"c:\temp\vhwc_showjob_20240228.png", 2000)

print(x)



