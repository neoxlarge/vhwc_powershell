def add_watermark(image_path, text, output_path):
  """
  在 PNG 檔中加入文字浮水印

  Args:
    image_path: 要加入浮水印的 PNG 檔路徑
    text: 浮水印文字
    output_path: 加入浮水印後的 PNG 檔輸出路徑

  Returns:
    None
  """

  # 載入 PNG 檔
  image = Image.open(image_path)

  # 建立文字浮水印
  watermark = Image.new("RGBA", image.size, (0, 0, 0, 0))
  draw = ImageDraw.Draw(watermark)
  text_size = draw.textsize(text, font=ImageFont.truetype("arial.ttf", 40))

  # 將文字浮水印放在圖片最下方
  position = (image.size[0] - text_size[0], image.size[1] - text_size[1])
  draw.text(position, text, font=ImageFont.truetype("arial.ttf", 40), fill=(255, 255, 255, 128))

  # 將文字浮水印加入 PNG 檔
  image.paste(watermark, position, watermark)

  # 儲存加入浮水印後的 PNG 檔
  image.save(output_path)


# 範例
add_watermark("c:\\temp\\vhwc_showjob_20240228_2.png", "Copyright © 2024", "c:\\temp\\vhwc_2.png")
