from PIL import Image, ImageDraw

SIZES = (16, 24, 32, 48, 64, 128, 256)
CANVAS = 256
image = Image.new("RGBA", (CANVAS, CANVAS), "#142B3A")
draw = ImageDraw.Draw(image)
draw.rounded_rectangle((28, 28, 228, 228), radius=48, outline="#D9E7EA", width=8)
for y, opacity in ((88, 255), (128, 170), (168, 90)):
    color = (112, 199, 191, opacity)
    draw.rounded_rectangle((72, y - 7, 184, y + 7), radius=7, fill=color)
draw.rounded_rectangle((121, 66, 135, 190), radius=7, fill=(112, 199, 191, 235))
image.save("assets/app.ico", sizes=[(size, size) for size in SIZES])
image.save("assets/app-icon.png")
