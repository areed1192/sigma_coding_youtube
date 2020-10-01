import win32com.client as win32

# Grab the Active Instance of Excel.
adobe_app = win32.GetActiveObject("Illustrator.Application")

# Define the Document we will be working with.
adobe_file = adobe_app.ActiveDocument

# Let's start with my Layer, that should contain my other objects.
thumbnail_layer = adobe_file.Layers("ThumbnailVideo")

# Inside my layer I have a TextFrame for both my Video Title.
video_title_frame = thumbnail_layer.TextFrames("VideoTitleFrame")

# and Video Series...
video_series_frame = thumbnail_layer.TextFrames("VideoSeriesFrame")

# I also have a PathItem that represents my Background.
thumbnail_background_path_item = thumbnail_layer.PathItems(
    "ThumbnailBackground"
)

# Grab the Series Text.
video_series_text = video_series_frame.TextRange.Contents

# Grab the Video Text.
video_title_text = video_title_frame.TextRange.Contents

# Inside of Adobe we have Colors, let's create a new Color that will represent our Background.
black_background_color = win32.Dispatch("Illustrator.RGBColor")
black_background_color.Red = 25
black_background_color.Green = 24
black_background_color.Blue = 24

# Define the Export PNG24 Options.
png_export_options = win32.Dispatch("Illustrator.ExportOptionsPNG24")
png_export_options.AntiAliasing = True
png_export_options.Transparency = True
png_export_options.MatteColor = black_background_color

# Export the document.
adobe_file.Export(
    ExportFile=r"C:\Users\Alex\OneDrive\Growth - Tutorial Videos\{file_name}".format(
        file_name="my_first_png"
    ),
    ExportFormat=5,
    Options=png_export_options
)
