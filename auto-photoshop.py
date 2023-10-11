import os
import win32com.client
from PIL import Image


SILENT_CLOSE = 2
curdir = os.path.abspath(os.path.dirname(__file__))
source_path = os.path.join(curdir, "photo_frame_by_the_houseplant_todo.psd")
gen_path = os.path.join(curdir, "final.jpg")


ps = win32com.client.Dispatch("Photoshop.Application")
ps.DisplayDialogs = 3
ps.Preferences.RulerUnits = 1

file = open('photo_frame_by_the_houseplant_todo_config.txt', 'r')
smartName = file.read()
print(smartName)

source_image = ps.Open(source_path)   #open example.psd
psdoc=ps.Application.ActiveDocument

for layer in source_image.Layers:
    if (layer.name=="Frame"):
        for sub_layer in layer.Layers:
            if (sub_layer.kind == 17):
                smart_object_layer = sub_layer

print(smart_object_layer.name)

psdoc.ActiveLayer = smart_object_layer

ps.ExecuteAction(ps.StringIDToTypeID("placedLayerEditContents"))
smartDoc = ps.activeDocument

for layer in smartDoc.Layers:
    for sub_layer in layer.Layers:
        if (sub_layer.name == smartName):
            smart_origin_layer=sub_layer

#smart_layer = smartDoc.LayerSets["PLACEHOLDER]"]
#smart_origin_layer = smart_layer.ArtLayers["Graphic Elements"]
smartDoc.ActiveLayer=smart_origin_layer

target_image_width = smart_origin_layer.Bounds[2]-smart_origin_layer.Bounds[0]
target_image_height = smart_origin_layer.Bounds[3]-smart_origin_layer.Bounds[1]

#print(target_image_height, target_image_width)

image = Image.open("demo.jpg")
resize_image = image.resize((int(target_image_width), int(target_image_height)))
resize_image.save("image.jpg")

insert_path = os.path.join(curdir, "image.jpg")


insert_image=ps.Open(insert_path)   #open demo.jpg
insert_image_layer=insert_image.ArtLayers.Item(1)


insert_image_layer.Copy()
insert_image.Close(SILENT_CLOSE)

pasted = smartDoc.Paste()
smartDoc.Save()
smartDoc.Close(SILENT_CLOSE)


jpgSaveOptions = win32com.client.Dispatch("Photoshop.JPEGSaveOptions")
source_image.Save()
source_image.SaveAs(gen_path, jpgSaveOptions, True, 2)


#source_image.Close(SILENT_CLOSE)
#ps.Quit()