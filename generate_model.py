import os
import FreeCAD as App
import Part
import Import

# Find the project and output folders
project_dir = os.path.dirname(os.path.abspath(__file__))
output_dir = os.path.join(project_dir, "demo_dataset")
os.makedirs(output_dir, exist_ok=True)

iges_path = os.path.join(output_dir, "2280_model.iges")
freecad_path = os.path.join(output_dir, "2280_model.FCStd")

# Create a new FreeCAD document
doc = App.newDocument("Part_2280")

# Main plate: 100 x 60 x 10 mm
shape = Part.makeBox(100, 60, 10)

# Four diameter-8 mm through-holes
corner_holes = [
    (10, 10),
    (90, 10),
    (10, 50),
    (90, 50),
]

for x, y in corner_holes:
    hole = Part.makeCylinder(
        4,                       # Radius = 4 mm
        10,                      # Through the 10 mm thickness
        App.Vector(x, y, 0)
    )
    shape = shape.cut(hole)

# One diameter-20 mm centre through-hole
center_hole = Part.makeCylinder(
    10,                          # Radius = 10 mm
    10,
    App.Vector(50, 30, 0)
)
shape = shape.cut(center_hole)

# Remove unnecessary split edges where possible
shape = shape.removeSplitter()

# Add final part to document
part = doc.addObject("PartDesign::Feature", "CNC_Mounting_Plate")
part.Label = "Part 2280 - CNC Mounting Plate"
part.Shape = shape

# Add useful metadata
part.addProperty("App::PropertyString", "PartNumber", "Manufacturing")
part.PartNumber = "2280"

part.addProperty("App::PropertyString", "Material", "Manufacturing")
part.Material = "Aluminium 6061-T6"

part.addProperty("App::PropertyString", "Revision", "Manufacturing")
part.Revision = "A"

part.addProperty("App::PropertyString", "Process", "Manufacturing")
part.Process = "CNC milling and drilling"

doc.recompute()

# Save editable FreeCAD model
doc.saveAs(freecad_path)

# Export IGES model
Import.export([part], iges_path)

# Report geometry facts
box = shape.BoundBox

print("")
print("3D model generated successfully")
print("IGES file:", iges_path)
print("FreeCAD file:", freecad_path)
print(
    "Bounding box: {:.2f} x {:.2f} x {:.2f} mm".format(
        box.XLength,
        box.YLength,
        box.ZLength
    )
)
print("Volume: {:.2f} cubic mm".format(shape.Volume))
print("Surface area: {:.2f} square mm".format(shape.Area))
print("Solids:", len(shape.Solids))
print("Faces:", len(shape.Faces))
print("Edges:", len(shape.Edges))

App.closeDocument("Part_2280")