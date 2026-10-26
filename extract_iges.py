import json
import os

import FreeCAD as App
import Part


# ------------------------------------------------------------
# File locations
# ------------------------------------------------------------

PROJECT_DIR = os.path.dirname(os.path.abspath(__file__))

# The application can override these with environment variables.
# Otherwise, the demonstration files are used.
INPUT_PATH = os.environ.get(
    "IGES_INPUT",
    os.path.join(PROJECT_DIR, "demo_dataset", "2280_model.iges")
)

OUTPUT_PATH = os.environ.get(
    "IGES_OUTPUT",
    os.path.join(PROJECT_DIR, "demo_dataset", "2280_geometry.json")
)

INPUT_PATH = os.path.abspath(INPUT_PATH)
OUTPUT_PATH = os.path.abspath(OUTPUT_PATH)


# ------------------------------------------------------------
# IGES analysis
# ------------------------------------------------------------

def analyze_iges(input_path):
    document = App.newDocument("IGES_Analysis")

    try:
        App.setActiveDocument(document.Name)

        # Import the IGES geometry into the FreeCAD document.
        Part.insert(input_path, document.Name)
        document.recompute()

        shapes = []

        for obj in document.Objects:
            if not hasattr(obj, "Shape"):
                continue

            shape = obj.Shape

            if shape is not None and not shape.isNull():
                shapes.append(shape)

        if not shapes:
            raise ValueError(
                "The IGES file was imported, but no usable geometry "
                "was detected."
            )

        if len(shapes) == 1:
            final_shape = shapes[0]
        else:
            final_shape = Part.makeCompound(shapes)

        bounding_box = final_shape.BoundBox
        solid_count = len(final_shape.Solids)
        volume = float(final_shape.Volume)

        result = {
            "status": "success",
            "source_file": os.path.basename(input_path),
            "bounding_box_mm": {
                "length_x": round(bounding_box.XLength, 3),
                "width_y": round(bounding_box.YLength, 3),
                "height_z": round(bounding_box.ZLength, 3)
            },
            "volume_mm3": round(volume, 3),
            "surface_area_mm2": round(float(final_shape.Area), 3),
            "solid_count": solid_count,
            "shell_count": len(final_shape.Shells),
            "face_count": len(final_shape.Faces),
            "edge_count": len(final_shape.Edges),
            "vertex_count": len(final_shape.Vertexes),
            "imported_object_count": len(shapes),
            "volume_reliability": (
                "high" if solid_count > 0 else "limited"
            )
        }

        if solid_count == 0:
            result["warning"] = (
                "The IGES geometry was not recognized as a formally "
                "closed solid. Bounding-box dimensions and surface area "
                "were extracted, but the volume should be treated as "
                "potentially unreliable."
            )

        return result

    finally:
        App.closeDocument(document.Name)


# ------------------------------------------------------------
# Main execution
# ------------------------------------------------------------

def main():
    try:
        if not os.path.isfile(INPUT_PATH):
            raise FileNotFoundError(
                f"IGES input file was not found: {INPUT_PATH}"
            )

        output_directory = os.path.dirname(OUTPUT_PATH)

        if output_directory:
            os.makedirs(output_directory, exist_ok=True)

        result = analyze_iges(INPUT_PATH)

    except Exception as error:
        result = {
            "status": "error",
            "source_file": os.path.basename(INPUT_PATH),
            "message": str(error)
        }

    with open(OUTPUT_PATH, "w", encoding="utf-8") as output_file:
        json.dump(result, output_file, indent=2)

    print()
    print("GEOMETRY ANALYSIS RESULT")
    print(json.dumps(result, indent=2))


if __name__ == "__main__":
    main()
main()
