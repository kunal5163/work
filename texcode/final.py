import subprocess
import os

def latex_to_docx(tex_file, output_docx, image_dir="images"):
    # Ensure the .tex file exists
    if not os.path.isfile(tex_file):
        raise FileNotFoundError(f"{tex_file} not found")

    # Build pandoc command
    command = [
        "pandoc",
        "-s", tex_file,
        "--resource-path", image_dir,  # for image resolution
        "-o", output_docx
    ]

    try:
        # Run the conversion
        subprocess.run(command, check=True)
        print(f"✅ Word file generated: {output_docx}")
    except subprocess.CalledProcessError as e:
        print("❌ Pandoc failed:", e)

if __name__ == "__main__":
    latex_to_docx("small_highlighted_output.tex", "converted_output.docx", image_dir="images")
