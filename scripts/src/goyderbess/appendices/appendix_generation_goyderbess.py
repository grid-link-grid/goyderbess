import os
import glob
import shutil
from typing import List, Optional, Callable

DEFAULT_CLASS_ASSETS_LIST = [
        os.path.join(os.path.dirname(__file__), "grid-link-appendix-template.cls"),
        os.path.join(os.path.dirname(__file__), "report-assets")
        ]

def clean_tex_string(str_input: str):
    return str_input.replace("\\", "/").replace(" ", r"\space")

def generate_appendix_goyderbess(
        project_name : str,
        client : str,
        title : str,
        doc_number : str,
        issued_date : str,
        revision : str,
        revision_history_csv_path : str,
        plots_directory : str,
        output_path : str,
        class_assets: List[str] = DEFAULT_CLASS_ASSETS_LIST,
        bess_charging: Optional[bool] = False,
        caption_preprocessor_fn: Optional[Callable] = None,
         
        ):
    temp_dir = os.path.join(os.getcwd(), "_temp_latex")
    if os.path.exists(temp_dir):
        shutil.rmtree(temp_dir)
    os.makedirs(temp_dir)

    # Copy class assets into temp directory
    for asset_path in class_assets:
        destination_path = os.path.join(temp_dir, os.path.basename(asset_path))
        if os.path.exists(destination_path):
            continue

        if os.path.isfile(asset_path):
            shutil.copy2(asset_path, temp_dir)
        elif os.path.isdir(asset_path):
            shutil.copytree(asset_path, os.path.join(temp_dir, os.path.basename(asset_path)))

    # Copy plots into temp directory
    temp_plots_dir = os.path.join(temp_dir, "plots")
    if os.path.exists(temp_plots_dir):
        shutil.rmtree(temp_plots_dir)
    if os.path.exists(plots_directory):
        shutil.copytree(plots_directory, temp_plots_dir)
        
    if bess_charging is not None:
        all_png_files = sorted(glob.glob(os.path.join(temp_plots_dir, '*.png')))
        png_files = sorted([f for f in all_png_files if ("CRG" in os.path.basename(f)) == bess_charging])
    else:
        png_files = sorted(glob.glob(os.path.join(temp_plots_dir, '*.png')))
 


    lines = []
    lines.append("\\documentclass{grid-link-appendix-template}")
    lines.append("\\project{" + project_name + "}")
    lines.append("\\client{" + client + "}")
    lines.append("\\title{" + title + "}")
    lines.append("\\docnumber{" + doc_number + "}")
    lines.append("\\issueddate{" + issued_date + "}")
    lines.append("\\revision{" + revision + "}")
    lines.append("\\revisionhistorycsvpath{" + clean_tex_string(revision_history_csv_path) + "}")

    lines.append("\\begin{document}")
    lines.append("\\frontmatter")
    lines.append("\\maketitle")
    lines.append("\\listoffigures")
    lines.append("\\mainmatter")
    lines.append("\\uselandscape")

    for png_path in png_files:
            png_relpath = os.path.relpath(png_path, temp_dir)
            png_basename = os.path.basename(png_relpath).split(".")[0]
            caption = png_basename if caption_preprocessor_fn is None else caption_preprocessor_fn(png_basename)
            lines.append("\\begin{figure}[htbp]")
            lines.append("\\centering")
            lines.append("\\begin{textblock*}{\\paperwidth}(0cm,3.5cm)")
            lines.append("\\includegraphics[height=15cm,keepaspectratio]{" + clean_tex_string(png_relpath) + "}")
            lines.append("\\caption{" + caption + "}")
            lines.append("\\label{fig:" + png_basename + "}")
            lines.append("\\end{textblock*}")
            lines.append("\\end{figure}")
            lines.append("\\clearpage")




    lines.append(r"""\end{document}""")

    print(output_path)
    tex_path = os.path.join(temp_dir, f"{os.path.basename(output_path).split('.')[0]}.tex")
    print(tex_path)
    if os.path.exists(tex_path) and os.path.isfile(tex_path):
        os.remove(tex_path)
    with open(tex_path, 'w') as f:
        f.writelines("\n".join(lines))

    # Change into the directory before running latex
    cwd = os.getcwd()
    os.chdir(temp_dir)
    os.system(f'xelatex.exe -synctex=1 -interaction=nonstopmode "{tex_path}"')
    os.system(f'xelatex.exe -synctex=1 -interaction=nonstopmode "{tex_path}"')
    os.chdir(cwd)

    if os.path.exists(output_path) and os.path.isfile(output_path):
        shutil.rmtree(output_path)
    if not os.path.exists(os.path.dirname(output_path)):
        os.makedirs(os.path.dirname(output_path))

    temp_pdf_path = tex_path.split(".tex")[0] + ".pdf"
    shutil.copy2(src=temp_pdf_path, dst=output_path)

