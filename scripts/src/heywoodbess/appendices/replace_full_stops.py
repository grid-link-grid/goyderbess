import os
"RUN THIS SCRIPT THREE TIMES TO FIND ALL THE FULL STOPS"




def replace_full_stops(file_dir):
    FILES_DIR = file_dir
    for i in range(3):
        for root, _, files in os.walk(FILES_DIR):
            for file in files:
                filename, ext = os.path.splitext(file)

                new_filename = f'{filename.replace(".", "p")}{ext}'

                if new_filename != filename:
                    old_path = os.path.join(root, file)
                    new_path = os.path.join(root, new_filename)
                    # print(old_path)
                    # print(new_path)
                    os.rename(old_path, new_path)

