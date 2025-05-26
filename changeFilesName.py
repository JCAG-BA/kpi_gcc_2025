import os

def rename_files_by_prefix(directory, original_file, prefix):
    try:
        # Validate if the directory exists
        if not os.path.exists(directory):
            print(f"Error: The directory '{directory}' does not exist.")
            return

        # List all files in the directory
        # print(files)
        # Chekear si existe el archivo original
        old_file = os.path.join(directory, original_file)
        # print(old_file)
        if os.path.isfile(old_file):
                    print("Entro!!")
                    # If the file exists, remove it before renaming
                    if os.path.exists(old_file):
                        print("Existe")
                        os.remove(old_file)
                    print(f"Eliminado: {old_file}")


        files = os.listdir(directory)
        for file in files:
            if file.startswith(prefix):
                old_file_path = os.path.join(directory, file)
                print(old_file_path)
                # Split the file into name and extension
                file_name, file_extension = os.path.splitext(file)

                new_name = os.path.join(directory, original_file)
                print(f"rename {new_name}")
                # Replace the prefix in the file name
                os.rename(old_file_path, new_name)

        print("File renaming process completed.")

    except Exception as e:
        print(f"An error occurred: {e}")




# Example usage
if __name__ == "__main__":

    # rename_files_by_prefix(r"C:\Users\jcarchila\Downloads".strip(), "tiempos_originación".strip())
    rename_files_by_prefix(r"C:\Users\jcarchila\Downloads".strip(), "tiempos_análisis.xlsx".strip(), "tiempos_análisis")
    rename_files_by_prefix(r"C:\Users\jcarchila\Downloads".strip(), "tiempos_originación.xlsx".strip(), "tiempos_originación")
    rename_files_by_prefix(r"C:\Users\jcarchila\Downloads".strip(), "creditos_mayores.xlsx".strip(), "creditos_mayores")
    rename_files_by_prefix(r"C:\Users\jcarchila\Downloads".strip(), "carga_laboral.xlsx".strip(), "carga_laboral")
    rename_files_by_prefix(r"C:\Users\jcarchila\Downloads".strip(), "tiempos_incidentes.xlsx".strip(), "tiempos_incidentes")

