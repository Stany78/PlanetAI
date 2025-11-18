import os
import zipfile

SOURCE_DIR = "Omi"
MAX_ZIP_SIZE_MB = 80   # ogni zip massimo 80MB
MAX_ZIP_SIZE = MAX_ZIP_SIZE_MB * 1024 * 1024


def main():
    if not os.path.isdir(SOURCE_DIR):
        print(f"ERRORE: la cartella '{SOURCE_DIR}' non esiste!")
        return

    files = sorted(os.listdir(SOURCE_DIR))

    print(f"📁 Trovati {len(files)} file in {SOURCE_DIR}")

    zip_index = 1
    current_zip_size = 0
    zip_filename = f"Omi_{zip_index}.zip"
    zipf = zipfile.ZipFile(zip_filename, 'w', zipfile.ZIP_DEFLATED)

    print(f"📦 Creo {zip_filename}")

    for fname in files:
        path = os.path.join(SOURCE_DIR, fname)
        fsize = os.path.getsize(path)

        # Se aggiungendo questo file supero 80MB -> chiudo ZIP e ne apro uno nuovo
        if current_zip_size + fsize > MAX_ZIP_SIZE:
            zipf.close()
            print(f"✔ Completato {zip_filename} ({current_zip_size/1024/1024:.1f} MB)")

            zip_index += 1
            zip_filename = f"Omi_{zip_index}.zip"
            zipf = zipfile.ZipFile(zip_filename, 'w', zipfile.ZIP_DEFLATED)
            print(f"📦 Creo {zip_filename}")
            current_zip_size = 0

        zipf.write(path, arcname=fname)
        current_zip_size += fsize

    # Chiudo l'ultimo zip
    zipf.close()
    print(f"✔ Completato {zip_filename} ({current_zip_size/1024/1024:.1f} MB)")

    print("\n🎉 COMPLETATO!")
    print("Puoi caricare i file Omi_*.zip su GitHub. Tutti <80MB.")
    print("Nel progetto NON devi includere la cartella Omi originale.")
    

if __name__ == "__main__":
    main()
