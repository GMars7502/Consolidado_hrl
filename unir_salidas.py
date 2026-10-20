import os
import glob
import time

def main():
    base_dir = os.path.dirname(os.path.abspath(__file__))
    salida_dir = os.path.join(base_dir, "SALIDA")
    f_out = os.path.join(base_dir, "consolidado_final.csv")
    
    archivos = glob.glob(os.path.join(salida_dir, "CONSOLIDADO_*.csv"))
    if not archivos:
        print("No se encontraron archivos procesados en la carpeta SALIDA.")
        return
        
    print(f"Iniciando unión de {len(archivos)} archivos...")
    start_time = time.time()
    
    with open(f_out, 'w', encoding='utf-8') as out:
        header_written = False
        
        for ruta_archivo in archivos:
            with open(ruta_archivo, 'r', encoding='utf-8') as infile:
                header_skipped = False
                for line in infile:
                    if not header_written:
                        out.write(line)
                        header_written = True
                        header_skipped = True
                        continue
                    
                    if not header_skipped:
                        header_skipped = True
                        continue
                        
                    out.write(line)

    print(f"¡Archivos unidos exitosamente en {time.time() - start_time:.2f} segundos!")
    print(f"Ruta final: {f_out}")

if __name__ == "__main__":
    main()
