from pathlib import Path

def create_folder(folder_path):
    """
    Verilen yolda klasör oluşturur
    """
    klasor_yolu = Path(folder_path)
    return klasor_yolu.mkdir(parents=True, exist_ok=True)