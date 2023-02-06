import pandas as pd

class CaptionData:

  def __init__(self, title, name, university, material, size):
    self.title = str(title) if not pd.isna(title) else ""
    self.name = str(name) if not pd.isna(name) else ""
    self.university = str(university) if not pd.isna(university) else ""
    self.material = str(material) if not pd.isna(material) else ""
    self.size = str(size) if not pd.isna(size) else ""
