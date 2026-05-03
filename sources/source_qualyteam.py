import pandas as pd
from pathlib import Path

BASE = Path().absolute()
input_dir = BASE / "mock_data/output"

print("definindo df_qualyteam...")
df_qualyteam = pd.read_json(input_dir / "qualyteam_demandas.json")