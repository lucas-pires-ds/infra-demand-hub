import pandas as pd
from pathlib import Path

BASE = Path().absolute()
input_dir = BASE / "mock_data/output"

print("definindo df_call...")
df_call = pd.read_excel(input_dir / "call_chamados.xlsx")