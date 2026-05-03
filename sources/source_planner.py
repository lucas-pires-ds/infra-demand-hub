import pandas as pd
from pathlib import Path

BASE = Path().absolute()
input_dir = BASE / "mock_data/output"

print("definindo df_planner...")
df_planner = pd.read_csv(input_dir / "planner_tarefas.csv")