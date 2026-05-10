#%%
import pandas as pd
import requests
import os
from dotenv import load_dotenv
load_dotenv()

url = os.getenv("LINK_QUALYTEAM_API")
resp = requests.get(url).json()

data = resp["demandas"]

print("Criando df_qualyteam...")
df_qualyteam = pd.DataFrame(data)

