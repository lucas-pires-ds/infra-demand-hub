#%%
import pandas as pd
import requests
import os
from pathlib import Path
from io import BytesIO
from dotenv import load_dotenv
load_dotenv()

url = os.getenv("LINK_CALL")
resp = requests.get(url)

df_call = pd.read_excel(BytesIO(resp.content))
