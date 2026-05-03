#%%
import requests
import pandas as pd
import os
from io import StringIO
from dotenv import load_dotenv
load_dotenv()

url = os.getenv("LINK_PLANNER")

resp = requests.get(url)
df_planner = pd.read_csv(StringIO(resp.text))

