#%%
import pandas as pd
import os
import requests
from dotenv import load_dotenv
load_dotenv()


url = os.getenv("LINK_QUALYTEAM")
resp = requests.get(url).json()

df_qualyteam = pd.DataFrame(resp)
