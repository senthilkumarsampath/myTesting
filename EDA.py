import numpy as np
import pandas as pd
from pandas_profiling import ProfileReport
df = pd.read_csv("C:/Users/senthilps/Downloads/h1.csv")
profile = ProfileReport(df, title="Report",html={'style':{'full_width':True}},sort="ascending")
profile.to_file("C:/Users/senthilps/Downloads/h1.html")