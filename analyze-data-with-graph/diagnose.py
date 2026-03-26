import pandas as pd

df = pd.read_excel(r"K:\GoogleDrive\00. Quick Share\AIA생명\Application_Insight_021013_021016_original.xlsx")
# B컬럼 원본값 5개 출력
print(df.iloc[:5, 1].tolist())
print(df.iloc[:5, 1].dtype)