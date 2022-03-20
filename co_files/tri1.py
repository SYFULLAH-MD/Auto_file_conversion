import pandas as pd
read_file = pd.read_excel ("Book1.xlsx")
read_file.to_csv ("Test.txt", 
                  index = None,
                  header=True)
df = pd.DataFrame(pd.read_csv("Test.csv"))
df
