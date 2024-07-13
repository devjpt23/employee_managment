import pandas as pd

emp_dict = {}
readfile = pd.read_csv("C:\\Users\\25470\\Desktop\\new project meeting one\\thisReal.csv")

for i in readfile._values:
    if(i[1]) in emp_dict:
        emp_dict[i[1]].append(i)
    else:
        emp_dict[i[1]] = []
        emp_dict[i[1]].append(i)


df =  pd.DataFrame(dict([(key, pd.Series(value)) for key, value in emp_dict.items()]))
with pd.ExcelWriter('output.xlsx') as writer:
    df.to_excel(writer, index=False)