import pandas as pd

def sel(item):
    if item=="无":
        return 1000
    st=item.strftime(format("%H:%M:%S"))
    return int(st.split(":")[0])*100+int(st.split(":")[1])

df=pd.read_excel(r"C:\Users\Administrator\Desktop\aaaaa.xlsx",usecols=["部门","start","start1"])
print(df.dtypes)
#df[(df["start"].apply(sel))<906]["start1"]="正常"  执行不了
#print(df["start"].apply(sel)<906)
temp=df["start"].apply(sel)<906
dep=df["部门"].isin(["研发部","电网业务部"])
df.loc[df[temp & dep].index,["start1"]]="正常"
df.drop(df[df["start1"]=="正常"].index,axis=0,inplace=True)
df=df.reset_index(drop=True)
print(df)
