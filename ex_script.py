import pandas as pd
import matplotlib.pyplot as plt
import os 
Path=("G:/excel_files/supermarket.xlsx")
if os.path.exists(Path):
    df=pd.read_excel(Path,usecols='B:I',skiprows=7)

    # clean date format
    df['Order Date']=pd.to_datetime(df['Order Date'],errors='coerce').dt.strftime('%Y-%m-%d')
    df['Ship Date']=pd.to_datetime(df['Ship Date'],errors='coerce').dt.strftime('%Y-%m-%d')

    print(df.head(4))
    print(df.columns)
    total_usd=df[df['Total (USD)']>150]
    total_usd.to_excel('Total_usd.xlsx',index=False)
    print(total_usd)
    df.fillna(0,inplace=True)
    df.drop_duplicates(inplace=True)

# Generate report for each day
    daily_report=df.groupby('Order Date').agg({'Total_USD':'sum','Order No':'count'}).reset_index()
    print(daily_report)
    daily_report.rename(columns={'Total_USD':'Daily Total USD','Order ID':'Number of Orders'},inplace=True)
    daily_report.to_excel('Daily_report.xlsx',index=False)

    # simple visualization
    daily_report.plot(x='Order Date',y='Total_USD', kind="bar",figsize=(12,6),title="Daily Total USD Report",color='skyblue',edgecolor='black')

    plt.xlabel("Order Date",font_size=12)
    plt.ylabel("Total USD",font_size=12)
    plt.xticks(rotation=45)
    plt.tight_layout()
    plt.show()

                                                  



    # df['Remark']=df['Tax (USD)'].apply(lambda 
    # x:'High ax' if x>20 else 'low Tax')
    df.insert(7,'Tax-Remarks',df['Tax_USD'].apply(lambda x:'High Tax' if x>20 else 'Low Tax'))
    df.rename(columns={'Total (USD)':'Total_USD','Tax (USD)':'Tax_USD'},inplace=True)
    df.rename(columns={'Tax_Remark':'Tax-Remarks'},inplace=True)

    #standardize column names
    df['Customer Name']=df['Customer Name'].str.strip().str.title()
    

    # Generate file to excel
    df.to_excel('Cleaned_supermarket_new.xlsx',index=False)
    

    df.drop(columns=['Tax-Remarks'],inplace=True)

    # total_usd_summary=df['Total_USD'].sum()
    # print(f"Total USD Summary: {total_usd_summary}")

    df.insert(8,'Tax_perProduct',df['Tax_USD']/df['Order Quantity'])
    df.insert(8,'Tax_percentage',df['Tax_USD']/df['Total_USD']*100)
   
     
    df.rename(columns={'Order Quantity':'Order_Quantinty'},inplace=True)
    order_than_three=df[df['Order_Quantinty']>3]
    print(order_than_three)


    # to save the new report
    with pd.ExcelWriter('sales_more_than_3_units.xlsx') as writer:
        order_than_three.to_excel(writer,sheet_name='Orders>3Units',index=False)

    



else:
    print("The specified file does not exist.")


