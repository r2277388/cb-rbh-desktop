import pandas as pd

def upload_ams_raw():
    ams_list_raw = ['Period','ISBN', 'AMS Spend','AMS Sales','AMS Units']
    df_ams = pd.read_csv('ams.csv',
                         usecols=ams_list_raw,
                         dtype={'ISBN': 'object'})
    df_ams.replace('#N/A', pd.NA, inplace=True)
    df_ams.dropna(inplace=True)
    df_ams['ISBN'] = df_ams['ISBN'].astype(str)
    df_ams['ISBN'] = df_ams['ISBN'].str.strip()
    # Ensure ISBN is 13 characters by padding with zeros if necessary
    df_ams['ISBN'] = df_ams['ISBN'].str.zfill(13)
    df_ams['Period'] = df_ams['Period'].astype(str)
    df_ams = df_ams[df_ams['AMS Spend']!= 0]
    df_ams = df_ams[df_ams['ISBN'] != '-']
    
    return df_ams


def upload_ams():
    ams_list = ['Period','ISBN', 'AMS Spend']
    
    df = upload_ams_raw()
    df.drop(columns = ams_list)
    
    return df

def main():
    df_raw = upload_ams_raw()
    print('Load AMS Raw ',df_raw.head())
    print('AMS Raw Info',df_raw.info())
    print("-"*60)
    print(upload_ams().head())
    print("AMS DF ", upload_ams().info())
    print("-"*60)
    
if __name__ == '__main__':
    main()