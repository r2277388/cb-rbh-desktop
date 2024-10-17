import pandas as pd

def upload_uc():
    df = pd.read_csv('unit_cost.csv',dtype = {'ISBN':object,'UC':float})
    df['ISBN'] = df['ISBN'].str.zfill(13)

    return df

def main():
    df = upload_uc()
    print(df.head())
    print(df.info())
    
if __name__ == '__main__':
    main()