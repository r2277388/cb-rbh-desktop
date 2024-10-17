import pandas as pd

def asin_isbn_key():
    """
    Helps us move from ASIN to ISBN's
    This is being updated manually
    """
    df = pd.read_csv("asin_isbn_key.csv")
    return df

if __name__ == '__main__':
    asin_key = asin_isbn_key()

    print(asin_key.head())