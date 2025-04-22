from load_cdu_comp_table import upload_cdu_table

def create_cdu_dict():
    df = upload_cdu_table()
    # Group by 'cdu' and only apply the function on the 'component' and 'qty' columns
    cdu_dict = df.groupby('cdu')[['component', 'qty']].apply(
        lambda x: dict(zip(x['component'], x['qty']))
    ).to_dict()
    
    return cdu_dict

def main():
    cdu_dict = create_cdu_dict()
    print(cdu_dict['0810073340169'])
    
if __name__ == "__main__":
    main()