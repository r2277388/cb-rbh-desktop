# -*- coding: utf-8 -*-
"""
Created on Sat Aug  3 18:31:57 2024

@author: RBH
"""

# %% Imports
from ssr_preparation import main as run_ssr_preparation
from ssr_summary import main as run_ssr_summary

def main():
    print("Choose an option:")
    print("1 - Run SSR Preparation")
    print("2 - Run SSR Summary")

    choice = input("Enter your choice (1 or 2): ")
    
    if choice == '1':
        print("Running SSR Preparation...")
        run_ssr_preparation()
    elif choice == '2':
        print("Running SSR Summary...")
        run_ssr_summary()
    else:
        print("Invalid choice. Please enter 1 or 2.")

if __name__ == "__main__":
    main()  # Call the main function