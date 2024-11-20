# -*- coding: utf-8 -*-
"""
Created on Sat Aug  3 18:31:57 2024

@author: RBH
"""

# %% Imports
from ssr_preparation import main as run_ssr_preparation
from ssr_summary import main as run_ssr_summary
from ssr_visualizations import create_viz

def main():
    print("Choose an option:")
    print("1 - Run SSR Preparation")
    print("2 - Run SSR Summary")
    print("3 - Run SSR Visualization")

    choice = input("Enter your choice (1, 2 or 3): ")
    
    if choice == '1':
        print("Running SSR Preparation...")
        run_ssr_preparation()
    elif choice == '2':
        print("Running SSR Summary...")
        run_ssr_summary()
    elif choice == '3':
        print("Running SSR Viz Process...")
        create_viz()
    else:
        print("Invalid choice. Please enter 1, ,2 or 3.")

if __name__ == "__main__":
    main()  # Call the main function