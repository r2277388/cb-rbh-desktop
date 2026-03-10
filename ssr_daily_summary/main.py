# -*- coding: utf-8 -*-
"""
Created on Sat Aug  3 18:31:57 2024

@author: RBH
"""

# %% Imports
import time

from ssr_preparation import main as run_ssr_preparation
from ssr_summary import main as run_ssr_summary
from ssr_visualizations import create_viz


def format_elapsed_time(seconds):
    total_seconds = int(round(seconds))
    minutes, secs = divmod(total_seconds, 60)
    hours, minutes = divmod(minutes, 60)

    if hours > 0:
        return f"{hours} hour(s), {minutes} minute(s), and {secs} second(s)"
    if minutes > 0:
        return f"{minutes} minute(s) and {secs} second(s)"
    return f"{secs} second(s)"


def run_with_timer(label, func):
    start_time = time.perf_counter()
    try:
        func()
    finally:
        elapsed = time.perf_counter() - start_time
        print(f"{label} took {format_elapsed_time(elapsed)} to run.")


def main():
    while True:
        print("\nChoose an option:")
        print("1 - Run SSR Daily Reporting")
        print("2 - Run SSR Aggregate Totals")
        print("3 - Run SSR Visualization")
        print("4 - Back to launcher")

        choice = input("Enter your choice (1, 2, 3, or 4): ").strip().lower()

        if choice == '1':
            print("Running SSR Daily Reporting...")
            run_with_timer("SSR Daily Reporting", run_ssr_preparation)
            continue
        elif choice == '2':
            print("Running SSR Aggregate Totals...")
            run_with_timer("SSR Aggregate Totals", run_ssr_summary)
            continue
        elif choice == '3':
            print("Running SSR Viz Process...")
            run_with_timer("SSR Visualization", create_viz)
            continue
        elif choice in ['4', 'back', 'b', 'exit', 'quit', 'q']:
            print("Returning to launcher...")
            return
        else:
            print("Invalid choice. Please enter 1, 2, 3, or 4.")

if __name__ == "__main__":
    main()  # Call the main function
