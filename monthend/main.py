try:
    from .bn_monthly_coop import main as run_bn_monthly_coop
except ImportError:
    from bn_monthly_coop import main as run_bn_monthly_coop


def main():
    while True:
        print("\nMonthend Reports")
        print()
        print("    1. Barnes & Noble Monthly Coop (Ailing)")
        print("    2. Back")
        print()

        choice = input("Choose an option: ").strip().lower()
        if choice == "1":
            try:
                run_bn_monthly_coop()
            except Exception as exc:
                print(f"Barnes & Noble Monthly Coop failed: {exc}")
        elif choice in {"2", "b", "back", "q", "quit", "exit"}:
            return
        else:
            print("Invalid choice.")


if __name__ == "__main__":
    main()
