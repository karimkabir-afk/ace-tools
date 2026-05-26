def get_float(prompt):
    while True:
        try:
            value = float(input(prompt).replace(",", ""))
            if value < 0:
                print("  Value must be non-negative. Try again.")
                continue
            return value
        except ValueError:
            print("  Invalid input. Please enter a numeric value.")


def main():
    print("\n=== Lease-Adjusted Leverage Ratio Calculator ===")
    print("Formula: (Funded Debt + 8× Annual Rent) ÷ TTM EBITDAR")
    print("Threshold: 5.75x\n")

    funded_debt = get_float("Funded Debt ($):          ")
    annual_rent = get_float("Trailing Twelve Month Rent ($):   ")
    ebitda      = get_float("Trailing Twelve Month EBITDA ($):  ")

    if ebitda == 0:
        print("\nError: EBITDA cannot be zero (division by zero).")
        return

    ebitdar = ebitda + annual_rent
    lease_adjusted_debt = funded_debt + 8 * annual_rent
    ratio = lease_adjusted_debt / ebitdar
    ebitda_needed = (lease_adjusted_debt / 5.75) - annual_rent

    print("\n--- Results ---")
    print(f"  Funded Debt:              ${funded_debt:>15,.2f}")
    print(f"  8× TTM Rent:              ${8 * annual_rent:>15,.2f}")
    print(f"  Lease-Adjusted Debt:      ${lease_adjusted_debt:>15,.2f}")
    print(f"  TTM EBITDA:               ${ebitda:>15,.2f}")
    print(f"  TTM EBITDAR (calculated): ${ebitdar:>15,.2f}")
    print(f"  Lease-Adjusted Leverage:  {ratio:>15.2f}x")
    print(f"  Threshold:                {'':>14} 5.75x")
    print(f"  TTM EBITDA needed for 5.75x: ${ebitda_needed:>12,.2f}")

    status = "PASS" if ratio <= 5.75 else "FAIL"
    print(f"\n  Result: {status}  ({'within' if ratio <= 5.75 else 'exceeds'} threshold by "
          f"{abs(ratio - 5.75):.2f}x)")
    print()


if __name__ == "__main__":
    main()
