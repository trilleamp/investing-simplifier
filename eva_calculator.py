"""
EVA (Economic Value Added) Calculator
Based on: "Forget ROE. This Is the Metric That Actually Matters" - Jimmy's Journal

This script calculates:
  - NOPAT  (Net Operating Profit After Tax)
  - ROIC   (Return on Invested Capital)
  - WACC   (Weighted Average Cost of Capital)
  - EVA    (Economic Value Added)

Usage: Run this file and enter the numbers when prompted.
"""


def calculate_nopat(ebit, tax_rate):
    """NOPAT = EBIT x (1 - Tax Rate)"""
    return ebit * (1 - tax_rate)


def calculate_capital_invested(total_equity, interest_bearing_debt):
    """Capital Invested = Total Equity + Interest-Bearing Debt"""
    return total_equity + interest_bearing_debt


def calculate_wacc(equity, debt, cost_of_equity, cost_of_debt, tax_rate):
    """
    WACC = (E/V x cost of equity) + (D/V x cost of debt x (1 - tax rate))
    where V = E + D
    """
    total = equity + debt
    if total == 0:
        return 0
    equity_weight = equity / total
    debt_weight = debt / total
    return (equity_weight * cost_of_equity) + (debt_weight * cost_of_debt * (1 - tax_rate))


def calculate_roic(nopat, capital_invested):
    """ROIC = NOPAT / Capital Invested"""
    if capital_invested == 0:
        return 0
    return nopat / capital_invested


def calculate_eva(nopat, capital_invested, wacc):
    """EVA = NOPAT - (Capital Invested x WACC)"""
    return nopat - (capital_invested * wacc)


def get_number(prompt):
    """Helper to get a number from the user."""
    while True:
        try:
            return float(input(prompt))
        except ValueError:
            print("  Please enter a valid number.")


def main():
    print("=" * 60)
    print("  EVA (Economic Value Added) Calculator")
    print("=" * 60)
    print()
    print("Enter the following values (all dollar amounts in millions):")
    print()

    # Inputs
    ebit = get_number("  EBIT (Operating Income):          $")
    tax_rate = get_number("  Tax Rate (e.g. 0.21 for 21%):     ")
    total_equity = get_number("  Total Shareholders' Equity:       $")
    total_debt = get_number("  Interest-Bearing Debt:             $")
    cost_of_equity = get_number("  Cost of Equity (e.g. 0.10 for 10%): ")
    cost_of_debt = get_number("  Cost of Debt   (e.g. 0.05 for 5%):  ")
    net_income = get_number("  Net Income (for ROE comparison):   $")

    # Calculations
    nopat = calculate_nopat(ebit, tax_rate)
    capital_invested = calculate_capital_invested(total_equity, total_debt)
    wacc = calculate_wacc(total_equity, total_debt, cost_of_equity, cost_of_debt, tax_rate)
    roic = calculate_roic(nopat, capital_invested)
    eva = calculate_eva(nopat, capital_invested, wacc)
    roe = net_income / total_equity if total_equity != 0 else 0

    # Results
    print()
    print("=" * 60)
    print("  RESULTS")
    print("=" * 60)
    print()
    print(f"  NOPAT (operating profit after tax):  ${nopat:,.2f}M")
    print(f"  Capital Invested:                    ${capital_invested:,.2f}M")
    print(f"  WACC:                                {wacc:.2%}")
    print(f"  ROIC:                                {roic:.2%}")
    print(f"  ROE (for comparison):                {roe:.2%}")
    print()
    print(f"  >>> EVA = ${eva:,.2f}M <<<")
    print()

    # Interpretation
    if eva > 0:
        print("  VERDICT: This company is CREATING VALUE.")
        print(f"  It earns {roic:.2%} on its capital, which is above")
        print(f"  the {wacc:.2%} that investors and lenders expect.")
    elif eva < 0:
        print("  VERDICT: This company is DESTROYING VALUE.")
        print(f"  It earns only {roic:.2%} on its capital, which is below")
        print(f"  the {wacc:.2%} that investors and lenders expect.")
        if roe > 0:
            print()
            print(f"  WARNING: ROE of {roe:.2%} looks positive, but EVA reveals")
            print("  the company is not earning enough to cover its cost of capital.")
    else:
        print("  VERDICT: Break-even. The company earns exactly its cost of capital.")

    print()
    print("  ROIC vs WACC spread: {:.2%}".format(roic - wacc))
    print()


if __name__ == "__main__":
    main()
