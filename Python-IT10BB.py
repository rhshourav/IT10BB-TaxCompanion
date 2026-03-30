# it10bb_smart_estimator.py
# Smart rough estimator for IT-10BB expense sections (Bangladesh).
# Inputs: total annual expense (BDT), location (dhaka/other), family size, has_kids, own_home, home_support_staff, mode.
# Output: whole number percentage + amount per IT-10BB section (100% accurate totals)

def compute_allocation(total_expense,
                       location='other_area',
                       family_size=3,
                       has_kids=True,
                       own_home=False,
                       home_support_staff=False,
                       mode='balanced'):
    """
    mode: 'balanced' (default), 'conservative' (cut festival/home support), 'comfortable' (more festival/home support)
    """
    weights = {
        "Food, Clothing and Other Essentials": 30.0,
        "Accommodation Expense": 28.0,
        "Electricity": 2.5,
        "Gas, Water, Sewer and Garbage": 3.0,
        "Phone, Internet, TV channels & Subscriptions": 3.5,
        "Home-Support Staff and Other Expenses": 7.0,
        "Education Expenses (for kids)": 10.0,
        "Festival, Party, Events": 6.0
    }

    if not has_kids:
        weights["Education Expenses (for kids)"] = 0.0

    loc = location.strip().lower()
    if loc in ('dhaka', 'dhaka_city', 'city', 'big_city', 'metro'):
        weights["Accommodation Expense"] *= 1.20
        weights["Food, Clothing and Other Essentials"] *= 1.05
        weights["Phone, Internet, TV channels & Subscriptions"] *= 1.05
        weights["Home-Support Staff and Other Expenses"] *= 1.10
    else:
        weights["Accommodation Expense"] *= 0.90

    extra_people = max(0, int(family_size) - 2)
    if extra_people > 0:
        weights["Food, Clothing and Other Essentials"] *= (1 + 0.05 * extra_people)
        if has_kids:
            weights["Education Expenses (for kids)"] *= (1 + 0.04 * extra_people)

    if own_home:
        weights["Accommodation Expense"] *= 0.60
        weights["Home-Support Staff and Other Expenses"] *= 1.05

    if not home_support_staff:
        weights["Home-Support Staff and Other Expenses"] *= 0.4

    if mode == 'conservative':
        weights["Festival, Party, Events"] *= 0.6
        weights["Home-Support Staff and Other Expenses"] *= 0.7
        weights["Food, Clothing and Other Essentials"] *= 1.08
    elif mode == 'comfortable':
        weights["Festival, Party, Events"] *= 1.3
        weights["Home-Support Staff and Other Expenses"] *= 1.2
        weights["Food, Clothing and Other Essentials"] *= 1.05

    # Normalize to raw percentage
    total_weight = sum([w for w in weights.values() if w > 0])
    raw_percentages = {k: (w / total_weight) * 100.0 if w > 0 else 0.0 for k, w in weights.items()}

    # Convert to integer percentages with remainders
    int_percentages = {}
    remainders = []
    for k, p in raw_percentages.items():
        ip = int(p)
        int_percentages[k] = ip
        remainders.append((p - ip, k))

    # Distribute remaining percent to top categories by remainder
    current_total = sum(int_percentages.values())
    missing = 100 - current_total
    if missing > 0:
        remainders.sort(reverse=True)
        for i in range(missing):
            _, key = remainders[i]
            int_percentages[key] += 1

    # Now compute integer amounts (Tk) and fix drift
    int_amounts = {}
    raw_amounts = {}
    total_allocated = 0
    remainder_amounts = []
    for k in int_percentages:
        amt_raw = total_expense * int_percentages[k] / 100
        amt_int = int(amt_raw)
        int_amounts[k] = amt_int
        total_allocated += amt_int
        raw_amounts[k] = amt_raw
        remainder_amounts.append((amt_raw - amt_int, k))

    # Fix any leftover amount due to rounding
    remaining_bdt = int(round(total_expense - total_allocated))
    if remaining_bdt > 0:
        remainder_amounts.sort(reverse=True)
        for i in range(remaining_bdt):
            _, key = remainder_amounts[i]
            int_amounts[key] += 1

    return int_percentages, int_amounts


def print_breakdown(percentages, amounts, total_expense):
    print("\n======= IT-10BB Rough Expense Breakdown =======")
    for k in percentages:
        print(f"{k:<45} : {percentages[k]:3d}% | Tk {amounts[k]:,}")
    print("------------------------------------------------")
    print(f"{'TOTAL':<45} : {sum(percentages.values()):3d}% | Tk {sum(amounts.values()):,}")
    print("================================================\n")


if __name__ == "__main__":
    print("Smart IT-10BB Rough Expense Estimator\n")
    try:
        total = float(input("Enter total annual expenses (BDT): ").strip())
    except ValueError:
        print("Invalid total. Please enter a number.")
        raise SystemExit

    loc = input("Location (dhaka / other_area) [default other_area]: ").strip() or 'other_area'
    try:
        fam = int(input("Family size (number of people) [default 3]: ").strip() or "3")
    except ValueError:
        fam = 3
    kids = input("Do you have kids? (y/N) [default N]: ").strip().lower() == 'y'
    own = input("Own home (no rent)? (y/N) [default N]: ").strip().lower() == 'y'
    staff = input("Have home-support staff (driver/housemaid)? (y/N) [default N]: ").strip().lower() == 'y'
    mode = input("Mode (balanced/conservative/comfortable) [default balanced]: ").strip().lower() or 'balanced'

    pct, amt = compute_allocation(total, loc, fam, kids, own, staff, mode)
    print_breakdown(pct, amt, total)
