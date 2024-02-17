import numpy_financial as npf

def calculate_mortgage_details(principal, annual_interest_rate, years, down_payment_percentage, check_in_years):
    down_payment_amount = principal * (down_payment_percentage / 100)
    loan_amount = principal - down_payment_amount
    monthly_interest_rate = annual_interest_rate / 12 / 100
    n_payments = years * 12
    monthly_payment = npf.pmt(monthly_interest_rate, n_payments, -loan_amount)
    
    results = []
    for check_in_year in check_in_years:
        n_check_in_payments = check_in_year * 12
        principal_paid_until_check_in = sum(npf.ppmt(monthly_interest_rate, month, n_payments, -loan_amount) for month in range(1, n_check_in_payments + 1))
        interest_paid_until_check_in = sum(npf.ipmt(monthly_interest_rate, month, n_payments, -loan_amount) for month in range(1, n_check_in_payments + 1))
        total_paid_until_check_in = monthly_payment * n_check_in_payments
        principal_remaining = loan_amount - principal_paid_until_check_in

        results.append({
            'check_in_year': check_in_year,
            'down_payment_amount': down_payment_amount,
            'principal_paid_until_check_in': principal_paid_until_check_in,
            'interest_paid_until_check_in': interest_paid_until_check_in,
            'total_paid_until_check_in': total_paid_until_check_in,
            'principal_remaining': principal_remaining,
        })
    
    total_interest_paid = sum(npf.ipmt(monthly_interest_rate, month, n_payments, -loan_amount) for month in range(1, n_payments + 1))
    total_principal_paid = loan_amount
    total_paid_over_life = monthly_payment * n_payments + down_payment_amount

    return {
        'loan_term_years': years,
        'interest_rate': annual_interest_rate,
        'down_payment_percentage': down_payment_percentage,
        'monthly_payment': monthly_payment,
        'total_interest_paid': total_interest_paid,
        'total_principal_paid': total_principal_paid,
        'total_paid_over_life': total_paid_over_life,
        'check_in_details': results
    }

def display_results(loan_amount, loan_details, down_payment_options, check_in_years):
    print(f"\nMortgage Calculation Results for Loan Amount: ${loan_amount:,.2f}\n")
    
    headers = ['Loan Term (Years)', 'Interest Rate (%)', 'Down Payment (%)', 'Down Payment Amount ($)']
    for year in check_in_years:
        headers += [f'Principal Paid by Year {year} ($)', f'Interest Paid by Year {year} ($)', f'Total Payment by Year {year} ($)', f'Principal Remaining after Year {year} ($)']
    headers += ['Monthly Payment ($)', 'Total Interest Paid ($)', 'Total Principal Paid ($)', 'Total Paid Over Life ($)']
    
    print(' | '.join(headers))
    
    for detail in loan_details:
        for down_payment in down_payment_options:
            result = calculate_mortgage_details(loan_amount, detail['interest_rate'], detail['years'], down_payment, check_in_years)
            row = [f"{detail['years']}", f"{detail['interest_rate']}%", f"{down_payment}%", f"{result['check_in_details'][0]['down_payment_amount']:,.2f}"]
            for check_in_detail in result['check_in_details']:
                row += [f"{check_in_detail['principal_paid_until_check_in']:,.2f}", f"{check_in_detail['interest_paid_until_check_in']:,.2f}", f"{check_in_detail['total_paid_until_check_in']:,.2f}", f"{check_in_detail['principal_remaining']:,.2f}"]
            row += [f"{result['monthly_payment']:,.2f}", f"{result['total_interest_paid']:,.2f}", f"{result['total_principal_paid']:,.2f}", f"{result['total_paid_over_life']:,.2f}"]
            print(' | '.join(row))

import csv

def save_results_to_csv(loan_amount, loan_details, down_payment_options, check_in_years):
    with open('mortgage_results.csv', mode='w', newline='') as file:
        writer = csv.writer(file)
        
        # Write header row
        headers = ['Loan Term (Years)', 'Interest Rate (%)', 'Down Payment (%)', 'Down Payment Amount ($)']
        for year in check_in_years:
            headers += [f'Principal Paid by Year {year} ($)', f'Interest Paid by Year {year} ($)', f'Total Payment by Year {year} ($)', f'Principal Remaining after Year {year} ($)']
        headers += ['Monthly Payment ($)', 'Total Interest Paid ($)', 'Total Principal Paid ($)', 'Total Paid Over Life ($)']
        writer.writerow(headers)
        
        # Write data rows
        for detail in loan_details:
            for down_payment in down_payment_options:
                result = calculate_mortgage_details(loan_amount, detail['interest_rate'], detail['years'], down_payment, check_in_years)
                row = [f"{detail['years']}", f"{detail['interest_rate']}%", f"{down_payment}%", f"{result['check_in_details'][0]['down_payment_amount']:,.2f}"]
                for check_in_detail in result['check_in_details']:
                    row += [f"{check_in_detail['principal_paid_until_check_in']:,.2f}", f"{check_in_detail['interest_paid_until_check_in']:,.2f}", f"{check_in_detail['total_paid_until_check_in']:,.2f}", f"{check_in_detail['principal_remaining']:,.2f}"]
                row += [f"{result['monthly_payment']:,.2f}", f"{result['total_interest_paid']:,.2f}", f"{result['total_principal_paid']:,.2f}", f"{result['total_paid_over_life']:,.2f}"]
                writer.writerow(row)

def main():
    loan_amount = float(input("Enter loan amount: $"))
    n_loan_lengths = int(input("How many loan lengths do you want to compare? "))
    loan_details = [{'years': int(input(f"Enter loan length #{i+1} (in years): ")), 'interest_rate': float(input(f"Enter interest rate for this loan length (in %): "))} for i in range(n_loan_lengths)]
    
    down_payment_options = [float(input("Enter down payment percentage: ")) for _ in range(int(input("How many down payment options? ")))]

    check_in_points = int(input("How many check-in points do you want to evaluate? "))
    check_in_years = [int(input(f"Enter check-in year #{i+1}: ")) for i in range(check_in_points)]

    # Process and display results for each combination
    display_results(loan_amount, loan_details, down_payment_options, check_in_years)
    
    # Call save_results_to_csv function to save results to CSV
    save_results_to_csv(loan_amount, loan_details, down_payment_options, check_in_years)

if __name__ == "__main__":
    main()

