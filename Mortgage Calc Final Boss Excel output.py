import numpy_financial as npf
import xlsxwriter
import os

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
    total_paid_over_life = monthly_payment * n_payments + down_payment_amount

    return {
        'monthly_payment': monthly_payment,
        'total_interest_paid': total_interest_paid,
        'total_paid_over_life': total_paid_over_life,
        'check_in_details': results
    }

def get_next_filename():
    base_filename = 'Mortgage_results'
    extension = '.xlsx'
    counter = 1
    while os.path.isfile(f"{base_filename}_{counter}{extension}"):
        counter += 1
    return f"{base_filename}_{counter}{extension}"

def save_results_to_excel(loan_amount, loan_details, down_payment_options, check_in_years):
    filename = get_next_filename()
    workbook = xlsxwriter.Workbook(filename)
    worksheet = workbook.add_worksheet()

    # Header Format
    header_format = workbook.add_format({'bold': True, 'bg_color': 'yellow', 'border': 1})
    money_format = workbook.add_format({'num_format': '$#,##0.00', 'border': 1})
    percent_format = workbook.add_format({'num_format': '0.00%', 'border': 1})

    # Headers
    headers = ['Loan Term (Years)', 'Interest Rate', 'Down Payment %', 'Monthly Payment', 'Total Interest Paid', 'Total Paid Over Life']
    for year in check_in_years:
        headers.extend([f'Principal Paid by Year {year}', f'Interest Paid by Year {year}', f'Total Payment by Year {year}', f'Principal Remaining after Year {year}'])
    worksheet.write_row(0, 0, headers, header_format)

    row = 1
    for detail in loan_details:
        for down_payment in down_payment_options:
            result = calculate_mortgage_details(loan_amount, detail['interest_rate'], detail['years'], down_payment, check_in_years)
            data = [detail['years'], detail['interest_rate'], down_payment, result['monthly_payment'], result['total_interest_paid'], result['total_paid_over_life']]
            for check_in_detail in result['check_in_details']:
                data.extend([check_in_detail['principal_paid_until_check_in'], check_in_detail['interest_paid_until_check_in'], check_in_detail['total_paid_until_check_in'], check_in_detail['principal_remaining']])
            worksheet.write_row(row, 0, data, money_format)
            row += 1

    workbook.close()
    print(f"Results saved to {filename}")

def main():
    loan_amount = float(input("Enter loan amount: $"))
    n_loan_lengths = int(input("How many loan lengths do you want to compare? "))
    loan_details = [{'years': int(input(f"Enter loan length #{i+1} (in years): ")), 'interest_rate': float(input(f"Enter interest rate for this loan length (in %): "))} for i in range(n_loan_lengths)]
    
    down_payment_options = [float(input(f"Enter down payment percentage option #{i+1}: ")) for i in range(int(input("How many down payment options? ")))]
    check_in_points = int(input("How many check-in points do you want to evaluate? "))
    check_in_years = [int(input(f"Enter check-in year #{i+1}: ")) for i in range(check_in_points)]

    save_results_to_excel(loan_amount, loan_details, down_payment_options, check_in_years)

if __name__ == '__main__':
    main()
