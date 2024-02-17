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

def excel_column_letter(index):
    """
    Convert a zero-indexed column number to a column letter (A-Z, AA-AZ, ...).
    """
    letter = ''
    while index > 0:
        index, remainder = divmod(index - 1, 26)
        letter = chr(65 + remainder) + letter
    return letter or 'A'

def save_results_to_excel(loan_amount, loan_details, down_payment_options, check_in_years):
    filename = get_next_filename()
    workbook = xlsxwriter.Workbook(filename)
    worksheet = workbook.add_worksheet()

    # Formats
    header_format = workbook.add_format({'bold': True, 'bg_color': 'yellow', 'border': 1})
    sub_header_format = workbook.add_format({'bold': True, 'bg_color': '#DDEBF7', 'border': 1})
    check_in_year_format = workbook.add_format({'bg_color': '#E2EFDA', 'border': 1})  # Green fill for check-in years
    money_format = workbook.add_format({'num_format': '$#,##0.00', 'border': 1})    

    # Top Header - Calculate the range based on number of columns
    num_fixed_headers = 4
    num_check_in_columns = len(check_in_years) * 4
    num_total_columns = num_fixed_headers + num_check_in_columns + len(['Total Interest Paid', 'Total Principal Paid', 'Total Paid Over Life'])
    top_header_range = f'A1:{excel_column_letter(num_total_columns)}1'
    worksheet.merge_range(top_header_range, f"Loan Amount: ${loan_amount:,.2f}", header_format)

    # Fixed Headers
    fixed_headers = ['Term (Years)', 'Interest Rate', 'Down Payment %', 'Monthly Payment']
    worksheet.write_row('A2', fixed_headers, header_format)

    # Check-in Year Headers
    for i, check_in_year in enumerate(check_in_years):
        col_offset = num_fixed_headers + i * 4
        start_col_letter = excel_column_letter(col_offset)
        end_col_letter = excel_column_letter(col_offset + 3)
        worksheet.merge_range(f'{start_col_letter}2:{end_col_letter}2', f'{check_in_year} Year Mark', sub_header_format)

        check_in_headers = [
            f'Principal Paid {check_in_year} yr',
            f'Interest {check_in_year} yr',
            f'Total Payment {check_in_year} yr',
            f'Loan Left {check_in_year} yr'
        ]
        worksheet.write_row(f'{start_col_letter}3', check_in_headers, sub_header_format)

    # Totals Headers - Start after the check-in year columns
    totals_start_col = excel_column_letter(num_fixed_headers + num_check_in_columns)
    totals_headers = ['Total Interest Paid', 'Total Principal Paid', 'Total Paid Over Life']
    worksheet.write_row(f'{totals_start_col}2', totals_headers, header_format)

    # Write data
    row = 3  # Start from row 3 to account for the headers
    for detail in loan_details:
        for down_payment in down_payment_options:
            result = calculate_mortgage_details(loan_amount, detail['interest_rate'], detail['years'], down_payment, check_in_years)
            data = [detail['years'], detail['interest_rate'], down_payment, result['monthly_payment']]
            worksheet.write_row(row, 0, data, money_format)

            # Write check-in years data with specific formatting
            for i, check_in_detail in enumerate(result['check_in_details']):
                col_offset = num_fixed_headers + i * 4
                check_in_data = [
                    check_in_detail['principal_paid_until_check_in'],
                    check_in_detail['interest_paid_until_check_in'],
                    check_in_detail['total_paid_until_check_in'],
                    check_in_detail['principal_remaining']
                ]
                for col, value in enumerate(check_in_data, start=col_offset):
                    worksheet.write_number(row, col, value, check_in_year_format if i < len(check_in_years) else money_format)

            totals_data = [result['total_interest_paid'], detail['years'] * loan_amount, result['total_paid_over_life']]
            for col, value in enumerate(totals_data, start=num_fixed_headers + num_check_in_columns):
                worksheet.write_number(row, col, value, money_format)
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
