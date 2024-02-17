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

import xlsxwriter
import os
import random

def get_random_color():
    return "#"+''.join([random.choice('0123456789ABCDEF') for j in range(6)])

def save_results_to_excel(loan_amount, loan_details, down_payment_options, check_in_years):
    filename = get_next_filename()
    workbook = xlsxwriter.Workbook(filename)
    worksheet = workbook.add_worksheet()

    # Formats
    header_format = workbook.add_format({
        'bold': True,
        'align': 'center',
        'valign': 'vcenter',
        'fg_color': '#FFEB9C',
        'border': 1
    })
    money_format = workbook.add_format({'num_format': '$#,##0.00', 'border': 1})
    percent_format = workbook.add_format({'num_format': '0%', 'border': 1})
    bold = workbook.add_format({'bold': True, 'border': 1})
    
    # Write the fixed headers
    fixed_headers = ['Term', 'Interest', 'Down Payment', 'Monthly Mortgage']
    worksheet.write_row('A1', fixed_headers, header_format)

    # Add check-in times headers with random colors
    check_in_headers = []
    for check_in_year in check_in_years:
        color = get_random_color()
        check_in_format = workbook.add_format({'bg_color': color, 'border': 1})
        check_in_headers.extend([
            f'Principal Paid {check_in_year}', f'Interest {check_in_year}',
            f'All Pays {check_in_year}', f'Loan Left {check_in_year}'
        ])
        for i in range(4):  # Four headers per check-in year
            worksheet.write(0, len(fixed_headers) + i + (len(check_in_years) - 1) * 4, check_in_headers[i + (check_in_year - 1) * 4], check_in_format)

    # Write data to Excel
    row = 1
    for detail in loan_details:
        data = [
            detail['loan_term_years'],
            detail['interest_rate'],
            # Extract other fields from the 'detail' dictionary and format as needed
        ]
        # Write the row of data to Excel, applying formats
        worksheet.write_row(f'A{row + 1}', data, money_format)
        row += 1

    workbook.close()
    print(f"Saved Excel file to: {filename}")

# ... (rest of your code)


        # Example of data list, replace with actual data from 'detail'
        data = [
            detail['loan_term_years'],
            detail['interest_rate'],
            # ... add other fields from the 'detail' dictionary
        ]
        # Write data with formats
        worksheet.write(row, 0, data[0], bold)  # Term, assuming it's an integer
        worksheet.write(row, 1, data[1], percent_format)  # Interest
        # ... write other data cells with appropriate formats
        row += 1

    workbook.close()
    print(f"Saved Excel file to: {filename}")

# ... (rest of your code)


