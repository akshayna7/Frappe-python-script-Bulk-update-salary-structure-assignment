import frappe
import frappe.utils
import pandas as pd
import datetime
from frappe import _
from frappe.model.document import Document




def insert_salary_structure_data():
    try:
        today = frappe.utils.today()
        cdatetime = frappe.utils.now()

    
        last_entry = frappe.db.sql("""
            SELECT file
            FROM `tabUpload Salary` WHERE status = 'pending'
            ORDER BY creation DESC
            LIMIT 1
        """, as_dict=True)
        if not last_entry:
            return
        file_path = last_entry[0].file.strip()
        owner = frappe.db.sql("""
            SELECT owner
            FROM `tabUpload Salary` WHERE status = 'pending'
            ORDER BY creation DESC
            LIMIT 1
        """, as_dict=True)
        by = owner[0].owner.strip()
        from_entry = frappe.db.sql("""
            SELECT fromdate
            FROM `tabUpload Salary` WHERE status = 'pending'
            ORDER BY creation DESC
            LIMIT 1
        """, as_dict=True)
        from_date = str(from_entry[0]['fromdate'])
        to_entry = frappe.db.sql("""
            SELECT todate
            FROM `tabUpload Salary` WHERE status = 'pending'
            ORDER BY creation DESC
            LIMIT 1
        """, as_dict=True)
        
        to_date = str(to_entry[0]['todate'])
        content = frappe.get_doc("File", {"file_url": file_path}).get_content()
        path = frappe.get_site_path() + file_path
        df = pd.read_excel(path)
        basic = frappe.get_doc('Salary Component', 'Basic')
        bns = frappe.get_doc('Salary Component', 'Bonus')
        hra = frappe.get_doc('Salary Component', 'House Rent Allowance')
        ma = frappe.get_doc('Salary Component', 'Medical Allowance')
        ca = frappe.get_doc('Salary Component', 'Conveyance Allowance')
        pfc = frappe.get_doc('Salary Component', 'Provident Fund')
        sa = frappe.get_doc('Salary Component', 'Special Allowance')
        ea = frappe.get_doc('Salary Component', 'Expense Allowance')
        epfee = frappe.get_doc('Salary Component', 'EPF Employee @ 12%')
        epfer = frappe.get_doc('Salary Component', 'EPF Employer @ 12%')
        mip = frappe.get_doc('Salary Component', 'Medical Insurance Premium')
        lwf = frappe.get_doc('Salary Component', 'Labour Welfare Fund')
        gty = frappe.get_doc('Salary Component', 'Gratuity')
        vp = frappe.get_doc('Salary Component', 'Variable Pay')
        esia = frappe.get_doc('Salary Component', 'ESI @ 0.75%')
        esib = frappe.get_doc('Salary Component', 'ESI @ 3.25%')
        taxc = frappe.get_doc('Salary Component', 'Professional Tax')
        asa = frappe.get_doc('Salary Component', 'Salary Advance')
        tdsc = frappe.get_doc('Salary Component', 'TDS')
        docstatus = 1
        insertdata = []
        
        for index, row in df.iterrows():
            if index == 1:
                for col_index, col_value in enumerate(row):
                    if col_value == 'Particulars':
                        nameIndex = col_index
                    elif col_value == 'Employee Number':
                        codeIndex = col_index
                    elif col_value == 'Bank Name':
                        bankIndex = col_index
                    elif col_value == 'Account Number':
                        accountIndex = col_index
                    elif col_value == 'Universal Account':
                        universalAccountIndex = col_index
                    elif col_value == 'PF Account Number':
                        pfAccountIndex = col_index
                    elif col_value == 'ESI Number':
                        esiNumberIndex = col_index
                    elif col_value == 'Designation':
                        designationIndex = col_index
                    elif col_value == 'Function':
                        functionIndex = col_index
                    elif col_value == 'Date of Joining':
                        joiningDateIndex = col_index
                    elif col_value == 'Date of Resignation':
                        resignationDateIndex = col_index
                    elif col_value == 'BDA':
                        bdaIndex = col_index
                    elif col_value == 'BONUS ALLOWANCE':
                        bonusAllowanceIndex = col_index
                    elif col_value == 'CA':
                        caIndex = col_index
                    elif col_value == 'EPF':
                        epfIndex = col_index
                    elif col_value == 'ESIC ALLOWANCE':
                        esicAllowanceIndex = col_index
                    elif col_value == 'EXPENSES REIMBURSEMENT':
                        expensesReimbursementIndex = col_index
                    elif col_value == 'GRATUITY 2020-21':
                        gratuityIndex = col_index
                    elif col_value == 'HRA':
                        hraIndex = col_index
                    elif col_value == 'Stipend':
                        stipendIndex = col_index
                    elif col_value == 'LWF ALLOWANCE':
                        lwfAllowanceIndex = col_index
                    elif col_value == 'MED ALLOWANCE':
                        medAllowanceIndex = col_index
                    elif col_value == 'ARREARS':
                        arrearsIndex = col_index
                    elif col_value == 'SA':
                        saIndex = col_index
                    elif col_value == 'Variable Pay':
                        variablePayIndex = col_index
                    elif col_value == 'Total Earnings':
                        totalEarningsIndex = col_index
                    elif col_value == 'EPF EMPLOYEE  @ 12%':
                        epfEmployeeIndex = col_index
                    elif col_value == 'EPF EMPLOYER @ 12%':
                        epfEmployerIndex = col_index
                    elif col_value == 'ESI @ 0.75%':
                        esiEmployeeIndex = col_index
                    elif col_value == 'ESI @ 3.25%':
                        esiEmployerIndex = col_index
                    elif col_value == 'LABOUR WELFARE FUND':
                        lwfIndex = col_index
                    elif col_value == 'MIP':
                        mipIndex = col_index
                    elif col_value == 'PROFESSIONAL TAX':
                        professionalTaxIndex = col_index
                    elif col_value == 'STAFF SALARY ADVANCE':
                        staffSalaryAdvanceIndex = col_index
                    elif col_value == 'TDS':
                        tdsIndex = col_index
                    elif col_value == 'Total Deductions':
                        totalDeductionsIndex = col_index
                    elif col_value == 'Net Amount':
                        netAmountIndex = col_index


            if index >= 2:
                nameuser = row[nameIndex]
                employee_code = str(row[codeIndex]).strip()
                employee_code = employee_code.replace("_x000D_", "")
                if employee_code != 'nan':
                    #logger.info(employee_code)
                    ssname = 'salary_structure_' + str(employee_code) + '_' + str(cdatetime)
                    parent_id = 'salary_structure_' + str(employee_code) + '_' + str(cdatetime)
                    totalearning = row[totalEarningsIndex]
                    totaldeduction = row[totalDeductionsIndex]
                    netpay = row[netAmountIndex]
                
                
                    
                    salary_structure = frappe.new_doc('Salary Structure')
                    salary_structure.name = ssname
                    salary_structure.docstatus = docstatus
                    salary_structure.net_pay = netpay
                    salary_structure.total_deduction = totaldeduction
                    salary_structure.total_earning = totalearning
                    salary_structure.insert()
                    
                    #basic
                    
                    if str(row[bdaIndex]) != 'nan':
                        name = basic.salary_component_abbr + str(employee_code) + '_' + str(cdatetime)

                        insertdata.append(
                            (
                            name,
                            basic.name,
                            basic.salary_component_abbr,
                            basic.type.lower() + 's',
                            'Salary Structure',
                            row[bdaIndex],
                            docstatus,
                            parent_id
                            )
                        )
                    
                    #hra
                    
                    if str(row[hraIndex]) != 'nan':
                        name = hra.salary_component_abbr + str(employee_code) + '_' + str(cdatetime)

                        insertdata.append(
                            (
                            name,
                            hra.name,
                            hra.salary_component_abbr,
                            hra.type.lower() + 's',
                            'Salary Structure',
                            row[hraIndex],
                            docstatus,
                            parent_id
                            )
                        )
                    
                    #ma
                    if str(row[medAllowanceIndex]) != 'nan':
                        name = ma.salary_component_abbr + str(employee_code) + '_' + str(cdatetime)

                        insertdata.append(
                            (
                            name,
                            ma.name,
                            ma.salary_component_abbr,
                            ma.type.lower() + 's',
                            'Salary Structure',
                            row[medAllowanceIndex],
                            docstatus,
                            parent_id
                            )
                        )
                    
                    #ca
                    if str(row[caIndex]) != 'nan':
                        name = ca.salary_component_abbr + str(employee_code) + '_' + str(cdatetime)

                        insertdata.append(
                            (
                            name,
                            ca.name,
                            ca.salary_component_abbr,
                            ca.type.lower() + 's',
                            'Salary Structure',
                            row[caIndex],
                            docstatus,
                            parent_id
                            )
                        )
                    
                    #pfc

                    if str(row[epfIndex]) != 'nan':
                        name = pfc.salary_component_abbr + str(employee_code) + '_' + str(cdatetime)

                        insertdata.append(
                            (
                            name,
                            pfc.name,
                            pfc.salary_component_abbr,
                            pfc.type.lower() + 's',
                            'Salary Structure',
                            row[epfIndex],
                            docstatus,
                            parent_id
                            )
                        )
                    
                    #sa
                    if str(row[saIndex]) != 'nan':
                        name = sa.salary_component_abbr + str(employee_code) + '_' + str(cdatetime)

                        insertdata.append(
                            (
                            name,
                            sa.name,
                            sa.salary_component_abbr,
                            sa.type.lower() + 's',
                            'Salary Structure',
                            row[saIndex],
                            docstatus,
                            parent_id
                            )
                        )
                    
                    #ea
                    if str(row[expensesReimbursementIndex]) != 'nan':
                        name = ea.salary_component_abbr + str(employee_code) + '_' + str(cdatetime)

                        insertdata.append(
                            (
                            name,
                            ea.name,
                            ea.salary_component_abbr,
                            ea.type.lower() + 's',
                            'Salary Structure',
                            row[expensesReimbursementIndex],
                            docstatus,
                            parent_id
                            )
                        )
                    
                    
                    #epfee
                    if str(row[epfEmployeeIndex]) != 'nan':
                        name = epfee.salary_component_abbr + str(employee_code) + '_' + str(cdatetime)

                        insertdata.append(
                            (
                            name,
                            epfee.name,
                            epfee.salary_component_abbr,
                            epfee.type.lower() + 's',
                            'Salary Structure',
                            row[epfEmployeeIndex],
                            docstatus,
                            parent_id
                            )
                        )
                    
                    #epfer

                    if str(row[epfEmployerIndex]) != 'nan':
                        name = epfer.salary_component_abbr + str(employee_code) + '_' + str(cdatetime)

                        insertdata.append(
                            (
                            name,
                            epfer.name,
                            epfer.salary_component_abbr,
                            epfer.type.lower() + 's',
                            'Salary Structure',
                            row[epfEmployerIndex],
                            docstatus,
                            parent_id
                            )
                        )
                    
                    #mip

                    if str(row[mipIndex]) != 'nan':
                        name = mip.salary_component_abbr + str(employee_code) + '_' + str(cdatetime)

                        insertdata.append(
                            (
                            name,
                            mip.name,
                            mip.salary_component_abbr,
                            mip.type.lower() + 's',
                            'Salary Structure',
                            row[mipIndex],
                            docstatus,
                            parent_id
                            )
                        )
                    
                    #lwf

                    if str(row[lwfIndex]) != 'nan':
                        name = lwf.salary_component_abbr + str(employee_code) + '_' + str(cdatetime)

                        insertdata.append(
                            (
                            name,
                            lwf.name,
                            lwf.salary_component_abbr,
                            lwf.type.lower() + 's',
                            'Salary Structure',
                            row[lwfIndex],
                            docstatus,
                            parent_id
                            )
                        )
                    
                    
                    gratuityamt = row[gratuityIndex]
                    if str(gratuityamt) != 'nan':
                        gratuity = float(gratuityamt)
                        
                        if gratuity and gratuity > 0:
                            name = gty.salary_component_abbr + str(employee_code) + '_' + str(cdatetime)
                        
                            #gty

                            insertdata.append(
                                (
                                name,
                                gty.name,
                                gty.salary_component_abbr,
                                gty.type.lower() + 's',
                                'Salary Structure',
                                gratuity,
                                docstatus,
                                parent_id
                                )
                            )
                        
                        
                    variablpayamt = row[variablePayIndex]
                    if str(variablpayamt) != 'nan':
                        variablpay = float(variablpayamt)
                        
                        if variablpay and variablpay > 0:
                            name = vp.salary_component_abbr + str(employee_code) + '_' + str(cdatetime)
                        
                            #vp
                            

                            insertdata.append(
                                (
                                name,
                                vp.name,
                                vp.salary_component_abbr,
                                vp.type.lower() + 's',
                                'Salary Structure',
                                variablpay,
                                docstatus,
                                parent_id
                                )
                            )
                        
                    esioneamt = row[esiEmployeeIndex]
                    if str(esioneamt) != 'nan':
                        esione = float(esioneamt)
                        if esione and esione > 0:
                            name = esia.salary_component_abbr + str(employee_code) + '_' + str(cdatetime)
                        
                            #esia

                            insertdata.append(
                                (
                                name,
                                esia.name,
                                esia.salary_component_abbr,
                                esia.type.lower() + 's',
                                'Salary Structure',
                                esione,
                                docstatus,
                                parent_id
                                )
                            )
                        
                    esitamt = row[esiEmployerIndex]
                    if str(esitamt) != 'nan':
                        esit = float(esitamt)
                        
                        if esit and esit > 0:
                            name = esib.salary_component_abbr + str(employee_code) + '_' + str(cdatetime)
                        
                            #esib

                            insertdata.append(
                                (
                                name,
                                esib.name,
                                esib.salary_component_abbr,
                                esib.type.lower() + 's',
                                'Salary Structure',
                                esit,
                                docstatus,
                                parent_id
                                )
                            )
                        
                    taxamt = row[professionalTaxIndex]
                    if str(taxamt) != 'nan':
                        tax = float(taxamt)
                        
                        if tax and tax > 0:
                            name = taxc.salary_component_abbr + str(employee_code) + '_' + str(cdatetime)
                        
                            #taxc

                            insertdata.append(
                                (
                                name,
                                taxc.name,
                                taxc.salary_component_abbr,
                                taxc.type.lower() + 's',
                                'Salary Structure',
                                tax,
                                docstatus,
                                parent_id
                                )
                            )
                        
                    advamt = row[staffSalaryAdvanceIndex]
                    if str(advamt) != 'nan':
                        adv = float(advamt)
                        
                        if adv and adv > 0:
                            name = asa.salary_component_abbr + str(employee_code) + '_' + str(cdatetime)
                        
                            #asa

                            insertdata.append(
                                (
                                name,
                                asa.name,
                                asa.salary_component_abbr,
                                asa.type.lower() + 's',
                                'Salary Structure',
                                adv,
                                docstatus,
                                parent_id
                                )
                            )
                        
                    tdsamt = row[tdsIndex]
                    if str(tdsamt) != 'nan':
                        tds = float(tdsamt)
                        
                        if tds and tds > 0:
                            name = tdsc.salary_component_abbr + str(employee_code) + '_' + str(cdatetime)
                        
                            #tdsc

                            insertdata.append(
                                (
                                name,
                                tdsc.name,
                                tdsc.salary_component_abbr,
                                tdsc.type.lower() + 's',
                                'Salary Structure',
                                tds,
                                docstatus,
                                parent_id
                                )
                            )
                    advamt = row[staffSalaryAdvanceIndex]
                    if str(advamt) != 'nan':
                        adv = float(advamt)
                        
                        if adv and adv > 0:
                            name = asa.salary_component_abbr + str(employee_code) + '_' + str(cdatetime)
                        
                            #asa

                            insertdata.append(
                                (
                                name,
                                asa.name,
                                asa.salary_component_abbr,
                                asa.type.lower() + 's',
                                'Salary Structure',
                                adv,
                                docstatus,
                                parent_id
                                )
                            )
                        
                    bnsamt = row[bonusAllowanceIndex]
                    if str(bnsamt) != 'nan':
                        bonus = float(bnsamt)
                        
                        if bonus and bonus > 0:
                            name = bns.salary_component_abbr + str(employee_code) + '_' + str(cdatetime)
                        
                            #bns

                            insertdata.append(
                                (
                                name,
                                bns.name,
                                bns.salary_component_abbr,
                                bns.type.lower() + 's',
                                'Salary Structure',
                                bonus,
                                docstatus,
                                parent_id
                                )
                            )
                    
                    employee = frappe.get_doc('Employee', employee_code)
                    user_id = employee.user_id
                    
                    if str(row[accountIndex]) != 'nan' and str(row[pfAccountIndex]) != 'nan' and str(row[bankIndex]) != 'nan':
                        employee.bank_ac_no = row[accountIndex]
                        employee.bank_name = row[bankIndex]
                        employee.provident_fund_account = row[pfAccountIndex]
                        employee.save()
                    user = frappe.get_doc('User', user_id)
                    user.salary_structure = name
                    user.save()
                    
                    
                    salary_structure_assignment = frappe.new_doc('Salary Structure Assignment')
                    salary_structure_assignment.employee = employee_code
                    salary_structure_assignment.docstatus = docstatus
                    salary_structure_assignment.salary_structure = ssname
                    joining = employee.date_of_joining
                    date_format = "%Y-%m-%d"
                    date = datetime.datetime.strptime(from_date, date_format)
                    frm = date.date()
                    if frm > joining:
                        salary_structure_assignment.from_date = from_date
                    else:
                        salary_structure_assignment.from_date = joining
                    now = datetime.datetime.now()
                    current_date = now.date()
                    relieving = employee.relieving_date
                    if relieving:
                        if(relieving < current_date):
                            salary_structure_assignment.save()
                    else:
                        salary_structure_assignment.save()
                    
                else:
                    break
                
        frappe.db.bulk_insert("Salary Detail", fields=["name","salary_component","abbr","parentfield","parenttype","amount","docstatus","parent"], values=set(insertdata))   
        update_query = """
                UPDATE `tabUpload Salary`
                SET status = %s
                WHERE status = %s
            """
        frappe.db.sql(update_query, ('completed', 'pending'))
        frappe.sendmail(
          recipients=[by, "email@gmail.com"],
          reference_doctype="User",
          reference_name="Administrator",
          subject="Salary Import",
          message="Salary import is completed!"
        )
        frappe.db.commit()
    except Exception:

        frappe.db.rollback()
        update_query = """
                UPDATE `tabUpload Salary`
                SET status = %s , log = %s
                WHERE status = %s
            """
        frappe.db.sql(update_query, ('failed',frappe.get_traceback(),'pending'))
        frappe.sendmail(
          recipients=[by, "email@gmail.com"],
          reference_doctype="User",
          reference_name="Administrator",
          subject="Salary Import",
          message="Salary import has failed. Check logs! "+frappe.get_traceback()
        )

def insert_salary_structure():
    frappe.enqueue('your-path', queue='short', timeout=2000)

