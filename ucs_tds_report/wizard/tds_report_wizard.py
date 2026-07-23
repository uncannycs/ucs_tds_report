# -*- coding: utf-8 -*-

from odoo import models, fields, api, _
from odoo.exceptions import UserError, ValidationError
import io
import xlsxwriter
import base64
import re
import calendar
import pytz
from datetime import date, datetime
from collections import defaultdict

class TDSReportWizard(models.TransientModel):
    _name = 'tds.report.wizard'
    _description = 'TDS Transaction Report Wizard'

    date_from = fields.Date(string='Start Date', required=True)
    date_to = fields.Date(string='End Date', required=True)
    tds_account_ids = fields.Many2many(
        'account.account',
        string='TDS Accounts',
        required=True,
        domain="[('account_type', 'in', ('asset_receivable', 'liability_payable', 'liability_current', 'asset_current', 'expense', 'income'))]"
    )

    @api.model
    def default_get(self, fields_list):
        res = super(TDSReportWizard, self).default_get(fields_list)
        today = date.today()
        first_day = today.replace(day=1)
        last_day_number = calendar.monthrange(today.year, today.month)[1]
        last_day = today.replace(day=last_day_number)
        res.update({
            'date_from': first_day,
            'date_to': last_day,
        })
        return res

    def get_tds_section(self, tax_name, line_label, move_ref):
        """ Extract TDS/TCS section from tax name, line label, or entry reference """
        search_texts = [tax_name, line_label, move_ref]
        # Regex to match common Indian TDS/TCS section codes (194X, 206X)
        pattern = r'\b(194[A-Z0-9\(\)]*|206[A-Z0-9\(\)]*)\b'
        
        for text in search_texts:
            if not text:
                continue
            if isinstance(text, dict):
                text_str = text.get('en_US') or (list(text.values())[0] if text.values() else '')
            else:
                text_str = str(text)
            
            match = re.search(pattern, text_str, re.IGNORECASE)
            if match:
                return match.group(1).upper()
        return "N/A"

    def action_generate_excel_report(self):
        self.ensure_one()
        if self.date_from > self.date_to:
            raise UserError(_("The 'Start Date' must be earlier than or equal to the 'End Date'."))

        # SQL Query to fetch the TDS journal lines and related details
        query = """
            SELECT 
                aml.date,
                am.name AS move_name,
                am.ref AS move_ref,
                aml.name AS line_label,
                t.name AS tax_name,
                rp.name AS partner_name,
                rp.l10n_in_pan AS pan,
                aml.credit AS balance_amount,
                (SELECT price_subtotal FROM account_move_line aml_base
                 JOIN account_move_line_account_tax_rel tax_rel ON tax_rel.account_move_line_id = aml_base.id
                 WHERE aml_base.move_id = aml.move_id AND tax_rel.account_tax_id = aml.tax_line_id
                 LIMIT 1) AS base_amount,
                aa.id AS account_id,
                aa.name AS account_name
            FROM account_move_line aml
            JOIN account_move am ON aml.move_id = am.id
            JOIN account_account aa ON aml.account_id = aa.id
            LEFT JOIN account_tax t ON aml.tax_line_id = t.id
            LEFT JOIN res_partner rp ON aml.partner_id = rp.id
            WHERE aml.account_id IN %s
              AND aml.date >= %s
              AND aml.date <= %s
              AND am.state = 'posted'
              AND aml.credit > 0
            ORDER BY aml.date ASC, aml.id ASC;
        """
        
        self.env.cr.execute(query, [tuple(self.tds_account_ids.ids), self.date_from, self.date_to])
        rows = self.env.cr.dictfetchall()

        if not rows:
            raise UserError(_("No posted journal entries found for the selected period and accounts."))

        # Get local generation date
        user_tz = self.env.user.tz or 'Asia/Kolkata'
        now_utc = datetime.utcnow()
        now_local = pytz.utc.localize(now_utc).astimezone(pytz.timezone(user_tz))
        gen_date_str = now_local.strftime('%d/%m/%Y %H:%M:%S')

        # Setup in-memory workbook
        output = io.BytesIO()
        workbook = xlsxwriter.Workbook(output, {'in_memory': True})

        # Formats
        title_format = workbook.add_format({
            'bold': True,
            'font_size': 10,
            'align': 'center',
            'valign': 'vcenter',
            'font_color': '#1B365D'
        })
        
        subtitle_format = workbook.add_format({
            'bold': True,
            'align': 'center',
            'font_size': 12
        })

        info_format = workbook.add_format({
            'bold': True,
            'font_size': 10,
            'align': 'left',
            'font_color': '#555555'
        })

        header_format = workbook.add_format({
            'bold': True,
            'font_size': 11,
            'align': 'center',
            'valign': 'vcenter',
            'bg_color': '#F2F2F2',
            'font_color': '#000000',
            'border': 1,
            'border_color': '#D3D3D3'
        })

        data_left = workbook.add_format({
            'align': 'left',
            'valign': 'vcenter',
            'border': 1,
            'border_color': '#EEEEEE'
        })

        data_center = workbook.add_format({
            'align': 'center',
            'valign': 'vcenter',
            'border': 1,
            'border_color': '#EEEEEE'
        })

        amount_format = workbook.add_format({
            'num_format': '₹ #,##,##0.00',
            'align': 'right',
            'valign': 'vcenter',
            'border': 1,
            'border_color': '#EEEEEE'
        })

        total_format = workbook.add_format({
            'num_format': '₹ #,##,##0.00',
            'bold': True,
            'align': 'right',
            'valign': 'vcenter',
            'top': 1,
            'bottom': 6,
            'border_color': '#D3D3D3'
        })

        total_label_format = workbook.add_format({
            'bold': True,
            'align': 'right',
            'valign': 'vcenter',
            'top': 1,
            'bottom': 6,
            'border_color': '#D3D3D3'
        })

        data_left_bold = workbook.add_format({
            'bold': True,
            'align': 'left',
            'valign': 'vcenter',
            'border': 1,
            'border_color': '#EEEEEE'
        })

        amount_bold_format = workbook.add_format({
            'num_format': '₹ #,##,##0.00',
            'bold': True,
            'align': 'right',
            'valign': 'vcenter',
            'border': 1,
            'border_color': '#EEEEEE'
        })

        # Process and group data
        # 1. Summary Grouping: {(account_code, account_name, partner_name, pan): {'amount': 0.0, 'base_amount': 0.0}}
        summary_data = defaultdict(lambda: {'amount': 0.0, 'base_amount': 0.0})
        # 2. Monthly Grouping: {(year, month, month_str): [rows]}
        monthly_data = defaultdict(list)

        account_codes = {acc.id: acc.code for acc in self.tds_account_ids}

        for r in rows:
            line_date = r['date']
            acc_code = account_codes.get(r['account_id'], '')
            r['account_code'] = acc_code
            acc_name = r['account_name']
            if isinstance(acc_name, dict):
                acc_name = acc_name.get('en_US') or (list(acc_name.values())[0] if acc_name.values() else '')
            else:
                acc_name = str(acc_name) if acc_name else ''

            partner = r['partner_name'] or 'N/A'
            if isinstance(partner, dict):
                partner = partner.get('en_US') or (list(partner.values())[0] if partner.values() else 'N/A')
            else:
                partner = str(partner)

            pan = r['pan'] or 'N/A'
            amount = r['balance_amount'] or 0.0
            base_amount = r['base_amount'] or 0.0

            # Summary aggregation
            summary_key = (acc_code, acc_name, partner, pan)
            summary_data[summary_key]['amount'] += amount
            summary_data[summary_key]['base_amount'] += base_amount

            # Save strings to dict for monthly sheets
            r['account_name_str'] = acc_name
            r['partner_name_str'] = partner

            # Monthly grouping
            month_key = (line_date.year, line_date.month, line_date.strftime('%B %Y'))
            monthly_data[month_key].append(r)

        # Calculate account totals for cached values in the SUMIF formulas
        account_totals = defaultdict(lambda: {'amount': 0.0, 'base_amount': 0.0})
        total_amount = 0.0
        total_base_amount = 0.0
        for key, vals in summary_data.items():
            acc_code, acc_name, partner, pan = key
            account_totals[(acc_code, acc_name)]['amount'] += vals['amount']
            account_totals[(acc_code, acc_name)]['base_amount'] += vals['base_amount']
            total_amount += vals['amount']
            total_base_amount += vals['base_amount']

        # Get unique accounts in summary
        unique_accounts = sorted(list(set((k[0], k[1]) for k in summary_data.keys())), key=lambda k: k[0])
        K = len(unique_accounts)
        sorted_summary_keys = sorted(summary_data.keys(), key=lambda k: (k[0], k[2]))
        P = len(sorted_summary_keys)

        # Calculate row boundaries for partner table (needed for SUMIF formulas in account table)
        # Note: We add K summary rows to the partner table, making the total rows P + K.
        detail_data_start_row = 8 + K + 5 # 1-based index
        detail_data_end_row = 8 + K + 4 + P + K # 1-based index

        # ----------------------------------------------------
        # SHEET 1: Summary Sheet
        # ----------------------------------------------------
        summary_sheet = workbook.add_worksheet('Summary')
        summary_sheet.merge_range('A1:E1', 'TECH-LONG PACKAGING MACHINERY INDIA PRIVATE LIMITED', title_format)
        summary_sheet.merge_range('A2:E2', 'Consolidated TDS Transaction Summary Report', subtitle_format)
        
        from_str = self.date_from.strftime('%d/%m/%Y')
        to_str = self.date_to.strftime('%d/%m/%Y')
        summary_sheet.merge_range('A4:E4', f"Period: {from_str} to {to_str}", info_format)
        summary_sheet.merge_range('A5:E5', f"Generated By: {self.env.user.name} | Date: {gen_date_str}", info_format)

        # Column widths
        summary_sheet.set_column('A:A', 35) # TDS Account
        summary_sheet.set_column('B:B', 50) # Partner Name
        summary_sheet.set_column('C:C', 20) # PAN Number
        summary_sheet.set_column('D:D', 20) # BASE Amount
        summary_sheet.set_column('E:E', 20) # Amount

        # 1. Account-wise Summary Table
        summary_sheet.set_row(7, 28)
        summary_sheet.merge_range('A8:C8', 'TDS Account', header_format)
        summary_sheet.write('D8', 'BASE Amount', header_format)
        summary_sheet.write('E8', 'Amount', header_format)

        row = 8
        sum_start_row = row + 1 # 1-based index
        for acc_code, acc_name in unique_accounts:
            summary_sheet.set_row(row, 20)
            summary_sheet.merge_range(row, 0, row, 2, f"{acc_code} - {acc_name}", data_left)
            # SUMIF formulas
            base_formula = f"=SUMIF(A${detail_data_start_row}:A${detail_data_end_row}, \"  \" & A{row+1}, D${detail_data_start_row}:D${detail_data_end_row})"
            amount_formula = f"=SUMIF(A${detail_data_start_row}:A${detail_data_end_row}, \"  \" & A{row+1}, E${detail_data_start_row}:E${detail_data_end_row})"
            acc_vals = account_totals[(acc_code, acc_name)]
            summary_sheet.write_formula(row, 3, base_formula, amount_format, acc_vals['base_amount'])
            summary_sheet.write_formula(row, 4, amount_formula, amount_format, acc_vals['amount'])
            row += 1
        sum_end_row = row # 1-based index

        # Account-wise Total
        summary_sheet.set_row(row, 22)
        summary_sheet.merge_range(row, 0, row, 2, 'Total', total_label_format)
        summary_sheet.write_formula(row, 3, f"=SUM(D{sum_start_row}:D{sum_end_row})", total_format, total_base_amount)
        summary_sheet.write_formula(row, 4, f"=SUM(E{sum_start_row}:E{sum_end_row})", total_format, total_amount)
        row += 1

        # Space and Section Header for Partner Breakdown
        row += 1
        summary_sheet.merge_range(row, 0, row, 4, 'Partner-wise Breakdown', subtitle_format)
        row += 1

        # 2. Partner-wise Breakdown Table
        summary_sheet.set_row(row, 28)
        sum_headers = ["TDS Account", "Partner Name", "PAN Number", "BASE Amount", "Amount"]
        for col_idx, h in enumerate(sum_headers):
            summary_sheet.write(row, col_idx, h, header_format)
        
        # Apply outline settings so that collapse buttons are above
        summary_sheet.outline_settings(visible=True, symbols_below=False, symbols_right=True, auto_style=False)
        partner_header_row = row
        row += 1

        # Group and write partner rows
        partner_groups = defaultdict(list)
        for key in sorted_summary_keys:
            acc_code, acc_name, partner, pan = key
            partner_groups[(acc_code, acc_name)].append(key)

        for (acc_code, acc_name), keys in partner_groups.items():
            acc_vals = account_totals[(acc_code, acc_name)]
            
            # Account Summary Row (level 1)
            summary_sheet.set_row(row, 20, None, {'level': 1})
            summary_sheet.write(row, 0, f"{acc_code} - {acc_name}", data_left_bold)
            summary_sheet.write(row, 1, "", data_left_bold)
            summary_sheet.write(row, 2, "", data_center)
            
            acc_summary_row = row
            row += 1
            
            # Partner Detail Rows (level 2)
            detail_start_excel_row = row + 1
            for key in keys:
                _acc_code, _acc_name, partner, pan = key
                vals = summary_data[key]
                amount = vals['amount']
                base_amount = vals['base_amount']
                
                # Check for empty base amount
                base_val = base_amount if base_amount else ""
                
                summary_sheet.set_row(row, 20, None, {'level': 2})
                summary_sheet.write(row, 0, f"  {acc_code} - {acc_name}", data_left)
                summary_sheet.write(row, 1, partner, data_left)
                summary_sheet.write(row, 2, pan, data_center)
                if base_val != "":
                    summary_sheet.write(row, 3, base_val, amount_format)
                else:
                    summary_sheet.write(row, 3, "", data_left)
                summary_sheet.write(row, 4, amount, amount_format)
                row += 1
            detail_end_excel_row = row
            
            # Write SUM formulas to Columns D and E of Account Summary row
            summary_sheet.write_formula(acc_summary_row, 3, f"=SUM(D{detail_start_excel_row}:D{detail_end_excel_row})", amount_bold_format, acc_vals['base_amount'])
            summary_sheet.write_formula(acc_summary_row, 4, f"=SUM(E{detail_start_excel_row}:E{detail_end_excel_row})", amount_bold_format, acc_vals['amount'])

        partner_end_row = row
        summary_sheet.autofilter(partner_header_row, 0, partner_end_row - 1, 4)

        # Partner-wise Grand Total Row (level 1)
        account_total_excel_row = sum_end_row + 1 # 1-based index of Account-wise Total row
        summary_sheet.set_row(row, 22, None, {'level': 1})
        summary_sheet.merge_range(row, 0, row, 2, 'Grand Total', total_label_format)
        summary_sheet.write_formula(row, 3, f"=D{account_total_excel_row}", total_format, total_base_amount)
        summary_sheet.write_formula(row, 4, f"=E{account_total_excel_row}", total_format, total_amount)

        # ----------------------------------------------------
        # SHEET 2+: Monthly Sheets
        # ----------------------------------------------------
        sorted_month_keys = sorted(monthly_data.keys(), key=lambda k: (k[0], k[1]))

        for month_key in sorted_month_keys:
            year, month_num, month_str = month_key
            month_rows = monthly_data[month_key]

            # Shorten tab name if necessary (max 31 chars)
            tab_name = month_str[:31]
            month_sheet = workbook.add_worksheet(tab_name)

            # Titles
            month_sheet.merge_range('A1:H1', 'TECH-LONG PACKAGING MACHINERY INDIA PRIVATE LIMITED', title_format)
            month_sheet.merge_range('A2:H2', f'TDS Transaction Details - {month_str}', subtitle_format)
            month_sheet.merge_range('A4:H4', f"Period: {month_str}", info_format)
            month_sheet.merge_range('A5:H5', f"Generated By: {self.env.user.name} | Date: {gen_date_str}", info_format)

            # Column widths
            month_sheet.set_column('A:A', 15) # Date
            month_sheet.set_column('B:B', 22) # Entry No.
            month_sheet.set_column('C:C', 15) # TDS Section
            month_sheet.set_column('D:D', 45) # Partner Name
            month_sheet.set_column('E:E', 35) # TDS Account
            month_sheet.set_column('F:F', 20) # BASE Amount
            month_sheet.set_column('G:G', 18) # Amount
            month_sheet.set_column('H:H', 18) # PAN Number

            # Set padded row height for table header
            month_sheet.set_row(6, 28)

            # Headers
            det_headers = ["Date", "Entry No.", "TDS Section", "Partner Name", "TDS Account", "BASE Amount", "Amount", "PAN Number"]
            for col_idx, h in enumerate(det_headers):
                month_sheet.write(6, col_idx, h, header_format)

            # Sort month's rows by account code, then date
            sorted_month_rows = sorted(month_rows, key=lambda r: (r['account_code'], r['date']))
            
            # Apply outline settings so collapse buttons are above
            month_sheet.outline_settings(visible=True, symbols_below=False, symbols_right=True, auto_style=False)
            
            m_row = 7
            m_start_row = m_row + 1

            prev_acc_code = None
            for r in sorted_month_rows:
                line_date = r['date']
                entry_no = r['move_name'] or ''
                date_str = line_date.strftime('%d/%m/%Y')
                tds_section = r['line_label'] or ''
                partner_name = r['partner_name_str']
                acc_code = r['account_code']
                acc_name = r['account_name_str']
                amount = r['balance_amount'] or 0.0
                base_amount = r['base_amount'] or 0.0
                pan = r['pan'] or 'N/A'
                
                # Check for empty base amount
                base_val = base_amount if base_amount else ""

                # If this is the same account as the previous line, group it under level 1
                if prev_acc_code == acc_code:
                    month_sheet.set_row(m_row, 20, None, {'level': 1})
                else:
                    month_sheet.set_row(m_row, 20, None, {'level': 0})
                    prev_acc_code = acc_code

                month_sheet.write(m_row, 0, date_str, data_center)
                month_sheet.write(m_row, 1, entry_no, data_center)
                month_sheet.write(m_row, 2, tds_section, data_center)
                month_sheet.write(m_row, 3, partner_name, data_left)
                month_sheet.write(m_row, 4, f"{acc_code} - {acc_name}", data_left)
                if base_val != "":
                    month_sheet.write(m_row, 5, base_val, amount_format)
                else:
                    month_sheet.write(m_row, 5, "", data_left)
                month_sheet.write(m_row, 6, amount, amount_format)
                month_sheet.write(m_row, 7, pan, data_center)
                m_row += 1

            m_end_row = m_row
            month_sheet.autofilter(6, 0, m_end_row - 1, 7)

            # Month Total Row (level 0)
            month_sheet.set_row(m_row, 22, None, {'level': 0})
            month_sheet.merge_range(m_row, 0, m_row, 4, 'Total', total_label_format)
            if m_end_row >= m_start_row:
                monthly_total = sum(r['balance_amount'] or 0.0 for r in sorted_month_rows)
                monthly_base_total = sum(r['base_amount'] or 0.0 for r in sorted_month_rows)
                month_sheet.write_formula(m_row, 5, f"=SUM(F{m_start_row}:F{m_end_row})", total_format, monthly_base_total)
                month_sheet.write_formula(m_row, 6, f"=SUM(G{m_start_row}:G{m_end_row})", total_format, monthly_total)
            else:
                month_sheet.write(m_row, 5, 0.0, total_format)
                month_sheet.write(m_row, 6, 0.0, total_format)
            month_sheet.write(m_row, 7, '', total_label_format)

        # Save and return attachment
        workbook.close()
        output.seek(0)
        excel_file_base64 = base64.b64encode(output.read())

        file_name = f"TDS_Report_{self.date_from.strftime('%d%m%Y')}_{self.date_to.strftime('%d%m%Y')}.xlsx"
        
        attachment = self.env['ir.attachment'].create({
            'name': file_name,
            'type': 'binary',
            'datas': excel_file_base64,
            'mimetype': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            'res_model': self._name,
            'res_id': self.id,
        })

        return {
            'type': 'ir.actions.act_url',
            'url': f'/web/content/{attachment.id}?download=true',
            'target': 'self',
        }
