# -*- coding: utf-8 -*-
{
    'name': 'UCS TDS Report',
    'version': '19.0.1.0.0',
    'summary': 'TDS Transaction Report Generator in Excel',
    'description': """
        Generates consolidated TDS transaction details (Excel) by selecting a date period and relevant TDS accounts.
    """,
    'category': 'Accounting/Reporting',
    'author': 'Uncanny Consulting Services LLP',
    'depends': ['account', 'l10n_in'],
    'data': [
        'security/security_groups.xml',
        'security/ir.model.access.csv',
        'wizard/tds_report_wizard_view.xml',
    ],
    'images': ["static/description/banner.gif"],
    'installable': True,
    'application': False,
    'license': 'LGPL-3',
    'price': 50,
    'currency': 'USD',
}
