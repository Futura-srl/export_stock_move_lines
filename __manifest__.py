{
    'name': 'export_stock_move_line',
    'version': '16',
    'author': "Luca Cocozza",
    'application': True,
    'description': "Esportazione stock move line.",
     'data': [
        # # Settaggi per accesso ai contenuti
        'security/ir.model.access.csv',
        # Cron job
        'data/cron.xml',
    ],
}
