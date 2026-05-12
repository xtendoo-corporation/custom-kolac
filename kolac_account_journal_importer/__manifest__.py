{
    "name": "Kolac Account Journal Importer",
    "version": "19.0.1.0.0",
    "category": "Accounting",
    "summary": "Importa apuntes contables desde Excel",
    "author": "Kolac",
    "depends": ["account", "account_asset"],
    "data": [
        "security/ir.model.access.csv",
        "views/account_asset_views.xml",
        "views/account_import_batch_views.xml",
        "views/account_subaccount_map_views.xml",
        "views/account_journal_import_wizard_views.xml",
    ],
    "installable": True,
    "application": False,
    "license": "LGPL-3",
}

