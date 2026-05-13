import importlib.util

from odoo import fields, models

_ACCOUNT_ASSET_AVAILABLE = importlib.util.find_spec("odoo.addons.account_asset") is not None


class KolacAccountSubaccountMap(models.Model):
    _name = "kolac.account.subaccount.map"
    _description = "Kolac Subaccount Mapping"
    _rec_name = "source_code"
    _order = "source_code"

    MAPPING_TYPES = [
        ("direct_account", "Cuenta directa"),
        ("expense_nature", "Gasto por naturaleza"),
        ("expense_nature_by_supplier", "Gasto por naturaleza y proveedor"),
        ("grant_or_becario_payable", "Acreedor becarios/subvenciones"),
        ("partner_customer", "Cliente"),
        ("partner_supplier", "Proveedor"),
        ("partner_creditor", "Acreedor"),
        ("asset", "Activo"),
        ("accumulated_depreciation", "Amortización acumulada"),
        ("tax_output", "IVA repercutido"),
        ("tax_input", "IVA soportado"),
        ("bank", "Banco"),
        ("cash", "Caja"),
        ("review_required", "Revisión"),
    ]

    company_id = fields.Many2one(
        "res.company",
        string="Compañía",
        required=True,
        default=lambda self: self.env.company,
    )
    source_code = fields.Char(string="Subcuenta origen", required=True)
    source_name = fields.Char(string="Nombre origen")
    source_origin = fields.Selection(
        [("journal", "Libro diario"), ("account_list", "Listado de cuentas"), ("account_plan", "Plan de cuentas"), ("manual", "Manual")],
        string="Origen principal",
        default="manual",
    )
    target_account_id = fields.Many2one(
        "account.account",
        string="Cuenta Odoo",
        required=True,
        domain="[('company_ids', 'in', [company_id])]",
    )
    target_account_code = fields.Char(string="Código Odoo", related="target_account_id.code", store=True)
    mapping_type = fields.Selection(MAPPING_TYPES, string="Tipo de mapeo", default="direct_account", required=True)
    partner_id = fields.Many2one("res.partner", string="Contacto")
    if _ACCOUNT_ASSET_AVAILABLE:
        asset_id = fields.Many2one("account.asset", string="Activo")
    tax_id = fields.Many2one("account.tax", string="Impuesto")
    action = fields.Selection(
        [("reuse", "Reutilizar"), ("create_partner", "Crear contacto"), ("create_asset", "Crear activo"), ("review", "Revisar")],
        string="Acción",
        default="reuse",
    )
    state = fields.Selection(
        [("auto", "Automático"), ("manual", "Manual"), ("review", "Revisión")],
        string="Estado",
        default="auto",
    )
    review_required = fields.Boolean(string="Revisión requerida")
    warning_message = fields.Text(string="Advertencias")
    import_batch_id = fields.Many2one("kolac.account.import.batch", string="Lote origen", readonly=True)
    note = fields.Char(string="Notas")

    _unique_source_code = models.Constraint(
        "UNIQUE(company_id, source_code)",
        "Ya existe una equivalencia para esa subcuenta y compañía.",
    )
