import hashlib
import importlib.util

from odoo import fields, models

_ACCOUNT_ASSET_AVAILABLE = importlib.util.find_spec("odoo.addons.account_asset") is not None


class KolacAccountImportBatch(models.Model):
    _name = "kolac.account.import.batch"
    _description = "Kolac Accounting Import Batch"
    _order = "create_date desc, id desc"

    name = fields.Char(string="Lote", required=True, default=lambda self: self._default_name())
    company_id = fields.Many2one(
        "res.company",
        string="Compañía",
        required=True,
        default=lambda self: self.env.company,
    )
    journal_id = fields.Many2one(
        "account.journal",
        string="Diario por defecto",
        required=True,
        domain="[('type', '=', 'general'), ('company_id', '=', company_id)]",
    )
    user_id = fields.Many2one(
        "res.users",
        string="Ejecutado por",
        required=True,
        default=lambda self: self.env.user,
        readonly=True,
    )
    state = fields.Selection(
        [
            ("draft", "Borrador"),
            ("analyzed", "Analizado"),
            ("done", "Importado"),
            ("error", "Con errores"),
        ],
        string="Estado",
        required=True,
        default="draft",
        readonly=True,
    )
    dry_run = fields.Boolean(string="Simulación")
    create_missing_accounts = fields.Boolean(string="Permitir crear subcuentas faltantes")
    post_moves = fields.Boolean(string="Publicar asientos válidos")

    diary_file_name = fields.Char(string="Libro Diario")
    account_list_file_name = fields.Char(string="Listado de cuentas")
    account_plan_file_name = fields.Char(string="Plan de cuentas")
    diary_file_hash = fields.Char(string="Hash Libro Diario", readonly=True)
    account_list_file_hash = fields.Char(string="Hash Listado", readonly=True)
    account_plan_file_hash = fields.Char(string="Hash Plan", readonly=True)
    import_signature = fields.Char(string="Firma de importación", readonly=True, index=True)

    mapping_line_count = fields.Integer(string="Mapeos", readonly=True)
    asset_line_count = fields.Integer(string="Activos detectados", readonly=True)
    warning_count = fields.Integer(string="Advertencias", readonly=True)
    error_count = fields.Integer(string="Errores", readonly=True)
    move_count = fields.Integer(string="Asientos creados", readonly=True)
    line_count = fields.Integer(string="Líneas creadas", readonly=True)
    review_report = fields.Text(string="Informe de revisión", readonly=True)
    imported_at = fields.Datetime(string="Importado el", readonly=True)

    trace_line_ids = fields.One2many(
        "kolac.account.import.trace",
        "batch_id",
        string="Trazabilidad",
        readonly=True,
    )
    asset_preview_line_ids = fields.One2many(
        "kolac.account.import.batch.asset",
        "batch_id",
        string="Detalle activos detectados",
        readonly=True,
    )

    def _default_name(self):
        return self.env["ir.sequence"].next_by_code("kolac.account.import.batch") or self.env._(
            "Importación contable"
        )

    @staticmethod
    def build_hash(file_content):
        if not file_content:
            return False
        return hashlib.sha256(file_content).hexdigest()

    @classmethod
    def build_signature(cls, diary_hash, account_list_hash, account_plan_hash, company_id):
        payload = "|".join(filter(None, [str(company_id), diary_hash or "", account_list_hash or "", account_plan_hash or ""]))
        return hashlib.sha256(payload.encode("utf-8")).hexdigest()


class KolacAccountImportTrace(models.Model):
    _name = "kolac.account.import.trace"
    _description = "Kolac Accounting Import Trace"
    _order = "move_date desc, id desc"

    batch_id = fields.Many2one(
        "kolac.account.import.batch",
        string="Lote",
        required=True,
        ondelete="cascade",
    )
    company_id = fields.Many2one(related="batch_id.company_id", store=True, readonly=True)
    move_id = fields.Many2one("account.move", string="Asiento Odoo", ondelete="set null")
    move_line_id = fields.Many2one("account.move.line", string="Línea Odoo", ondelete="set null")
    old_move_number = fields.Char(string="Asiento origen", required=True, index=True)
    old_move_ref = fields.Char(string="Referencia origen", index=True)
    move_date = fields.Date(string="Fecha origen", required=True, index=True)
    old_account_code = fields.Char(string="Cuenta antigua", required=True, index=True)
    old_account_name = fields.Char(string="Nombre antiguo")
    original_label = fields.Char(string="Etiqueta original")
    target_account_id = fields.Many2one("account.account", string="Cuenta Odoo", ondelete="set null")
    mapped_account_code = fields.Char(related="target_account_id.code", string="Código cuenta final", store=True, readonly=True)
    partner_id = fields.Many2one("res.partner", string="Contacto", ondelete="set null")
    if _ACCOUNT_ASSET_AVAILABLE:
        asset_id = fields.Many2one("account.asset", string="Activo", ondelete="set null")
    tax_id = fields.Many2one("account.tax", string="Impuesto", ondelete="set null")
    mapping_type = fields.Selection(
        [
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
        ],
        string="Tipo de mapeo",
    )
    source_file = fields.Char(string="Fichero origen")
    source_line = fields.Integer(string="Línea origen")
    debit = fields.Monetary(string="Debe", currency_field="company_currency_id")
    credit = fields.Monetary(string="Haber", currency_field="company_currency_id")
    note = fields.Char(string="Notas")
    company_currency_id = fields.Many2one(related="company_id.currency_id", readonly=True)


class KolacAccountImportBatchAsset(models.Model):
    _name = "kolac.account.import.batch.asset"
    _description = "Kolac Imported Asset Preview"
    _order = "id"

    batch_id = fields.Many2one(
        "kolac.account.import.batch",
        string="Lote",
        required=True,
        ondelete="cascade",
    )
    company_id = fields.Many2one(related="batch_id.company_id", store=True, readonly=True)
    source_code = fields.Char(string="Cuenta activo antigua", required=True, index=True)
    old_account_name = fields.Char(string="Nombre antiguo")
    asset_name = fields.Char(string="Activo detectado", required=True)
    target_account_id = fields.Many2one("account.account", string="Cuenta activo Odoo", ondelete="set null")
    depreciation_source_code = fields.Char(string="Cuenta amortización antigua")
    old_depreciation_account_name = fields.Char(string="Nombre amortización antigua")
    depreciation_target_account_id = fields.Many2one("account.account", string="Cuenta amortización Odoo", ondelete="set null")
    original_value = fields.Monetary(string="Valor original", currency_field="company_currency_id")
    accumulated_value = fields.Monetary(string="Amortización acumulada", currency_field="company_currency_id")
    residual_value = fields.Monetary(string="Pendiente", currency_field="company_currency_id")
    acquisition_date = fields.Date(string="Fecha adquisición")
    create_asset = fields.Boolean(string="Crear activo")
    if _ACCOUNT_ASSET_AVAILABLE:
        asset_id = fields.Many2one("account.asset", string="Activo Odoo", ondelete="set null")
    state = fields.Selection(
        [
            ("mapped_with_asset", "Listo para crear activo"),
            ("fully_depreciated", "Totalmente amortizado"),
            ("review_required", "Revisión requerida"),
            ("error", "Error contable"),
        ],
        string="Estado",
        required=True,
        default="review_required",
    )
    source_move_number = fields.Char(string="Asiento origen")
    source_reference = fields.Char(string="Referencia origen")
    source_label = fields.Char(string="Descripción original")
    source_file = fields.Char(string="Fichero origen")
    warning_message = fields.Text(string="Advertencias")
    company_currency_id = fields.Many2one(related="company_id.currency_id", readonly=True)
