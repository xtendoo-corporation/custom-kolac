from collections import defaultdict

from odoo import Command, fields, models


class AccountAsset(models.Model):
    _inherit = "account.asset"

    kolac_old_asset_account_code = fields.Char(string="KOLAC cuenta antigua activo", readonly=True, copy=False)
    kolac_old_asset_account_name = fields.Char(string="KOLAC nombre antiguo activo", readonly=True, copy=False)
    kolac_old_depreciation_account_code = fields.Char(string="KOLAC cuenta antigua amortización", readonly=True, copy=False)
    kolac_old_depreciation_account_name = fields.Char(string="KOLAC nombre antiguo amortización", readonly=True, copy=False)
    kolac_original_value_detected = fields.Monetary(string="KOLAC valor detectado", currency_field="currency_id", readonly=True, copy=False)
    kolac_accumulated_depreciation_detected = fields.Monetary(string="KOLAC amortización detectada", currency_field="currency_id", readonly=True, copy=False)
    kolac_net_book_value_detected = fields.Monetary(string="KOLAC valor neto detectado", currency_field="currency_id", readonly=True, copy=False)
    kolac_acquisition_date_detected = fields.Date(string="KOLAC fecha detectada", readonly=True, copy=False)
    kolac_import_batch_id = fields.Many2one("kolac.account.import.batch", string="KOLAC lote importación", readonly=True, copy=False)
    kolac_source_file = fields.Char(string="KOLAC fichero origen", readonly=True, copy=False)
    kolac_source_move_number = fields.Char(string="KOLAC asiento origen", readonly=True, copy=False)
    kolac_source_move_ref = fields.Char(string="KOLAC referencia origen", readonly=True, copy=False)
    kolac_source_label = fields.Char(string="KOLAC descripción original", readonly=True, copy=False)
    kolac_asset_review_state = fields.Selection(
        [
            ("mapped_with_asset", "Listo para crear activo"),
            ("fully_depreciated", "Totalmente amortizado"),
            ("review_required", "Revisión requerida"),
            ("error", "Error contable"),
        ],
        string="KOLAC estado revisión",
        readonly=True,
        copy=False,
    )
    kolac_warning_message = fields.Text(string="KOLAC advertencias", readonly=True, copy=False)

    def _link_kolac_original_move_lines(self):
        trace_model = self.env["kolac.account.import.trace"]
        posted_assets = self.filtered(lambda asset: asset.state in ("draft", "model"))
        if not posted_assets:
            return

        traces = trace_model.search([
            ("asset_id", "in", posted_assets.ids),
            ("mapping_type", "=", "asset"),
            ("move_line_id", "!=", False),
            ("move_id.state", "=", "posted"),
        ])
        if not traces:
            return

        lines_by_asset = defaultdict(lambda: self.env["account.move.line"])
        for trace in traces:
            move_line = trace.move_line_id.exists()
            if not move_line:
                continue
            asset = trace.asset_id
            if asset.account_asset_id and move_line.account_id != asset.account_asset_id:
                continue
            lines_by_asset[asset.id] |= move_line

        for asset in posted_assets:
            lines_to_link = lines_by_asset.get(asset.id, self.env["account.move.line"]) - asset.original_move_line_ids
            if lines_to_link:
                asset.write({
                    "original_move_line_ids": [Command.link(line.id) for line in lines_to_link],
                })

    def validate(self):
        self._link_kolac_original_move_lines()
        return super().validate()
