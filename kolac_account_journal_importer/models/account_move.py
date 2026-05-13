import importlib.util

from odoo import models

_ACCOUNT_ASSET_AVAILABLE = importlib.util.find_spec("odoo.addons.account_asset") is not None


class AccountMove(models.Model):
    _inherit = "account.move"

    def action_post(self):
        result = super().action_post()
        if not _ACCOUNT_ASSET_AVAILABLE:
            return result
        trace_model = self.env["kolac.account.import.trace"]
        assets = trace_model.search([
            ("move_id", "in", self.ids),
            ("asset_id", "!=", False),
            ("mapping_type", "=", "asset"),
        ]).mapped("asset_id")
        if assets:
            assets._link_kolac_original_move_lines()
        return result
