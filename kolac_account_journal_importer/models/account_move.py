from odoo import models


class AccountMove(models.Model):
    _inherit = "account.move"

    def action_post(self):
        result = super().action_post()
        trace_model = self.env["kolac.account.import.trace"]
        assets = trace_model.search([
            ("move_id", "in", self.ids),
            ("asset_id", "!=", False),
            ("mapping_type", "=", "asset"),
        ]).mapped("asset_id")
        if assets:
            assets._link_kolac_original_move_lines()
        return result
