import importlib.util

# account_asset is an Odoo Enterprise module; only load the extension if available
if importlib.util.find_spec("odoo.addons.account_asset") is not None:
    from . import account_asset  # noqa: F401

from . import account_move
from . import account_import_batch
from . import account_subaccount_map
