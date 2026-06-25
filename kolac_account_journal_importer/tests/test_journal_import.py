import base64
import importlib.util
import io
import unittest
from datetime import date
from pathlib import Path

from odoo import Command
from odoo.exceptions import AccessError, UserError
from odoo.tests import tagged
from odoo.tests.common import TransactionCase

try:
    import openpyxl
    HAS_OPENPYXL = True
except ImportError:
    HAS_OPENPYXL = False

_ACCOUNT_ASSET_AVAILABLE = importlib.util.find_spec("odoo.addons.account_asset") is not None


DIARY_HEADERS = [
    "Asiento",
    "Fecha",
    "Subcuenta",
    "Concepto",
    "Importe Debe",
    "Importe Haber",
    "Descripción",
    "Referencia",
]
REFERENCE_HEADERS = ["Subcuenta", "Nombre", "NIF", "Tipo"]
OFFICIAL_DIARY_HEADERS = [
    "Fecha",
    "Asto.",
    "Ord",
    "Dia",
    "Cuenta",
    "Ttulo",
    "Concepto",
    "Debe",
    "Haber",
]


def _make_xlsx(headers, data_rows, sheet_name="Datos"):
    workbook = openpyxl.Workbook()
    sheet = workbook.active
    sheet.title = sheet_name
    sheet.append(headers)
    for row in data_rows:
        sheet.append(row)
    buffer = io.BytesIO()
    workbook.save(buffer)
    return base64.b64encode(buffer.getvalue())


def _make_official_diary_xlsx(data_rows, year=2024, include_year_line=True):
    workbook = openpyxl.Workbook()
    sheet = workbook.active
    sheet.title = "Diario de movimientos oficial"
    if include_year_line:
        sheet.append(["ISJALU 2015 SOCIEDAD LIMITADA"])
    else:
        sheet.append(["EMPRESA DE EJEMPLO SOCIEDAD LIMITADA"])
    sheet.append(["Usuario: Carlos"])
    if include_year_line:
        sheet.append(["Fecha: 08/05/2026 - 12:40:56"])
    else:
        sheet.append(["Fecha de emisión: 08/05 - 12:40:56"])
    sheet.append(["Diario de movimientos oficial"])
    if include_year_line:
        sheet.append([f"Movimientos desde el da 01/01/{year} hasta el 31/12/{year} (Euros)"])
    else:
        sheet.append(["Movimientos del ejercicio (Euros)"])
    sheet.append([])
    sheet.append(OFFICIAL_DIARY_HEADERS)
    for row in data_rows:
        sheet.append(row)
    buffer = io.BytesIO()
    workbook.save(buffer)
    return base64.b64encode(buffer.getvalue())


@tagged("post_install", "-at_install", "kolac_account_journal_import")
class TestJournalImport(TransactionCase):

    @classmethod
    def _get_or_create_account(cls, code, name, account_type, **extra_vals):
        account = cls.env["account.account"].search([
            ("code", "=", code),
            ("company_ids", "in", cls.env.company.id),
        ], limit=1)
        if account:
            return account

        vals = {
            "name": name,
            "code": code,
            "account_type": account_type,
        }
        vals.update(extra_vals)
        return cls.env["account.account"].create(vals)

    def setUp(self):
        super().setUp()
        if not HAS_OPENPYXL:
            self.skipTest("openpyxl no está disponible")

        self.company = self.env.company
        self.journal = self.env["account.journal"].search([
            ("type", "=", "general"),
            ("company_id", "=", self.company.id),
        ], limit=1)
        if not self.journal:
            self.journal = self.env["account.journal"].create({
                "name": "Diario Test",
                "code": "DTST",
                "type": "general",
                "company_id": self.company.id,
            })

        self.env["kolac.account.subaccount.map"].search([
            ("company_id", "=", self.company.id),
            ("source_code", "in", [
                "430000001",
                "400000123",
                "410000055",
                "216000079",
                "217001087",
                "281600079",
                "281701087",
                "410002227",
                "629010014",
                "629050002",
                "629040000",
                "412000000",
            ]),
        ]).unlink()

        self.account_121 = self._get_or_create_account("121000", "Fondos propios", "equity_unaffected")
        self.account_430 = self._get_or_create_account("430000", "Clientes", "asset_receivable", reconcile=True)
        self.account_400 = self._get_or_create_account("400000", "Proveedores", "liability_payable", reconcile=True)
        self.account_410 = self._get_or_create_account("410000", "Acreedores", "liability_payable", reconcile=True)
        self.account_216 = self._get_or_create_account("216000", "Mobiliario", "asset_fixed")
        self.account_211 = self._get_or_create_account("211000", "Construcciones", "asset_fixed")
        self.account_100 = self._get_or_create_account("100000", "Capital social", "equity")
        self.account_217 = self._get_or_create_account("217000", "Equipos para procesos de información", "asset_fixed")
        self.account_218 = self._get_or_create_account("218000", "Elementos de transporte", "asset_fixed")
        self.account_219 = self._get_or_create_account("219000", "Otro inmovilizado material", "asset_fixed")
        self.account_281100 = self._get_or_create_account("281100", "Amortización construcciones", "asset_non_current")
        self.account_281600 = self._get_or_create_account("281600", "Amortización mobiliario", "asset_non_current")
        self.account_281700 = self._get_or_create_account("281700", "Amortización equipos informáticos", "asset_non_current")
        self.account_281800 = self._get_or_create_account("281800", "Amortización elementos de transporte", "asset_non_current")
        self.account_281900 = self._get_or_create_account("281900", "Amortización otro inmovilizado material", "asset_non_current")
        self.account_681700 = self._get_or_create_account("681700", "Dotación amortización equipos informáticos", "expense_depreciation")
        self.account_623100 = self._get_or_create_account("623100", "Contabilidad", "expense")
        self.account_600 = self._get_or_create_account("600000", "Compras de mercaderías", "expense_direct_cost")
        self.account_700 = self._get_or_create_account("700000", "Ventas", "income")
        self.account_752 = self._get_or_create_account("752000", "Ingresos por arrendamientos", "income_other")
        self.account_472 = self._get_or_create_account("472000", "IVA soportado", "asset_current")
        self.account_477 = self._get_or_create_account("477000", "IVA repercutido", "liability_current")
        self._ensure_tax("purchase")
        self._ensure_tax("sale")

    def _ensure_tax(self, tax_use):
        tax = self.env["account.tax"].search([
            ("company_id", "=", self.company.id),
            ("type_tax_use", "=", tax_use),
            ("amount", "=", 21.0),
        ], limit=1)
        if tax:
            return tax
        return self.env["account.tax"].create({
            "name": f"IVA 21% {'Compras' if tax_use == 'purchase' else 'Ventas'}",
            "type_tax_use": tax_use,
            "amount_type": "percent",
            "amount": 21.0,
            "company_id": self.company.id,
            "invoice_repartition_line_ids": [
                Command.create({"repartition_type": "base", "factor_percent": 100}),
                Command.create({"repartition_type": "tax", "factor_percent": 100}),
            ],
            "refund_repartition_line_ids": [
                Command.create({"repartition_type": "base", "factor_percent": 100}),
                Command.create({"repartition_type": "tax", "factor_percent": 100}),
            ],
        })

    def _make_diary_file(self, rows):
        return _make_xlsx(DIARY_HEADERS, rows, sheet_name="Libro Diario")

    def _make_official_diary_file(self, rows, year=2024, include_year_line=True):
        return _make_official_diary_xlsx(rows, year=year, include_year_line=include_year_line)

    def _make_account_list_file(self, rows):
        return _make_xlsx(REFERENCE_HEADERS, rows, sheet_name="Listado")

    def _make_account_plan_file(self, rows):
        return _make_xlsx(REFERENCE_HEADERS, rows, sheet_name="Plan")

    def _create_wizard(self, diary_rows, account_list_rows=None, account_plan_rows=None, **extra_vals):
        vals = {
            "diary_file": self._make_diary_file(diary_rows),
            "diary_file_name": "Libro Diario.xlsx",
            "company_id": self.company.id,
            "journal_id": self.journal.id,
            "create_missing_accounts": False,
            "post_moves": False,
        }
        if account_list_rows is not None:
            vals.update({
                "account_list_file": self._make_account_list_file(account_list_rows),
                "account_list_file_name": "Listado de cuentas.xlsx",
            })
        if account_plan_rows is not None:
            vals.update({
                "account_plan_file": self._make_account_plan_file(account_plan_rows),
                "account_plan_file_name": "Plan de cuentas.xlsx",
            })
        vals.update(extra_vals)
        return self.env["kolac.account.journal.import.wizard"].create(vals)

    def _create_official_wizard(self, diary_rows, year=2024, include_year_line=True, **extra_vals):
        vals = {
            "diary_file": self._make_official_diary_file(
                diary_rows, year=year, include_year_line=include_year_line
            ),
            "diary_file_name": f"ISJ {year} Diario de movimientos oficial.XLSX",
            "company_id": self.company.id,
            "journal_id": self.journal.id,
            "create_missing_accounts": False,
            "post_moves": False,
        }
        vals.update(extra_vals)
        return self.env["kolac.account.journal.import.wizard"].create(vals)

    def _default_account_list_rows(self):
        return [
            ["430000001", "Cliente A", "", "cliente"],
            ["400000123", "Proveedor X", "", "proveedor"],
            ["410000055", "Acreedor Y", "", "acreedor"],
            ["216000079", "Mesa oficina principal", "", "activo"],
            ["281600079", "Amortización Mesa oficina principal", "", "amortizacion"],
            ["217001087", "NAP TS-INF. SERRANO Y MAS (P24-22)", "", "activo"],
            ["281701087", "Amort. NAP TS-INF. SERRANO Y MAS (P24-22)", "", "amortizacion"],
            ["472000921", "IVA soportado 21%", "", "impuesto"],
        ]

    def _default_account_plan_rows(self):
        return [
            ["430000001", "Clientes detalle", "", "cliente"],
            ["400000123", "Proveedores detalle", "", "proveedor"],
            ["410000055", "Acreedores detalle", "", "acreedor"],
            ["216000079", "Mobiliario", "", "activo"],
            ["281600079", "Amortización acumulada mobiliario", "", "amortizacion"],
            ["217001087", "Equipos proceso información", "", "activo"],
            ["281701087", "Amortización acumulada equipos proceso información", "", "amortizacion"],
            ["700000123", "Ventas detalle", "", "venta"],
            ["472000921", "IVA soportado 21%", "", "impuesto"],
        ]

    def _create_account_user(self):
        return self.env["res.users"].with_context(no_reset_password=True).create({
            "name": "Usuario Contabilidad Importador",
            "login": "importador_contable",
            "email": "importador_contable@example.com",
            "company_id": self.company.id,
            "company_ids": [(6, 0, [self.company.id])],
            "group_ids": [(6, 0, [
                self.env.ref("base.group_user").id,
                self.env.ref("account.group_account_user").id,
            ])],
        })

    def _get_tax(self, tax_use):
        tax = self._ensure_tax(tax_use)
        self.assertTrue(tax)
        return tax

    def test_01_imports_without_reference_excels(self):
        wizard = self._create_wizard([
            [1, date(2026, 1, 1), "430000001", "APERTURA", "121,00", "", "Cliente A", "FAC-1"],
            [1, date(2026, 1, 1), "121000000", "APERTURA", "", "121,00", "Contrapartida", "FAC-1"],
        ], create_missing_accounts=True)

        wizard.action_analyze()

        self.assertFalse(wizard.log_line_ids.filtered(lambda line: line.category == "sources"))

        action = wizard.action_confirm_import()
        move = self.env["account.move"].browse(action["res_id"])
        receivable_line = move.line_ids.filtered(lambda line: line.account_id.code == "430000001")

        self.assertTrue(receivable_line)
        self.assertEqual(receivable_line.partner_id.name, "Cliente A")

    def test_01b_official_diary_allows_confirm_with_single_excel(self):
        wizard = self._create_official_wizard([
            ["15-Ene.", 2, 1, 15, "430000007", "SOLORA PROYECT SL", "SOLORA PROYECT SL N. FRA: 005/2024", 121.0, None],
            ["15-Ene.", 2, 2, 15, "752000001", "INGRESOS POR ARRENDAMIENTOS ARJONA", "SOLORA PROYECT SL N. FRA: 005/2024", None, 100.0],
            ["15-Ene.", 2, 3, 15, "477000000", "HACIENDA PBLICA, IVA REPERCUTIDO", "SOLORA PROYECT SL N. FRA: 005/2024", None, 21.0],
        ], create_missing_accounts=True)

        wizard.action_analyze()

        self.assertFalse(wizard.log_line_ids.filtered(lambda line: line.category == "sources"))
        self.assertEqual(wizard.detected_move_count, 1)
        self.assertEqual(wizard.move_preview_line_ids.move_date, date(2024, 1, 15))

        action = wizard.action_confirm_import()
        move = self.env["account.move"].browse(action["res_id"])
        receivable_line = move.line_ids.filtered(lambda line: line.account_id.code == "430000007")
        income_line = move.line_ids.filtered(lambda line: line.account_id.code == "752000001")

        self.assertTrue(receivable_line)
        self.assertEqual(receivable_line.partner_id.name, "SOLORA PROYECT SL")
        self.assertTrue(income_line)
        self.assertEqual(move.date, date(2024, 1, 15))

    def test_01c_official_diary_detects_building_asset_and_depreciation(self):
        wizard = self._create_official_wizard([
            ["01-Ene.", 3, 1, 1, "211000001", "OFICINA ARJONA", "ASIENTO DE APERTURA", 100000.0, None],
            ["01-Ene.", 3, 2, 1, "281100001", "A/A OFICINA ARJONA", "ASIENTO DE APERTURA", None, 20000.0],
            ["01-Ene.", 3, 3, 1, "100000004", "CS- KOLAC CAPITAL", "ASIENTO DE APERTURA", None, 80000.0],
        ], create_missing_accounts=True)

        wizard.action_analyze()

        asset_mapping = wizard.mapping_line_ids.filtered(lambda line: line.source_code == "211000001")
        depreciation_mapping = wizard.mapping_line_ids.filtered(lambda line: line.source_code == "281100001")
        asset_preview = wizard.asset_line_ids.filtered(lambda line: line.source_code == "211000001")

        self.assertEqual(asset_mapping.mapping_type, "asset")
        self.assertEqual(asset_mapping.target_account_code, "211000001")
        self.assertEqual(depreciation_mapping.mapping_type, "accumulated_depreciation")
        self.assertEqual(depreciation_mapping.target_account_code, "281100001")
        self.assertEqual(asset_preview.target_account_id.code, "211000001")
        self.assertEqual(asset_preview.depreciation_target_account_id.code, "281100001")
        self.assertEqual(asset_preview.original_value, 100000.0)
        self.assertEqual(asset_preview.accumulated_value, 20000.0)
        self.assertEqual(asset_preview.residual_value, 80000.0)

    def test_01d_can_analyze_real_kolac_workbook(self):
        workbook_path = Path(__file__).resolve().parents[2] / "ISJ 2024 Diario de movimientos oficial.XLSX"
        if not workbook_path.exists():
            self.skipTest("No se ha encontrado el Excel real de KOLAC en el repositorio")

        wizard = self.env["kolac.account.journal.import.wizard"].create({
            "diary_file": base64.b64encode(workbook_path.read_bytes()),
            "diary_file_name": workbook_path.name,
            "company_id": self.company.id,
            "journal_id": self.journal.id,
            "create_missing_accounts": False,
            "post_moves": False,
        })

        wizard.action_analyze()

        self.assertGreaterEqual(wizard.detected_move_count, 300)
        self.assertGreaterEqual(wizard.detected_line_count, 1500)
        self.assertGreaterEqual(wizard.detected_account_count, 100)
        self.assertFalse(wizard.log_line_ids.filtered(lambda line: line.category == "sources"))

    def test_01e_report_year_resolves_dates_when_file_has_no_year(self):
        wizard = self._create_official_wizard(
            [
                ["15-Ene.", 2, 1, 15, "430000007", "SOLORA PROYECT SL", "FRA 005", 121.0, None],
                ["15-Ene.", 2, 2, 15, "752000001", "ARRENDAMIENTOS", "FRA 005", None, 121.0],
            ],
            include_year_line=False,
            diary_file_name="Diario de movimientos oficial.XLSX",
            report_year=2024,
            create_missing_accounts=True,
        )

        wizard.action_analyze()

        self.assertEqual(wizard.detected_move_count, 1)
        self.assertEqual(wizard.move_preview_line_ids.move_date, date(2024, 1, 15))

    def test_01f_report_year_overrides_year_guessed_from_file(self):
        wizard = self._create_official_wizard(
            [
                ["15-Ene.", 2, 1, 15, "430000007", "SOLORA PROYECT SL", "FRA 005", 121.0, None],
                ["15-Ene.", 2, 2, 15, "752000001", "ARRENDAMIENTOS", "FRA 005", None, 121.0],
            ],
            year=2024,
            report_year=2025,
            create_missing_accounts=True,
        )

        wizard.action_analyze()

        self.assertEqual(wizard.move_preview_line_ids.move_date, date(2025, 1, 15))

    def test_01g_report_year_resolves_numeric_day_month_dates(self):
        wizard = self._create_official_wizard(
            [
                ["15/01", 2, 1, 15, "430000007", "SOLORA PROYECT SL", "FRA 005", 121.0, None],
                ["15/01", 2, 2, 15, "752000001", "ARRENDAMIENTOS", "FRA 005", None, 121.0],
            ],
            include_year_line=False,
            diary_file_name="Diario de movimientos oficial.XLSX",
            report_year=2024,
            create_missing_accounts=True,
        )

        wizard.action_analyze()

        self.assertEqual(wizard.move_preview_line_ids.move_date, date(2024, 1, 15))

    def test_01h_missing_year_raises_actionable_error(self):
        wizard = self._create_official_wizard(
            [
                ["15-Ene.", 2, 1, 15, "430000007", "SOLORA PROYECT SL", "FRA 005", 121.0, None],
                ["15-Ene.", 2, 2, 15, "752000001", "ARRENDAMIENTOS", "FRA 005", None, 121.0],
            ],
            include_year_line=False,
            diary_file_name="Diario de movimientos oficial.XLSX",
            create_missing_accounts=True,
        )

        with self.assertRaises(UserError) as error:
            wizard.action_analyze()

        self.assertIn("Año del ejercicio", error.exception.args[0])

    def test_01i_invalid_report_year_is_rejected(self):
        wizard = self._create_official_wizard(
            [
                ["15-Ene.", 2, 1, 15, "430000007", "SOLORA PROYECT SL", "FRA 005", 121.0, None],
                ["15-Ene.", 2, 2, 15, "752000001", "ARRENDAMIENTOS", "FRA 005", None, 121.0],
            ],
            report_year=12,
            create_missing_accounts=True,
        )

        with self.assertRaises(UserError):
            wizard.action_analyze()

    def test_02_maps_customer_account_to_partner_and_preserves_subaccount(self):
        wizard = self._create_wizard([
            [2, date(2026, 1, 2), "430000001", "FACTURA", "121,00", "", "Cliente A", "FAC-2"],
            [2, date(2026, 1, 2), "700000123", "FACTURA", "", "100,00", "Venta", "FAC-2"],
            [2, date(2026, 1, 2), "472000921", "FACTURA", "", "21,00", "IVA soportado 21%", "FAC-2"],
        ], self._default_account_list_rows(), self._default_account_plan_rows(), create_missing_accounts=True)

        wizard.action_analyze()
        mapping = wizard.mapping_line_ids.filtered(lambda line: line.source_code == "430000001")
        self.assertEqual(mapping.mapping_type, "partner_customer")
        self.assertEqual(mapping.target_account_code, "430000001")

        action = wizard.action_confirm_import()
        move = self.env["account.move"].browse(action["res_id"])
        receivable_line = move.line_ids.filtered(lambda line: line.account_id.code == "430000001")
        self.assertTrue(receivable_line.partner_id)
        self.assertEqual(receivable_line.partner_id.name, "Cliente A")

    def test_03_maps_supplier_and_creditor_to_partners(self):
        wizard = self._create_wizard([
            [3, date(2026, 1, 3), "400000123", "COMPRA", "", "50,00", "Proveedor X", "FAC-P-1"],
            [3, date(2026, 1, 3), "410000055", "COMPRA", "", "10,00", "Acreedor Y", "FAC-P-1"],
            [3, date(2026, 1, 3), "121000000", "COMPRA", "60,00", "", "Tesorería", "FAC-P-1"],
        ], self._default_account_list_rows(), self._default_account_plan_rows(), create_missing_accounts=True)

        wizard.action_analyze()
        supplier_mapping = wizard.mapping_line_ids.filtered(lambda line: line.source_code == "400000123")
        creditor_mapping = wizard.mapping_line_ids.filtered(lambda line: line.source_code == "410000055")
        self.assertEqual(supplier_mapping.target_account_code, "400000123")
        self.assertEqual(creditor_mapping.target_account_code, "410000055")

        action = wizard.action_confirm_import()
        move = self.env["account.move"].browse(action["res_id"])
        self.assertEqual(move.line_ids.filtered(lambda line: line.account_id.code == "400000123").partner_id.name, "Proveedor X")
        self.assertEqual(move.line_ids.filtered(lambda line: line.account_id.code == "410000055").partner_id.name, "Acreedor Y")

    def test_04_detects_asset_and_accumulated_depreciation(self):
        wizard = self._create_wizard([
            [4, date(2026, 1, 4), "216000079", "APERTURA", "1500,00", "", "Mesa oficina principal", "ACT-1"],
            [4, date(2026, 1, 4), "281600079", "APERTURA", "", "900,00", "Amortización acumulada", "ACT-1"],
            [4, date(2026, 1, 4), "121000000", "APERTURA", "", "600,00", "Contrapartida", "ACT-1"],
        ], self._default_account_list_rows(), self._default_account_plan_rows(), create_missing_accounts=True)

        wizard.action_analyze()
        asset_mapping = wizard.mapping_line_ids.filtered(lambda line: line.source_code == "216000079")
        depreciation_mapping = wizard.mapping_line_ids.filtered(lambda line: line.source_code == "281600079")
        asset_preview = wizard.asset_line_ids.filtered(lambda line: line.source_code == "216000079")

        self.assertEqual(asset_mapping.mapping_type, "asset")
        self.assertEqual(asset_mapping.target_account_code, "216000079")
        self.assertEqual(depreciation_mapping.mapping_type, "accumulated_depreciation")
        self.assertEqual(depreciation_mapping.target_account_code, "281600079")
        self.assertEqual(asset_preview.original_value, 1500.0)
        self.assertEqual(asset_preview.accumulated_value, 900.0)
        self.assertEqual(asset_preview.residual_value, 600.0)

    def test_05_creates_exact_subaccounts_when_missing(self):
        wizard = self._create_wizard([
            [5, date(2026, 1, 5), "700000123", "VENTA", "", "100,00", "Venta", "FAC-5"],
            [5, date(2026, 1, 5), "121000000", "VENTA", "100,00", "", "Cobro", "FAC-5"],
        ], self._default_account_list_rows(), self._default_account_plan_rows(), create_missing_accounts=True)

        wizard.action_analyze()
        sales_mapping = wizard.mapping_line_ids.filtered(lambda line: line.source_code == "700000123")
        self.assertEqual(sales_mapping.target_account_code, "700000123")

        wizard.action_confirm_import()
        created_account = self.env["account.account"].search([("code", "=", "700000123")], limit=1)
        self.assertTrue(created_account)

    def test_06_blocks_unbalanced_entries(self):
        wizard = self._create_wizard([
            [6, date(2026, 1, 6), "430000001", "ERROR", "121,00", "", "Cliente A", "ERR-6"],
            [6, date(2026, 1, 6), "700000123", "ERROR", "", "100,00", "Venta", "ERR-6"],
        ], self._default_account_list_rows(), self._default_account_plan_rows(), create_missing_accounts=True)

        wizard.action_analyze()
        self.assertTrue(wizard.log_line_ids.filtered(lambda line: line.blocking and line.category == "entry"))
        with self.assertRaises(UserError):
            wizard.action_confirm_import()

    def test_07_allows_reimporting_previous_imports(self):
        diary_rows = [
            [7, date(2026, 1, 7), "430000001", "FACTURA", "121,00", "", "Cliente A", "FAC-7"],
            [7, date(2026, 1, 7), "700000123", "FACTURA", "", "100,00", "Venta", "FAC-7"],
            [7, date(2026, 1, 7), "472000921", "FACTURA", "", "21,00", "IVA soportado 21%", "FAC-7"],
        ]
        wizard = self._create_wizard(diary_rows, self._default_account_list_rows(), self._default_account_plan_rows(), create_missing_accounts=True)
        wizard.action_analyze()
        wizard.action_confirm_import()

        batch_model = self.env["kolac.account.import.batch"]
        first_done = batch_model.search_count([
            ("company_id", "=", self.company.id),
            ("state", "=", "done"),
        ])

        wizard_2 = self._create_wizard(diary_rows, self._default_account_list_rows(), self._default_account_plan_rows(), create_missing_accounts=True)
        wizard_2.action_analyze()
        result = wizard_2.action_confirm_import()

        self.assertEqual(result.get("res_model"), "account.move")
        second_done = batch_model.search_count([
            ("company_id", "=", self.company.id),
            ("state", "=", "done"),
        ])
        self.assertEqual(second_done, first_done + 1)

    def test_07b_reusing_wizard_keeps_previous_imports(self):
        first_diary_rows = [
            [701, date(2025, 1, 7), "430000001", "FACTURA", "121,00", "", "Cliente A", "FAC-701"],
            [701, date(2025, 1, 7), "700000123", "FACTURA", "", "100,00", "Venta", "FAC-701"],
            [701, date(2025, 1, 7), "472000921", "FACTURA", "", "21,00", "IVA soportado 21%", "FAC-701"],
        ]
        wizard = self._create_wizard(
            first_diary_rows,
            self._default_account_list_rows(),
            self._default_account_plan_rows(),
            create_missing_accounts=True,
        )
        wizard.action_analyze()
        wizard.action_confirm_import()
        first_batch = wizard.preview_batch_id
        first_moves = first_batch.trace_line_ids.move_id
        first_move_count = first_batch.move_count
        first_line_count = first_batch.line_count
        first_imported_at = first_batch.imported_at

        second_diary_rows = [
            [702, date(2026, 1, 8), "430000001", "FACTURA", "242,00", "", "Cliente A", "FAC-702"],
            [702, date(2026, 1, 8), "700000123", "FACTURA", "", "200,00", "Venta", "FAC-702"],
            [702, date(2026, 1, 8), "472000921", "FACTURA", "", "42,00", "IVA soportado 21%", "FAC-702"],
        ]
        wizard.write({
            "diary_file": self._make_diary_file(second_diary_rows),
            "diary_file_name": "Libro Diario 2026.xlsx",
        })
        wizard.action_analyze()

        self.assertNotEqual(wizard.preview_batch_id, first_batch)
        wizard.action_confirm_import()

        self.assertTrue(first_batch.exists())
        self.assertEqual(first_batch.state, "done")
        self.assertFalse(first_batch.dry_run)
        self.assertEqual(first_batch.move_count, first_move_count)
        self.assertEqual(first_batch.line_count, first_line_count)
        self.assertEqual(first_batch.imported_at, first_imported_at)
        self.assertEqual(first_moves.exists(), first_moves)

    def test_08_generates_tax_mapping_warning_but_preserves_tax_subaccount(self):
        wizard = self._create_wizard([
            [8, date(2026, 1, 8), "472000921", "FACTURA", "21,00", "", "IVA soportado 21%", "FAC-8"],
            [8, date(2026, 1, 8), "121000000", "FACTURA", "", "21,00", "Contrapartida", "FAC-8"],
        ], self._default_account_list_rows(), self._default_account_plan_rows(), create_missing_accounts=True)

        wizard.action_analyze()
        mapping = wizard.mapping_line_ids.filtered(lambda line: line.source_code == "472000921")
        self.assertEqual(mapping.target_account_code, "472000921")
        self.assertTrue(mapping.warning_message or mapping.target_tax_id)

    def test_09_persists_full_review_report_in_batch(self):
        wizard = self._create_wizard([
            [9, date(2026, 1, 9), "430000001", "FACTURA", "121,00", "", "Cliente A", "FAC-9"],
            [9, date(2026, 1, 9), "700000123", "FACTURA", "", "100,00", "Venta", "FAC-9"],
            [9, date(2026, 1, 9), "472000921", "FACTURA", "", "21,00", "IVA soportado 21%", "FAC-9"],
        ], self._default_account_list_rows(), self._default_account_plan_rows(), create_missing_accounts=True)

        wizard.action_analyze()
        wizard.action_confirm_import()

        batch = self.env["kolac.account.import.batch"].search([
            ("company_id", "=", self.company.id),
            ("diary_file_name", "=", "Libro Diario.xlsx"),
        ], order="id desc", limit=1)

        self.assertTrue(batch)
        self.assertIn("MAPEOS", batch.review_report)
        self.assertIn("ASIENTOS", batch.review_report)

    def test_09b_integrates_purchase_tax_into_move_lines(self):
        wizard = self._create_wizard([
            [91, date(2026, 1, 9), "600000000", "FACTURA", "100,00", "", "Licencias Odoo 19.0", "FAC-COMPRA-91"],
            [91, date(2026, 1, 9), "472000921", "FACTURA", "21,00", "", "IVA soportado 21%", "FAC-COMPRA-91"],
            [91, date(2026, 1, 9), "410000055", "FACTURA", "", "121,00", "Proveedor X", "FAC-COMPRA-91"],
        ], self._default_account_list_rows(), self._default_account_plan_rows(), create_missing_accounts=True)

        wizard.action_analyze()
        wizard.action_confirm_import()

        move = wizard.preview_batch_id.trace_line_ids.move_id[:1]
        expense_line = move.line_ids.filtered(lambda line: line.account_id.code == "600000000")
        tax_line = move.line_ids.filtered(lambda line: line.account_id.code == "472000921")

        self.assertTrue(expense_line.tax_ids)
        self.assertTrue(tax_line.tax_line_id)
        if expense_line.tax_ids.invoice_repartition_line_ids.tag_ids:
            self.assertTrue(expense_line.tax_tag_ids)
        if tax_line.tax_line_id.invoice_repartition_line_ids.tag_ids:
            self.assertTrue(tax_line.tax_tag_ids)

    def test_09c_integrates_sale_tax_into_move_lines(self):
        wizard = self._create_wizard([
            [92, date(2026, 1, 9), "430000001", "FACTURA", "121,00", "", "Cliente A", "FAC-VENTA-92"],
            [92, date(2026, 1, 9), "700000123", "FACTURA", "", "100,00", "Licencias Odoo 19.0", "FAC-VENTA-92"],
            [92, date(2026, 1, 9), "477000021", "FACTURA", "", "21,00", "IVA repercutido 21%", "FAC-VENTA-92"],
        ], self._default_account_list_rows(), self._default_account_plan_rows(), create_missing_accounts=True)

        wizard.action_analyze()
        wizard.action_confirm_import()

        move = wizard.preview_batch_id.trace_line_ids.move_id[:1]
        income_line = move.line_ids.filtered(lambda line: line.account_id.code == "700000123")
        tax_line = move.line_ids.filtered(lambda line: line.account_id.code == "477000021")

        self.assertTrue(income_line.tax_ids)
        self.assertTrue(tax_line.tax_line_id)
        if income_line.tax_ids.invoice_repartition_line_ids.tag_ids:
            self.assertTrue(income_line.tax_tag_ids)
        if tax_line.tax_line_id.invoice_repartition_line_ids.tag_ids:
            self.assertTrue(tax_line.tax_tag_ids)

    def test_09d_opening_asset_entry_removes_tax_ids_and_tax_tags(self):
        opening_journal = self.env["account.journal"].create({
            "name": "Apertura 2026",
            "code": "APER26",
            "type": "general",
            "company_id": self.company.id,
        })
        wizard = self._create_wizard([
            [93, date(2026, 1, 1), "217001079", "ASIENTO DE APERTURA", "1000,00", "", "IPAD PRO 11 - cuenta antigua 217001079", "AP-93"],
            [93, date(2026, 1, 1), "472000921", "ASIENTO DE APERTURA", "210,00", "", "IVA soportado 21%", "AP-93"],
            [93, date(2026, 1, 1), "121000000", "ASIENTO DE APERTURA", "", "1210,00", "Contrapartida", "AP-93"],
        ], self._default_account_list_rows() + [
            ["217001079", "IPAD PRO 11", "", "activo"],
        ], self._default_account_plan_rows() + [
            ["217001079", "Equipos proceso información", "", "activo"],
        ], create_missing_accounts=True, journal_id=opening_journal.id)

        wizard.action_analyze()
        wizard.action_confirm_import()

        move = wizard.preview_batch_id.trace_line_ids.move_id[:1]
        asset_line = move.line_ids.filtered(lambda line: line.account_id.code == "217001079")

        self.assertTrue(asset_line)
        self.assertFalse(asset_line.tax_ids)
        self.assertFalse(asset_line.tax_tag_ids)
        self.assertFalse(asset_line.tax_repartition_line_id)
        self.assertFalse(asset_line.tax_line_id)
        self.assertFalse(move.line_ids.filtered(lambda line: line.tax_ids or line.tax_tag_ids or line.tax_repartition_line_id or line.tax_line_id))
        self.assertIn("cuenta antigua 217001079", asset_line.name)
        self.assertIn("VALIDACIONES FISCALES", wizard.preview_batch_id.review_report)
        self.assertIn("Se han eliminado impuestos/etiquetas fiscales de una línea de apertura", wizard.preview_batch_id.review_report)

    def test_09e_opening_entry_ignores_default_taxes_from_asset_account(self):
        purchase_tax = self._get_tax("purchase")
        opening_journal = self.env["account.journal"].create({
            "name": "Saldo inicial 2026",
            "code": "SI2026",
            "type": "general",
            "company_id": self.company.id,
        })
        self.account_217.tax_ids = [Command.set([purchase_tax.id])]

        wizard = self._create_wizard([
            [94, date(2026, 1, 1), "217001080", "SALDO INICIAL", "850,00", "", "Equipo reacondicionado - cuenta antigua 217001080", "AP-94"],
            [94, date(2026, 1, 1), "121000000", "SALDO INICIAL", "", "850,00", "Contrapartida", "AP-94"],
        ], self._default_account_list_rows() + [
            ["217001080", "Equipo reacondicionado", "", "activo"],
        ], self._default_account_plan_rows() + [
            ["217001080", "Equipos proceso información", "", "activo"],
        ], create_missing_accounts=True, journal_id=opening_journal.id)

        wizard.action_analyze()
        wizard.action_confirm_import()

        move = wizard.preview_batch_id.trace_line_ids.move_id[:1]
        asset_line = move.line_ids.filtered(lambda line: line.account_id.code == "217001080")

        self.assertTrue(asset_line)
        self.assertFalse(asset_line.tax_ids)
        self.assertFalse(asset_line.tax_tag_ids)
        self.assertFalse(asset_line.tax_repartition_line_id)

    def test_09f_opening_entry_for_219_group_keeps_asset_without_fiscal_tags(self):
        opening_journal = self.env["account.journal"].create({
            "name": "Migración apertura 2026",
            "code": "MAP26",
            "type": "general",
            "company_id": self.company.id,
        })
        wizard = self._create_wizard([
            [95, date(2026, 1, 1), "219000123", "MIGRACIÓN", "500,00", "", "Otro inmovilizado - cuenta antigua 219000123", "AP-95"],
            [95, date(2026, 1, 1), "472000921", "MIGRACIÓN", "105,00", "", "IVA soportado 21%", "AP-95"],
            [95, date(2026, 1, 1), "121000000", "MIGRACIÓN", "", "605,00", "Contrapartida", "AP-95"],
        ], self._default_account_list_rows() + [
            ["219000123", "Otro inmovilizado", "", "activo"],
        ], self._default_account_plan_rows() + [
            ["219000123", "Otro inmovilizado material", "", "activo"],
        ], create_missing_accounts=True, journal_id=opening_journal.id)

        wizard.action_analyze()
        mapping = wizard.mapping_line_ids.filtered(lambda line: line.source_code == "219000123")
        self.assertEqual(mapping.mapping_type, "asset")
        self.assertEqual(mapping.target_account_code, "219000123")

        wizard.action_confirm_import()

        move = wizard.preview_batch_id.trace_line_ids.move_id[:1]
        asset_line = move.line_ids.filtered(lambda line: line.account_id.code == "219000123")

        self.assertTrue(asset_line)
        self.assertFalse(asset_line.tax_ids)
        self.assertFalse(asset_line.tax_tag_ids)
        self.assertFalse(asset_line.tax_repartition_line_id)
        self.assertFalse(move.line_ids.filtered(lambda line: line.tax_tag_ids))

    def test_10_maps_447_to_partner_and_creates_account_if_missing(self):
        account_447 = self.env["account.account"].search([
            ("code", "=", "447000123"),
            ("company_ids", "in", self.company.id),
        ], limit=1)
        if account_447 and not self.env["account.move.line"].search_count([("account_id", "=", account_447.id)]):
            account_447.unlink()
        account_447 = self.env["account.account"].search([
            ("code", "=", "447000123"),
            ("company_ids", "in", self.company.id),
        ], limit=1)

        wizard = self._create_wizard([
            [10, date(2026, 1, 10), "447000123", "SUBVENCIÓN", "250,00", "", "Organismo concedente", "SUB-10"],
            [10, date(2026, 1, 10), "121000000", "SUBVENCIÓN", "", "250,00", "Contrapartida", "SUB-10"],
        ], [
            ["447000123", "Organismo concedente", "", "cliente"],
        ], [
            ["447000123", "Deudores subvenciones", "", "cliente"],
        ], create_missing_accounts=True)

        wizard.action_analyze()
        mapping = wizard.mapping_line_ids.filtered(lambda line: line.source_code == "447000123")
        self.assertEqual(mapping.mapping_type, "partner_customer")
        self.assertEqual(mapping.target_account_code, "447000123")

        action = wizard.action_confirm_import()
        move = self.env["account.move"].browse(action["res_id"])
        receivable_line = move.line_ids.filtered(lambda line: line.account_id.code == "447000123")

        self.assertTrue(receivable_line)
        self.assertTrue(receivable_line.partner_id)
        self.assertEqual(receivable_line.partner_id.name, "Organismo concedente")

        created_or_reused_account = self.env["account.account"].search([
            ("code", "=", "447000123"),
            ("company_ids", "in", self.company.id),
        ], limit=1)
        self.assertTrue(created_or_reused_account)

    def test_11_maps_monitor_supplier_and_expense_preserving_subaccounts(self):
        wizard = self._create_wizard([
            [11, date(2026, 1, 11), "629010014", "SERVICIO", "100,00", "", "MONITOR INFORMÁTICA", "FAC-MON-11"],
            [11, date(2026, 1, 11), "410002227", "SERVICIO", "", "100,00", "MONITOR INFORMÁTICA", "FAC-MON-11"],
        ], [
            ["410002227", "Monitor Informática", "", "acreedor"],
            ["629010014", "Monitor Informática", "", "gasto"],
        ], [
            ["410002227", "Monitor Informática", "", "acreedor"],
            ["629010014", "Monitor Informática", "", "gasto"],
        ], create_missing_accounts=True)

        wizard.action_analyze()
        expense_mapping = wizard.mapping_line_ids.filtered(lambda line: line.source_code == "629010014")
        creditor_mapping = wizard.mapping_line_ids.filtered(lambda line: line.source_code == "410002227")

        self.assertEqual(expense_mapping.mapping_type, "expense_nature_by_supplier")
        self.assertEqual(expense_mapping.target_account_code, "629010014")
        self.assertEqual(expense_mapping.partner_name, "Monitor Informática")
        self.assertEqual(creditor_mapping.mapping_type, "partner_creditor")
        self.assertEqual(creditor_mapping.target_account_code, "410002227")

        expense_account = wizard._ensure_target_account(expense_mapping)
        creditor_account = wizard._ensure_target_account(creditor_mapping)
        expense_partner = wizard._ensure_partner(expense_mapping)
        creditor_partner = wizard._ensure_partner(creditor_mapping)

        self.assertEqual(expense_account.code, "629010014")
        self.assertEqual(creditor_account.code, "410002227")
        self.assertIn("monitor", expense_partner.name.lower())
        self.assertIn("monitor", creditor_partner.name.lower())
        self.assertTrue(self.env["account.account"].search([("code", "=", "629010014")], limit=1))
        self.assertTrue(self.env["account.account"].search([("code", "=", "410002227")], limit=1))

    def test_12_maps_phone_supplier_expense_and_creates_optional_partner(self):
        wizard = self._create_wizard([
            [12, date(2026, 1, 12), "629050002", "TELÉFONO", "30,00", "", "TELEFONÍA MÓVIL(VODAFONE)", "TEL-12"],
            [12, date(2026, 1, 12), "121000000", "TELÉFONO", "", "30,00", "Contrapartida", "TEL-12"],
        ], self._default_account_list_rows(), self._default_account_plan_rows(), create_missing_accounts=True)

        wizard.action_analyze()
        mapping = wizard.mapping_line_ids.filtered(lambda line: line.source_code == "629050002")

        self.assertEqual(mapping.mapping_type, "expense_nature_by_supplier")
        self.assertEqual(mapping.target_account_code, "629050002")
        self.assertEqual(mapping.partner_name, "Vodafone")

        target_account = wizard._ensure_target_account(mapping)
        partner = wizard._ensure_partner(mapping)

        self.assertEqual(target_account.code, "629050002")
        self.assertEqual(partner.name, "Vodafone")
        self.assertTrue(self.env["account.account"].search([("code", "=", "629050002")], limit=1))

    def test_12b_auto_creates_known_expense_accounts_without_global_flag(self):
        account_629010014 = self.env["account.account"].search([
            ("code", "=", "629010014"),
            ("company_ids", "in", self.company.id),
        ], limit=1)
        if account_629010014 and not self.env["account.move.line"].search_count([("account_id", "=", account_629010014.id)]):
            account_629010014.unlink()

        wizard = self._create_wizard([
            [121, date(2026, 1, 12), "629010014", "SERVICIO", "100,00", "", "MONITOR INFORMÁTICA", "FAC-MON-121"],
            [121, date(2026, 1, 12), "410002227", "SERVICIO", "", "100,00", "MONITOR INFORMÁTICA", "FAC-MON-121"],
        ], [
            ["410002227", "Monitor Informática", "", "acreedor"],
            ["629010014", "Monitor Informática", "", "gasto"],
        ], [
            ["410002227", "Monitor Informática", "", "acreedor"],
            ["629010014", "Monitor Informática", "", "gasto"],
        ])

        wizard.action_analyze()
        expense_mapping = wizard.mapping_line_ids.filtered(lambda line: line.source_code == "629010014")

        self.assertFalse(wizard.create_missing_accounts)
        self.assertEqual(expense_mapping.target_account_code, "629010014")
        if not expense_mapping.target_account_id:
            self.assertIn("se creará durante la importación", expense_mapping.warning_message)

        created_account = wizard._ensure_target_account(expense_mapping)

        persisted_account = self.env["account.account"].search([
            ("code", "=", "629010014"),
            ("company_ids", "in", self.company.id),
        ], limit=1)
        self.assertEqual(created_account.code, "629010014")
        self.assertEqual(persisted_account, created_account)

    def test_13_preserves_material_informatico_subaccount(self):
        wizard = self._create_wizard([
            [13, date(2026, 1, 13), "629040000", "COMPRA PORTÁTIL", "350,00", "", "Portátil Lenovo", "IT-13"],
            [13, date(2026, 1, 13), "121000000", "COMPRA PORTÁTIL", "", "350,00", "Contrapartida", "IT-13"],
        ], self._default_account_list_rows(), self._default_account_plan_rows(), create_missing_accounts=True)

        wizard.action_analyze()
        mapping = wizard.mapping_line_ids.filtered(lambda line: line.source_code == "629040000")

        self.assertEqual(mapping.mapping_type, "expense_nature")
        self.assertEqual(mapping.target_account_code, "629040000")

    def test_14_warns_for_412000000_but_keeps_specific_payable_account(self):
        wizard = self._create_wizard([
            [14, date(2026, 1, 14), "412000000", "PAGO BECARIO", "100,00", "", "BECARIOS", "BEC-14"],
            [14, date(2026, 1, 14), "121000000", "PAGO BECARIO", "", "100,00", "Contrapartida", "BEC-14"],
        ], [
            ["412000000", "BECARIOS", "", "acreedor"],
        ], [
            ["412000000", "BECARIOS", "", "acreedor"],
        ], create_missing_accounts=True)

        wizard.action_analyze()
        mapping = wizard.mapping_line_ids.filtered(lambda line: line.source_code == "412000000")

        self.assertEqual(mapping.mapping_type, "grant_or_becario_payable")
        self.assertEqual(mapping.target_account_code, "412000000")
        self.assertIn("Confirmar con asesor", mapping.warning_message)

    def test_15_matches_217_assets_with_2817_by_family_and_reference(self):
        wizard = self._create_wizard([
            [15, date(2026, 1, 15), "217001087", "APERTURA", "2556,20", "", "NAP TS-INF. SERRANO Y MAS (P24-22)", "P24-22"],
            [15, date(2026, 1, 15), "281701087", "APERTURA", "", "1200,00", "Amort. NAP TS-INF. SERRANO Y MAS (P24-22)", "P24-22"],
            [15, date(2026, 1, 15), "121000000", "APERTURA", "", "1356,20", "Contrapartida", "P24-22"],
        ], self._default_account_list_rows(), self._default_account_plan_rows(), create_missing_accounts=True)

        wizard.action_analyze()
        asset_mapping = wizard.mapping_line_ids.filtered(lambda line: line.source_code == "217001087")
        depreciation_mapping = wizard.mapping_line_ids.filtered(lambda line: line.source_code == "281701087")
        asset_preview = wizard.asset_line_ids.filtered(lambda line: line.source_code == "217001087")

        self.assertEqual(asset_mapping.mapping_type, "asset")
        self.assertEqual(asset_mapping.target_account_code, "217001087")
        self.assertEqual(depreciation_mapping.mapping_type, "accumulated_depreciation")
        self.assertEqual(depreciation_mapping.target_account_code, "281701087")
        self.assertEqual(asset_preview.depreciation_source_code, "281701087")
        self.assertEqual(asset_preview.depreciation_target_account_id.code, "281701087")
        self.assertEqual(asset_preview.accumulated_value, 1200.0)
        self.assertEqual(asset_preview.residual_value, 1356.2)
        self.assertEqual(asset_preview.state, "mapped_with_asset")
        self.assertTrue(asset_preview.create_asset)

    def test_16_marks_217_asset_without_depreciation_for_review(self):
        wizard = self._create_wizard([
            [16, date(2026, 1, 16), "217001089", "APERTURA", "900,00", "", "MONITOR SAMSUNG (P24-30)", "P24-30"],
            [16, date(2026, 1, 16), "121000000", "APERTURA", "", "900,00", "Contrapartida", "P24-30"],
        ], [
            ["217001089", "Monitor Samsung (P24-30)", "", "activo"],
        ], [
            ["217001089", "Equipos proceso información", "", "activo"],
        ], create_missing_accounts=True)

        wizard.action_analyze()
        asset_preview = wizard.asset_line_ids.filtered(lambda line: line.source_code == "217001089")

        self.assertEqual(asset_preview.target_account_id.code, "217001089")
        self.assertEqual(asset_preview.depreciation_target_account_id.code, "281701089")
        self.assertEqual(asset_preview.state, "review_required")
        self.assertFalse(asset_preview.create_asset)
        self.assertIn("No se ha encontrado amortización acumulada asociada", asset_preview.warning_message)

    def test_17_persists_detected_assets_in_review_batch(self):
        wizard = self._create_wizard([
            [17, date(2026, 1, 17), "217001087", "APERTURA", "2556,20", "", "NAP TS-INF. SERRANO Y MAS (P24-22)", "P24-22"],
            [17, date(2026, 1, 17), "281701087", "APERTURA", "", "1200,00", "Amort. NAP TS-INF. SERRANO Y MAS (P24-22)", "P24-22"],
            [17, date(2026, 1, 17), "121000000", "APERTURA", "", "1356,20", "Contrapartida", "P24-22"],
        ], self._default_account_list_rows(), self._default_account_plan_rows(), create_missing_accounts=True)

        wizard.action_analyze()

        self.assertTrue(wizard.preview_batch_id)
        self.assertTrue(wizard.preview_batch_id.dry_run)
        self.assertEqual(wizard.preview_batch_id.state, "analyzed")
        self.assertEqual(wizard.preview_batch_id.asset_line_count, 1)

        persisted_asset = wizard.preview_batch_id.asset_preview_line_ids.filtered(lambda line: line.source_code == "217001087")
        self.assertTrue(persisted_asset)
        self.assertEqual(persisted_asset.target_account_id.code, "217001087")
        self.assertEqual(persisted_asset.depreciation_target_account_id.code, "281701087")
        self.assertEqual(persisted_asset.state, "mapped_with_asset")
        self.assertEqual(persisted_asset.source_reference, "P24-22")

    def test_18_confirm_import_reuses_review_batch(self):
        wizard = self._create_wizard([
            [18, date(2026, 1, 18), "430000001", "FACTURA", "121,00", "", "Cliente A", "FAC-18"],
            [18, date(2026, 1, 18), "700000123", "FACTURA", "", "100,00", "Venta", "FAC-18"],
            [18, date(2026, 1, 18), "472000921", "FACTURA", "", "21,00", "IVA soportado 21%", "FAC-18"],
        ], self._default_account_list_rows(), self._default_account_plan_rows(), create_missing_accounts=True)

        wizard.action_analyze()
        preview_batch = wizard.preview_batch_id
        entries, hashes = wizard._build_import_entries()
        reused_batch = wizard._create_batch(hashes)

        self.assertEqual(reused_batch, preview_batch)
        self.assertEqual(wizard.preview_batch_id, preview_batch)
        self.assertFalse(reused_batch.dry_run)
        self.assertEqual(reused_batch.state, "analyzed")
        self.assertEqual(reused_batch.import_signature, preview_batch.import_signature)
        self.assertTrue(entries)

    def test_19_account_user_can_persist_asset_preview_lines(self):
        account_user = self._create_account_user()
        wizard = self._create_wizard([
            [19, date(2026, 1, 19), "217001087", "APERTURA", "1200,00", "", "NAP TS-INF. SERRANO Y MAS (P24-22)", "ACT-19"],
            [19, date(2026, 1, 19), "281701087", "APERTURA", "", "700,00", "Amort. NAP TS-INF. SERRANO Y MAS (P24-22)", "ACT-19"],
            [19, date(2026, 1, 19), "121000000", "APERTURA", "", "500,00", "Contrapartida", "ACT-19"],
        ], self._default_account_list_rows(), self._default_account_plan_rows(), create_missing_accounts=True)

        try:
            wizard.with_user(account_user).action_analyze()
        except AccessError as exc:
            self.fail(f"action_analyze no debe fallar por ACL al persistir activos detectados: {exc}")

        preview_batch = wizard.with_user(account_user).preview_batch_id
        self.assertTrue(preview_batch)
        self.assertEqual(len(preview_batch.asset_preview_line_ids), 1)
        self.assertEqual(preview_batch.asset_preview_line_ids.source_code, "217001087")

    def test_20_account_user_can_create_import_trace_lines_on_confirm(self):
        account_user = self._create_account_user()
        wizard = self._create_wizard([
            [20, date(2026, 1, 20), "430000001", "FACTURA", "121,00", "", "Cliente A", "FAC-20"],
            [20, date(2026, 1, 20), "700000123", "FACTURA", "", "100,00", "Venta", "FAC-20"],
            [20, date(2026, 1, 20), "472000921", "FACTURA", "", "21,00", "IVA soportado 21%", "FAC-20"],
        ], self._default_account_list_rows(), self._default_account_plan_rows(), create_missing_accounts=True)

        wizard = wizard.with_user(account_user)

        try:
            wizard.action_analyze()
            wizard.action_confirm_import()
        except AccessError as exc:
            self.fail(f"action_confirm_import no debe fallar por ACL al crear trazas: {exc}")

        batch = wizard.preview_batch_id
        self.assertTrue(batch)
        self.assertFalse(batch.dry_run)
        self.assertEqual(len(batch.trace_line_ids), 3)
        self.assertTrue(all(batch.trace_line_ids.mapped("move_line_id")))

    @unittest.skipUnless(_ACCOUNT_ASSET_AVAILABLE, "Requiere módulo Enterprise account_asset")
    def test_21_links_posted_imported_asset_move_lines_to_asset(self):
        wizard = self._create_wizard([
            [21, date(2026, 1, 21), "217001087", "APERTURA", "2556,20", "", "NAP TS-INF. SERRANO Y MAS (P24-22)", "P24-22"],
            [21, date(2026, 1, 21), "281701087", "APERTURA", "", "1200,00", "Amort. NAP TS-INF. SERRANO Y MAS (P24-22)", "P24-22"],
            [21, date(2026, 1, 21), "121000000", "APERTURA", "", "1356,20", "Contrapartida", "P24-22"],
        ], self._default_account_list_rows(), self._default_account_plan_rows(), create_missing_accounts=True, post_moves=False)

        wizard.action_analyze()
        wizard.action_confirm_import()

        asset = wizard.asset_line_ids.filtered(lambda line: line.source_code == "217001087").asset_id
        move = wizard.preview_batch_id.trace_line_ids.move_id[:1]

        self.assertTrue(asset)
        self.assertFalse(asset.original_move_line_ids)
        self.assertEqual(move.state, "draft")

        move.action_post()
        asset.invalidate_recordset()

        self.assertTrue(asset.original_move_line_ids)
        self.assertEqual(asset.original_move_line_ids.account_id.code, "217001087")
        self.assertEqual(asset.original_move_line_ids.move_id, move)
        self.assertIn(asset, move.asset_ids)

    @unittest.skipUnless(_ACCOUNT_ASSET_AVAILABLE, "Requiere módulo Enterprise account_asset")
    def test_22_copies_depreciation_accounts_from_asset_model_on_import(self):
        asset_model = self.env["account.asset"].search([
            ("state", "=", "model"),
            ("company_id", "=", self.company.id),
            ("account_asset_id", "=", self.account_217.id),
            ("account_depreciation_id", "=", self.account_281700.id),
        ], limit=1)
        if not asset_model:
            asset_model = self.env["account.asset"].create({
                "name": "Modelo equipos importador",
                "state": "model",
                "company_id": self.company.id,
                "account_asset_id": self.account_217.id,
                "account_depreciation_id": self.account_281700.id,
                "account_depreciation_expense_id": self.account_681700.id,
                "journal_id": self.journal.id,
                "method": "linear",
                "method_number": 5,
                "method_period": 12,
                "prorata_computation_type": "constant_periods",
            })
        else:
            asset_model.write({
                "account_depreciation_expense_id": self.account_681700.id,
                "journal_id": self.journal.id,
            })

        wizard = self._create_wizard([
            [22, date(2026, 1, 22), "217001087", "APERTURA", "2556,20", "", "NAP TS-INF. SERRANO Y MAS (P24-22)", "P24-22"],
            [22, date(2026, 1, 22), "281701087", "APERTURA", "", "1200,00", "Amort. NAP TS-INF. SERRANO Y MAS (P24-22)", "P24-22"],
            [22, date(2026, 1, 22), "121000000", "APERTURA", "", "1356,20", "Contrapartida", "P24-22"],
        ], self._default_account_list_rows(), self._default_account_plan_rows(), create_missing_accounts=True)

        wizard.action_analyze()
        wizard.action_confirm_import()

        asset = wizard.asset_line_ids.filtered(lambda line: line.source_code == "217001087").asset_id

        self.assertTrue(asset)
        self.assertEqual(asset.account_asset_id.code, "217001087")
        self.assertEqual(asset.account_depreciation_id.code, "281701087")
        self.assertEqual(asset.account_depreciation_expense_id, self.account_681700)
        self.assertEqual(asset.journal_id, self.journal)

    def test_11_maps_623_subaccounts_preserving_exact_code(self):
        wizard = self._create_wizard([
            [11, date(2026, 1, 11), "623010203", "GASTO", "75,00", "", "Servicios contables", "G-11"],
            [11, date(2026, 1, 11), "121000000", "GASTO", "", "75,00", "Contrapartida", "G-11"],
        ], [
            ["623010203", "Servicios contables", "", "gasto"],
        ], [
            ["623010203", "Servicios contables", "", "gasto"],
        ], create_missing_accounts=True)

        wizard.action_analyze()
        mapping = wizard.mapping_line_ids.filtered(lambda line: line.source_code == "623010203")
        self.assertEqual(mapping.mapping_type, "direct_account")
        self.assertEqual(mapping.target_account_code, "623010203")

        action = wizard.action_confirm_import()
        move = self.env["account.move"].browse(action["res_id"])
        expense_line = move.line_ids.filtered(lambda line: line.account_id.code == "623010203")
        self.assertTrue(expense_line)
