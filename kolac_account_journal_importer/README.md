Kolac Account Journal Importer
================================

Módulo para migrar contabilidad histórica desde Excel a una estructura nativa de Odoo 19.

Flujo funcional
---------------

El flujo está pensado para trabajar con el Diario de movimientos oficial de KOLAC.

- Libro Diario oficial

Con ese fichero el wizard:

- conserva el código original de cada subcuenta contable;
- detecta clientes, proveedores y acreedores para usar partner_id sin reagrupar cuentas;
- detecta activos y amortizaciones acumuladas para revisión previa;
- marca impuestos y advertencias de mapeo;
- permite importar directamente cuando el fichero ya viene en el formato oficial del diario de KOLAC;
- bloquea la importación definitiva solo si faltan datos mínimos o hay errores bloqueantes;
- genera trazabilidad por lote de importación.

Resultado buscado
-----------------

La migración respeta las subcuentas antiguas y crea en Odoo los mismos códigos cuando no existan todavía. En particular:

- 430xxxxxx mantiene su subcuenta de cliente y añade `partner_id` cuando procede.
- 400xxxxxx y 410xxxxxx mantienen su subcuenta de proveedor o acreedor y añaden `partner_id` cuando procede.
- 211/216/217/218/219xxxxxx mantienen su subcuenta exacta para inmovilizado.
- 2811/2816/2817/2818/2819xxxxxx mantienen su subcuenta exacta de amortización acumulada.
- el resto de grupos contables se importan respetando el código original del diario.

Menús
-----

- Contabilidad -> Accounting -> Transactions -> Importar apuntes contables
- Contabilidad -> Accounting -> Transactions -> Lotes importación contable
- Contabilidad -> Configuración -> Equivalencias subcuentas

Uso en Docker o Doodba
----------------------

1. Actualizar el módulo.

Ejecutar desde la raíz del workspace::

		docker compose run --rm -T odoo \
			odoo -d kolac19 -u kolac_account_journal_importer --stop-after-init

2. Ejecutar los tests del módulo::

		docker compose run --rm -T odoo \
			odoo -d kolac19 -u kolac_account_journal_importer \
			--test-tags /kolac_account_journal_importer --stop-after-init

3. Abrir el wizard en Odoo y seguir este orden:

- cargar el Diario oficial de KOLAC;
- ejecutar Analizar;
- revisar mapeos, activos, asientos preparados y logs;
- ejecutar Confirmar importación solo si no quedan errores bloqueantes.

Checklist operativo
-------------------

- Verificar que el Diario oficial de KOLAC tiene cabeceras tipo `Fecha`, `Asto.`, `Cuenta`, `Ttulo`, `Concepto`, `Debe`, `Haber`.
- Verificar que el diario seleccionado es de tipo general y de la compañía correcta.
- Revisar que las subcuentas creadas en Odoo respetan el código original del diario.
- Confirmar que las cuentas 430, 400 y 410 mantienen su subcuenta exacta y, cuando aplique, quedan vinculadas a un contacto.
- Confirmar que inmovilizado y amortización conservan su subcuenta exacta.
- Revisar advertencias fiscales de 472 y 477 antes de importar definitivamente.
- Revisar los asientos descuadrados o duplicados detectados por el análisis.
- Guardar equivalencias manuales útiles para reutilizarlas en importaciones posteriores.
- Consultar el lote creado tras la importación para validar trazabilidad e informe guardado.

Validación mínima recomendada
-----------------------------

- Actualizar el módulo en la base objetivo.
- Ejecutar la batería de tests del addon.
- Hacer una pasada de análisis con el Diario oficial real antes de la primera importación definitiva.
- Confirmar en un lote importado que el informe de revisión y la trazabilidad se han guardado correctamente.

