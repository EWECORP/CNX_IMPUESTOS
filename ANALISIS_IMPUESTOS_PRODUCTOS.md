# Impuestos asociados a productos en DIARCO-SGM

Análisis reconstruido solo desde evidencia del repositorio (se excluyó `sql-cdcapiworkersolution-sgmconnexaapi-recepcionesolution`).

**Estados:** CONFIRMADO (evidencia directa) · INFERIDO · NO DETERMINADO.

**Limitación estructural (CONFIRMADO):** el repo contiene código VB6 (`.frm/.bas`), SPs de la integración Connexa (`SQL/`) y C#. **No contiene DDL de tablas, vistas, triggers ni los SPs de facturación/compras** (`SD03_*`, `SD04_*`, `SD99_*`, SPs de grabación de comprobantes). Por eso tipos, nullability, PK/FK y los cálculos hechos dentro de SPs son **NO DETERMINADO**, salvo lo inferible por uso. Rutas abreviadas: `SUC`=`01_Gestion/04_Modulo Sucursal`, `MERC`=`01_Gestion/03_Modulo Mercaderias`, `EDI`=`01_Gestion/15  FC EDI`, `PAGO`=`01_Gestion/02_Modulo Pago a Proveedores`, `TES`=`01_Gestion/99_Modulo Tesoreria`. Los formularios están copiados en varios módulos (y hay `- copia`, `Copia de`, `Luke`): se cita la copia de referencia.

---

## PARTE I — RESUMEN EJECUTIVO

1. **Modelo.** El IVA no se guarda como porcentaje en el producto: el artículo (`T050_ARTICULOS`) guarda **códigos** (`C_IVA`, `C_IVA_VENTAS`, `C_IVA_VENTAS_CF`) que apuntan al maestro `T110_IVA`, cuyo campo `K_COEF_IVA` es el **coeficiente decimal** (0,21 = 21 %; el código lo multiplica por 100 para mostrarlo).
2. **Impuestos existentes (CONFIRMADO en código):** IVA (alícuotas general, reducida, exenta; la compra reconoce además 27 %), impuestos internos (por artículo/proveedor), percepción de IVA, percepciones de Ingresos Brutos por provincia (varias), percepción de bebidas, retención de IVA, Decreto 1344 (seguridad e higiene), TISH (Corrientes/Salta/Posadas/Comodoro), Tasa de protección ambiental (Pilar, sucursal 13) y percepción de ganancias (solo en compras).
3. **Asociación a productos.** Mecanismo = FK por código en `T050_ARTICULOS`. **Hay dos códigos de IVA de venta** (`C_IVA_VENTAS` para Responsable Inscripto y `C_IVA_VENTAS_CF` para el resto) más `C_IVA` (compra/costo, usado en OC y devoluciones a proveedor). Los impuestos internos se guardan como coeficiente por artículo-proveedor (`K_IMP_INTERNOS`). Las percepciones de IIBB dependen de `C_IB` del artículo + provincia/padrón del cliente.
4. **Carga/mantenimiento.** El maestro `T110_IVA` no tiene pantalla de ABM en el repo (solo lecturas) → alta probablemente por script/DBA (NO DETERMINADO). El código de IVA del artículo se edita en el ABM de artículos (`MERC/Articulos_Gestion_Baja.frm` carga `oCbo_Iva`; `GblF_Load_IvaCombobox`). Existe además un archivo `seguridad.ini` `[AlicuotaIva]` con alícuotas 1-3 y `MultipleAlicuota` que, en los módulos EDI/Pago a Proveedores, **reemplaza** al maestro.
5. **Tasa.** Se lee de `T110_IVA.K_COEF_IVA` por `C_IVA` (`GblF_CoefIVA`, redondeada a 3 decimales). **No hay vigencia ni histórico en `T110_IVA`** (sin columnas de fecha detectadas en ningún uso).
6. **Cálculo (ventas, CONFIRMADO).** El precio de venta de lista (`I_PRECIO_VTA`) es **neto**; precio final = `neto × (1 + coef)`; IVA de comprobante = `Round(NetoGrupo × coef, 3)` por grupo de alícuota. Para facturas A el precio unitario se muestra neto; para B/consumidor final, con IVA.
7. **Compras (CONFIRMADO).** Totales por alícuota 21 / 10,5 / 27 (`i_neto_2100`, `i_iva_2100`, ...), percepciones, impuestos internos; redondeo a **2 decimales** en comparaciones y notas de crédito (vs. 3 en ventas). Se compara lo facturado contra lo recibido (parte) y las diferencias generan SNC/SND.
8. **Persistencia.** Cada línea de venta persiste `@K_IVA` (coeficiente aplicado); la cabecera persiste condición IVA del cliente, `M_EXENTO_IVA`, percepciones. Los históricos conservan lo grabado (INFERIDO: depende de los SPs no incluidos) pero **los reportes de Libro IVA rotulan columnas con la alícuota vigente actual** (`Facturador_CptesConDescuentos.LcLF_GetCoefIva`).
9. **Riesgos principales:** doble fuente de alícuotas (tabla vs `seguridad.ini`), alícuotas y sucursales hardcodeadas (37 y 48 sin IVA/percepción), GUID de IVA 21 % hardcodeado en Connexa, lógica de percepción con fechas históricas hardcodeadas, código duplicado en ≥6 módulos, mezcla de 2 y 3 decimales, N+1 de consultas para obtener coeficientes.

---

## PARTE II — RESUMEN TÉCNICO

### 2.1 Arquitectura
VB6 → ADO → SQL Server; lectura del maestro mediante SQL dinámico (`SQL_DINAMICO` en ventas) o `SELECT` directo (compras). Integración con Connexa por SPs T-SQL (`SQL/`). Sin servicios C# con lógica impositiva (las referencias a "tax" en `SGMConnexa.Api` son artefactos genéricos de HelpPage).

### 2.2 Catálogo de impuestos (Matriz — PARTE IX)

| Impuesto | Código / identificador | Tasa | Aplica a | Base | Vigencia | Incluido en precio | Módulos |
|---|---|---|---|---|---|---|---|
| IVA | `T110_IVA.C_IVA`: 1 = alícuota 1 (general), 2 = alícuota 2 (reducida), 4 = exento (en grilla de venta `vaSpread1` col.20 = 4 → "ALICUOTA EXENTA"). `3` aparece en `C_COSTO_IVA` con otro significado (ver 2.4) | `K_COEF_IVA` (decimal). Valores concretos de 1 y 2: NO DETERMINADO en repo (docs EDI citan 21 / 10,5 / 27 como `[AlicuotaIva]` del ini) | Artículo (venta y compra) | Neto de línea tras descuentos | No hay (sin fechas) | Precio de venta de lista: **no** incluido (neto). Lista de proveedor: depende de `C_COSTO_IVA` | Sucursal, Clientes, Mercaderías, EDI, Pago Prov., Cartelería, Libro IVA |
| Impuesto interno | `K_IMP_INTERNOS` / `K_IMP_INTERNO` (coef.), `I_IMP_INTERNO`, `I_TOTAL_IMP_INTERNO` | Coeficiente por artículo-proveedor | Artículo-proveedor (compras); en venta se resta del precio antes del IVA (cartelería) | Costo base / precio | NO DETERMINADO | Se suma al costo de lista (`GlbF_Calculo_Plista_Prov`) | Mercaderías, EDI, Connexa |
| Percepción IVA | `T110_IVA_PERCEPCION` (`C_COND_IVA`, `K_COEF_IVA`, `I_BASE_MINIMA`) | Por condición IVA del cliente; neto alíc.2 a la mitad | Cliente (cond. IVA) | Neto1 + Neto2/2 | Sin fechas en tabla; en compras hay tramos por fecha hardcodeados | No | Sucursal, EDI (`GblF_CorrijoIvaPercepcion1`), Pago Prov. |
| Percepción / retención IIBB | `T050_ARTICULOS.C_IB`, `LclF_GetCoefIb`, `Gbl_AlicuotaIb*` por provincia, `C_PROVINCIA`/`Gbl_ProvinciaFiscal` | Padrón/alícuota por provincia | Artículo × cliente × provincia | Neto (o neto+IVA según provincia/condición) | NO DETERMINADO | No | Sucursal, Clientes |
| Decreto 1344, TISH, Tasa Prot. Ambiental | `oTxt_Decreto1344`, `Gbl_TasaProtAmb` | Parámetros | Cliente/sucursal | Neto | NO DETERMINADO | No | Facturador |
| Retención IVA | `I_RETENCION_IVA`, `M_AGENTE_RETENCION_IVA` | NO DETERMINADO | Cliente agente | — | — | — | Si es agente de retención IVA la percepción se anula (Facturador_Articulos ~11463) |

### 2.3 Tablas (inventario)

| Tabla | Propósito | Procesos que leen | Procesos que escriben |
|---|---|---|---|
| `T110_IVA` (`C_IVA`, `K_COEF_IVA`) | Maestro de alícuotas (clave presumida `C_IVA`) | `GblF_CoefIVA` (MERC/ModuloGeneral.bas:2113), `LcLF_GetCoefIva` (Facturador_CptesConDescuentos), JOINs en Facturador_Articulos, `OrdenCompra_Kikker_Generar`, `SP_SGM_CONNEXA_OC_GET`, `SP_STOCK_ADJUSTMENT_DEVOLUCION_PROVEEDOR` | Ninguno en el repo (NO DETERMINADO) |
| `T110_IVA_PERCEPCION` (`C_COND_IVA`, `K_COEF_IVA`, `I_BASE_MINIMA`) | Parámetros de percepción de IVA por condición | `TES/ModuloImpuestosEspeciales.bas:369,413`, `PAGO/Facturacion_SNC2.frm:1922` | Ninguno |
| `T050_ARTICULOS` (`C_IVA`, `C_IVA_VENTAS`, `C_IVA_VENTAS_CF`, `C_IB`) | Maestro de artículos | Facturador, cartelería, OC | ABM artículos (SP no incluido) |
| `T231_PEDIDOS_CABE` (`C_COND_IVA_CLIENTE`, `M_EXENTO_IVA`), `T231_PEDIDOS_DETA` (`I_PRECIO_UNIT_ART`, `I_PRECIO_UNIT_ART_FINAL`, `K_IVA`…) | Pedidos de venta | Facturador_Articulos | SPs de pedidos |
| `T230_FACTURADOR_PARAM_COND_IVA` | Parámetros del facturador por condición IVA (p. ej. `I_TOPE_MAXIMO_COMP`) | Facturador_Inicio:6238 | NO DETERMINADO |
| `T001_TABLA_CODIGO` (`C_TABLA=12`) | Descripciones de condición IVA | `GblF_GetDescripT001(12, …)` | — |
| `T020`/clientes (`C_COND_IVA`, `M_EXENTO_IVA`, `M_CERT_PYME`, `M_EXCLUIDO_PERCEPCION_IVA`, `F_EXCLUIDO_PERCEP_IVA_VENC`, `K_LIBERACION_PERCEPCION_IVA`, `M_AGENTE_RETENCION_IVA`) | Condición fiscal del cliente | Facturador_Inicio | ABM clientes |
| `T023_PROV_PARTES_DETA` (`I_TOTAL_IMP_INTERNO`), OC (`I_IVA_OC`, `I_IMP_INTERNO_OC`, líneas `C_IVA_EN_CALCULO`, `K_COEF_IVA`, `K_IMP_INTERNO`, `I_IMP_INTERNO`, `I_TOTAL_IMP_INTERNO`) | Partes y órdenes de compra | EDI, Connexa | `SP_SGM_CONEXXA_OC_CREATE`, EDI |
| Comprobantes proveedor (`i_neto_2100/1050/2700`, `i_iva_2100/1050/2700`, `i_perc_iva`, `i_perc_bebidas`, `i_perc_ganancias`, `i_impuestos_inter`) | Facturas/SNC/SND de proveedor | Pago Prov., Contable | SP invocado desde `EDI/ComprobanteProveedor_Alta.frm:8300-8373` (nombre/tabla no visibles aquí) |

### 2.4 Diccionario de columnas clave (PARTE VI)

| Campo | Tipo | Oblig. | Significado | Uso | Valores | Evidencia |
|---|---|---|---|---|---|---|
| `T110_IVA.C_IVA` | NO DET. (numérico) | Sí | Código de alícuota | JOIN con artículo | 1,2,4 observados | `Facturador_Articulos.frm:7738`, grilla col.20 (~11368) |
| `T110_IVA.K_COEF_IVA` | NO DET. (decimal) | Sí | **Coeficiente decimal** (0,21 = 21 %) | `precio*(1+K)`; ×100 para mostrar | — | `oTxt_CoefIva = K*100`, `.Tag = K` (7883) |
| `T050.C_IVA` | NO DET. | Sí | IVA del artículo en compras/OC/stock | OC, devolución prov. | — | `SP_SGM_CONNEXA_OC_GET.sql:111`, `Gblf_GetArticuloIva` |
| `T050.C_IVA_VENTAS` | NO DET. | Sí | IVA de venta si cliente RI (cond. 1) y no PyME | Facturador | — | `Facturador_Articulos.frm:7795-7811, 9004` |
| `T050.C_IVA_VENTAS_CF` | NO DET. | Sí | IVA de venta para resto de condiciones (CF, monotributo, etc.) | Facturador, cartelería | — | idem 7722-7738 |
| `C_COSTO_IVA` (lista prov.) | NO DET. | — | Modo en que el costo de lista trata el IVA | `GblF_CoefIvaArtComp` | 1 = "Medio IVA" → `(1+k)/(1+k/2)`; 2 = "No incluye" → factor 1; 3 = "Incluye IVA" → `(1+k)` | `MERC/ModuloGeneral.bas:2164-2177` |
| `K_IVA` (línea) | — | — | Coeficiente de IVA aplicado en la línea | Se envía a SP de grabación | — | `Facturador_Articulos.frm:13442,14000,17142` |
| `C_COND_IVA_CLIENTE` | — | — | Condición IVA del cliente en el comprobante | 1=RI (según código), 4=Consumidor Final, 5=No categorizado ("Solo Para No Categorizados"), 3 y 6 con tratamiento propio (descripción en `T001` tabla 12: NO DETERMINADO) | — | Facturador_Inicio ~6942-6990; Facturador_Articulos ~11452 |
| `M_EXENTO_IVA` | char S/N | — | Cliente/comprobante exento: si ≠"N" no se calcula precio con IVA | — | 'S','N' | Facturador_Articulos ~10546; `oCbo_iva_Change` |
| `I_BASE_MINIMA` | — | — | Mínimo de base/importe bajo el cual no se cobra percepción IVA | — | 9999999 si no hay fila | `ModuloImpuestosEspeciales.bas:GblF_MinimoPercepcionIva` |
| `K_IMP_INTERNOS` | — | — | Coef. de impuesto interno sobre costo base | lista proveedor | 0 = sin interno | `GlbF_Calculo_Plista_Prov` |
| `C_IVA_EN_CALCULO` (línea OC) | — | — | Mismo concepto que `C_COSTO_IVA` | Connexa lo fija en **2** (precio no incluye IVA) | — | `SP_SGM_CONEXXA_OC_CREATE.sql:390` |

---

## PARTE III — RELACIÓN PRODUCTO ↔ IMPUESTO

| Mecanismo | Origen | Regla | Prioridad / fallback |
|---|---|---|---|
| FK `T050.C_IVA_VENTAS` → `T110_IVA` | Artículo | Venta a cliente con `C_COND_IVA_CLIENTE=1` y `M_CERT_PYME='N'` | Si cliente ≠ RI o es PyME: `C_IVA_VENTAS_CF` |
| FK `T050.C_IVA_VENTAS_CF` → `T110_IVA` | Artículo | Resto | — |
| FK `T050.C_IVA` | Artículo | Compras, OC, stock/devolución a proveedor | — |
| Código por lista de proveedor `C_COSTO_IVA` | Artículo×proveedor×sucursal | Define si el costo de lista incluye IVA | — |
| Coeficiente `K_IMP_INTERNOS` | Artículo×proveedor | Impuesto interno | 0 = no aplica |
| `C_IB` + provincia | Artículo×cliente | Percepción IIBB | Coeficiente 0 para CF en Capital Federal (`LclF_GetCoefIb`) |
| Hardcode sucursal | Sucursal 37 y 48 | `oTxt_CoefIva=0` y `PercepIva=0` | Anula IVA (CONFIRMADO, Facturador_Articulos ~7882, 8995, 11490) |
| Fallback sin fila en `T110_IVA` | — | `LcLF_GetCoefIva` → 0; `GblF_CoefIVA` → 0 + MsgBox | **El producto queda sin IVA silenciosamente en el facturador**; EDI valida con error "No se encontro codigo de IVA en la Tabla T110_IVA" |

Un producto tiene **un** código de IVA por contexto (no hay tabla producto-impuesto N:M) → "varios impuestos" se logra por combinación de IVA + interno + IIBB, no por relación múltiple (CONFIRMADO por ausencia de tabla intermedia; la existencia en el DDL es NO DETERMINADO).

---

## PARTE IV — FLUJOS END-TO-END

**Flujo 1 – Configuración de alícuota.** `T110_IVA` (alta externa, NO DET.) → disponible vía SELECT en cada uso. Sin caché ni vigencia: **cambiar `K_COEF_IVA` afecta de inmediato todo cálculo nuevo y cualquier listado que re-lea la tabla**. Alternativa: `seguridad.ini [AlicuotaIva]` (`Alicuota_1..3`, `MultipleAlicuota`; `EDI/Module1.bas:265-300`, `PAGO/ModuloConexion.bas:374`): con `MultipleAlicuota=1`, `GblF_CalculoIVA` usa el ini; con 0 usa tabla.

**Flujo 2 – Asociación a producto.** ABM artículo → combos `C_IVA*` (cargados por `GblF_Load_IvaCombobox`) → grabación en `T050_ARTICULOS` (SP no incluido).

**Flujo 3 – Venta (Facturador).**
1. `Facturador_Inicio`: cliente → `C_COND_IVA`, `M_EXENTO_IVA`, `M_CERT_PYME`, IB, exclusiones de percepción. `oCbo_iva_Change` ajusta defaults (RI: percepción IVA no excluida; resto: liberación 100 %).
2. `Facturador_Articulos` (inicio): `Gbl_Alicuota1 = GblF_CoefIVA(1)`, `Gbl_Alicuota2 = GblF_CoefIVA(2)`; letra A si cond. IVA 1 ó 3, B si no; `Gbl_AlicuotaPercepcionIva`/`Gbl_MinimoPercepcionIva` desde `T110_IVA_PERCEPCION` por cond. IVA.
3. Por artículo: JOIN `T110_IVA` según `C_IVA_VENTAS` o `_CF` → `oTxt_CoefIva`.
4. Línea: col.5 = coef.; col.6 = precio neto `Round(PrecioBase,3)`; col.7 = precio con IVA `Round(PrecioBase*(1+coef),4)` (si `M_EXENTO_IVA="N"`); col.20 = tipo IVA (1/2/4) que decide en qué acumulador entra.
5. Totales: `Neto1`, `Neto2` (netos por alícuota tras descuentos + envases), `Iva1=Round(Neto1*coef1,3)`, `Iva2=Round(Neto2*coef2,3)`, `Exento`.
6. Percepciones y grabación: SPs reciben `@K_IVA` por línea, `@I_PERCEP_IVA`, `@C_COND_IVA_CLIENTE`, `@M_EXENTO_IVA`.

**Flujo 4 – Compra (EDI / FC electrónica).** Partes de recepción (`T023_PROV_PARTES_DETA`, incluye `I_total_imp_interno`) → `ComprobanteProveedor_Alta.frm` (netos/IVA 21/10,5/27, percepciones, internos) → comparación factura vs parte (coef. `vCoefIva1/2 = GblF_CoefIVA(1/2)`, `vCoefPercepIva = GblF_CorrijoIvaPercepcion1`) → generación de SNC/SND por diferencias (`vCoefDifMerca`, `vCoefDifPrecio`, `oMny_CoefDescProv`) → SP de alta de comprobante (`@i_neto_2100…@i_perc_iva…@i_impuestos_inter`) → SOAP/EDI y SAP según memoria del repo (anulaciones SOAP previas al `CommitTrans`).

**Flujo 5 – Orden de compra Connexa.**
- Entrada (`SP_SGM_CONEXXA_OC_CREATE`): del JSON `payload.taxes` toma `IvaRate` (si `tax.extCode='1'`, nombre 'IVA' o descripción 'IVA%') y `ImpuestosInternosRate` = suma de `customTaxAmount` de todo lo que no sea IVA. Calcula `I_IVA_OC = neto × rate/100`, `I_IMP_INTERNO_OC`, línea `K_COEF_IVA = rate/100`, `K_IMP_INTERNO`, `I_IMP_INTERNO = precio × rate/100`, `C_IVA_EN_CALCULO = 2`.
- Salida (`SP_CONNEXA_PO_SEND_ONE`): si `I_IVA_OC>0` envía un impuesto con `taxId` **GUID fijo "IVA 21"**; interno se envía como `customTaxName='Impuesto Interno SGM'`, `customTaxAmount = I_IMP_INTERNO_OC / I_NETO_OC × 100`.
- Consulta (`SP_SGM_CONNEXA_OC_GET`): `vatCoef` desde `T110_IVA` por `art.C_IVA` (puede diferir del IVA grabado en la OC).

**Flujo 6 – Cambio de alícuota.** Update de `T110_IVA` (externo). Nuevas operaciones usan el nuevo valor; líneas ya grabadas conservan `K_IVA` (INFERIDO: depende del SP). Reportes que usan `LcLF_GetCoefIva(1/2)` para títulos de columnas mostrarán la alícuota vigente aunque el período sea anterior (CONFIRMADO en rotulado; el cálculo de montos del libro: NO DET.).

---

## PARTE V — ALGORITMOS

**Venta (por línea y comprobante)**
```
coef  = T110_IVA.K_COEF_IVA[ cond_IVA=1 y no PyME ? art.C_IVA_VENTAS : art.C_IVA_VENTAS_CF ]
si sucursal in (37,48): coef = 0
neto_unit  = I_PRECIO_VTA (lista)            # almacenado SIN IVA
final_unit = Round(neto_unit*(1+coef), 3|4)  # 4 dec. en grilla col.7, 3 en recálculo de pedido
inverso    : neto = Round(final/(1+coef), 3) # pedidos con precio final (I_PRECIO_UNIT_ART)
por tipoIVA: 1→Neto1, 2→Neto2, 4→Exento
Neto1 = Round(Σneto1 − descuentos + envases, 3); Iva1 = Round(Neto1*coef1, 3)
Neto2 = Round(Σneto2 − descuentos, 3);          Iva2 = Round(Neto2*coef2, 3)
PercIVA(RI, no excluido) = Round(Neto1*c + Neto2*c/2, 3)  [× (1−liberación/100) si excluido]
PercIVA = 0 si agente de retención IVA, si < I_BASE_MINIMA, o sucursal 37/48
PercIVA(cond.5) = Round((Neto1+Iva1)*(coef1/2),3)+Round((Neto2+Iva2)*(coef2/2),3)  # sobrescrito luego por el bloque RI/excluido
IIBB = Round(Neto [×(1+coef) según provincia] × coefIB × (1−exclusión), 3)
Total = Neto1+Iva1+Neto2+Iva2+Exento+percepciones+otros
```
Orden real: precio → descuentos de línea/pie (sobre neto) → IVA por grupo → percepciones (base neto o neto+IVA) → total. Descuentos: el IVA se calcula **después** del descuento (`Neto1 = Σ − vTotalDescNeto1Art`). Promociones: ajustan el precio base (`vPrecioPromoUltimoBulto`, `K_DESC_APLIC_A_PRECIO_VTA`) antes del IVA (INFERIDO por flujo; ver Facturador_Articulos ~9423).

**Costo de lista de proveedor** (`GlbF_Calculo_Plista_Prov`): `calculo = costo_base / GblF_CoefIvaArtComp(C_COSTO_IVA, k_iva_1)`; si hay interno: `Round(costo*(1+k_iva)+costo*k_int, 3)`, descuentos secuenciales (código < 200 descuenta `×(1−k)`, ≥ 200 recarga `÷(1+k)`), y vuelve a neto `Round(calculo/(1+k_iva), 3)`. Más `I_ENVASE` y redondeo final a 3.

**Cartelería** (`Articulos_Cartel_ImpresionNew`): `PrecioConIVA = (Precio − ImpuestoInterno) × (1 + coef) `; el impuesto interno se trata como **no gravado por IVA y se suma después** (`Round(((P×(1+recargo)) − int) × (1+coef) + int, 3)`). Es una fórmula distinta a la de la lista de proveedor (donde `k_int × costo` se suma antes de descontar y se divide por (1+k_iva)).

**Compras:** `GblF_CalculoIVA(importe, cod)` = `Round(importe × coef, 3)` con coef desde ini (`MultipleAlicuota=1`) o tabla; en el comprobante: `Iva21Dif = Round(Neto21Dif × vCoefIva1, 2)`, `Iva10Dif = Round(Neto10Dif × vCoefIva2, 2)`; proporciones con 6 decimales; percepción IVA dif = `Round(|Neto − Parte| × coefPercep, 2)`.

**Percepción IVA en compras (histórica):** `GblF_CorrijoIvaPercepcion1` contiene **tramos por fecha hardcodeados** (p. ej. 23/05/2006–31/03/2007 → 0; 01/04/2007–30/09/2007 → 3 %/1,5 % ×0,01; mínimo exento 21,3). Es el único mecanismo de vigencia temporal encontrado y es código, no datos.

**Redondeo (CONFIRMADO):** ventas 3 decimales (4 en precio con IVA de grilla); compras 2 decimales en diferencias y 3 en `GblF_CalculoIVA`; coeficientes `Round(K,3)` en `GblF_CoefIVA` (**trunca alícuotas con más de 3 decimales**, p. ej. 0,105 OK, 0,1050x no).

---

## PARTE VI — TRAZABILIDAD

| Funcionalidad | Archivo | Método/SP | Tablas | Op. |
|---|---|---|---|---|
| Obtener coef. IVA | `MERC/ModuloGeneral.bas` | `GblF_CoefIVA` | T110_IVA | S |
| Obtener coef. (informes) | `SUC/Facturador_CptesConDescuentos.frm` | `LcLF_GetCoefIva` (vía `SQL_DINAMICO`) | T110_IVA | S |
| Modo IVA en costo de lista | `MERC/ModuloGeneral.bas` | `GblF_CoefIvaArtComp`, `GlbF_Calculo_Plista_Prov` | — | cálculo |
| IVA de venta por línea | `SUC/Facturador_Articulos.frm` | `LclF_GetPedidoArt`, rutina de totales ~11300-11500 | T050, T051, T110_IVA, T231_* | S |
| Percepción IVA | `TES/ModuloImpuestosEspeciales.bas` | `GblF_MinimoPercepcionIva`, `GblF_CoeficientePercepcionIva` | T110_IVA_PERCEPCION | S |
| IIBB | `SUC/Facturador_Articulos.frm` | `LclF_GetCoefIb` | padrones/T050.C_IB | S |
| Alícuotas por ini | `EDI/Module1.bas`, `PAGO/ModuloConexion.bas`, `PAGO/ModuloGeneral.bas` | `GblF_CalculoIVA` | seguridad.ini | cálculo |
| Compras: totales por alícuota | `EDI/ComprobanteProveedor_Alta.frm` | alta / SNC | SP alta comprobante | I |
| Internos de partes | idem | `LclF_ObtengoImpuestosInternos` | T023_PROV_PARTES_DETA | S |
| OC Connexa (entrada/salida/consulta) | `SQL/SP_SGM_CONEXXA_OC_CREATE.sql`, `SP_CONNEXA_PO_SEND_ONE.sql`, `SP_SGM_CONNEXA_OC_GET.sql` | idem | T110_IVA, OC | I/S |
| Devolución a proveedor | `SQL/SP_STOCK_ADJUSTMENT_DEVOLUCION_PROVEEDOR.sql` (+`SP Barrio`) | `I_IVA = SUM(I_COSTO_ESTADISTICO × K_COEF_IVA)` | T110_IVA | S/I |

Inventario SP/funciones/triggers/vistas con código visible: solo los SP citados arriba y las funciones VB6. Triggers y vistas: **NO DETERMINADO** (no hay DDL).

---

## PARTE VII — REGLAS DE NEGOCIO

| Regla | Evidencia | Estado |
|---|---|---|
| El IVA del artículo se determina por código a `T110_IVA`, no por porcentaje | `Facturador_Articulos` JOINs | CONFIRMADO |
| Precio de lista de venta es neto; final = neto×(1+coef) | col.6/7 grilla | CONFIRMADO |
| Cliente RI (cond.1) no PyME usa `C_IVA_VENTAS`; el resto `C_IVA_VENTAS_CF` | 7795, 9004 | CONFIRMADO |
| Letra A para cond. 1 y 3; B para el resto | ~7355 | CONFIRMADO |
| Sucursales 37 y 48 no cobran IVA ni percepción IVA | 7882, 8995, 11490 | CONFIRMADO (motivo NO DET.) |
| Cliente CF (4) no paga percepción, Decreto 1344, TISH; sin CUIT | `oCbo_iva_Change` | CONFIRMADO |
| Percepción IVA = 0 si agente de retención IVA, o menor al mínimo | ~11463-11475 | CONFIRMADO |
| Si no hay fila de percepción para la cond. IVA → mínimo 9.999.999 (no percibe) y coef. 0 | `ModuloImpuestosEspeciales.bas` | CONFIRMADO |
| IVA se calcula después de descuentos | Neto1 = Σ − desc | CONFIRMADO |
| Impuesto interno no lleva IVA en cartelería | fórmula cartel | CONFIRMADO (otra fórmula en lista prov.) |
| Notas de crédito de envase vacío no devuelven IIBB (Cba/Ctes) | ~11308-11318 | CONFIRMADO |
| Connexa: todo impuesto ≠ IVA se trata como interno | `OC_CREATE` | CONFIRMADO |
| Devoluciones/anulaciones usan config. original | grabación en línea (`K_IVA`) | INFERIDO |
| Tabla T110_IVA se mantiene fuera del sistema | sin escrituras en repo | INFERIDO |
| Significado de códigos 3 y 6 de condición IVA | — | NO DETERMINADO |

---

## PARTE VIII — DUDAS A VALIDAR
1. Valores productivos de `T110_IVA` (códigos 1–4, `K_COEF_IVA`) y de `T001` tabla 12.
2. ¿Quién/cómo se da de alta o modifica `T110_IVA`? ¿Existen triggers/auditoría?
3. ¿`MultipleAlicuota` está en 1 en producción? (define si prevalece el ini sobre la tabla).
4. Significado de `C_COSTO_IVA` en la tabla y su relación con `C_IVA_EN_CALCULO` (=2 en Connexa).
5. Por qué las sucursales 37 y 48 no cobran IVA (zona franca/régimen especial) y si hay otras.
6. DDL: tipos/PK/FK de `T110_IVA`, `T110_IVA_PERCEPCION`, `T050_ARTICULOS`; SP de grabación de comprobantes y libros IVA (`K_IVA`, `I_IVA*`).
7. Cómo Libro IVA calcula períodos históricos tras cambio de alícuota.
8. Códigos de tipo de IVA de la grilla (1/2/4) frente a `C_IVA` (¿3 = 27 %?).
9. Origen del GUID de IVA 21 % en Connexa y manejo de IVA 10,5 %/27 % en OC enviadas.

---

## PARTE IX — RIESGOS Y HALLAZGOS

**Problemas confirmados**
- Dos fuentes de alícuota: `T110_IVA` y `seguridad.ini` (`GblF_CalculoIVA` en EDI y Pago Prov. versus `GblF_CoefIVA` del maestro).
- `GblF_CoefIVA` redondea a 3 decimales; `GblF_CalculoIVA` redondea a 3; comprobantes de compra a 2; ventas a 3/4 → diferencias de centavos posibles entre línea y comprobante.
- Hardcodes: sucursales 37/48; sucursal 13 (tasa ambiental); GUID IVA 21 en `SP_CONNEXA_PO_SEND_ONE`; provincias numéricas (1, 5, 6, 18, 19, 24…) en `LclF_GetCoefIb`; tramos de fechas en `GblF_CorrijoIvaPercepcion1`; compras con campos fijos 21/10,5/27.
- Fórmula de impuesto interno distinta entre cartelería, lista de proveedor y OC Connexa (`I_IMP_INTERNO = precio × rate`, sin IVA).
- Connexa: `K_COEF_IVA` de la OC puede diferir de `T110_IVA` consultado luego (`OC_GET` usa `art.C_IVA`).
- Fallback silencioso a 0 en `LcLF_GetCoefIva` (código inexistente = sin IVA).
- N+1: `GblF_CoefIVA`/`LcLF_GetCoefIva` abren una conexión por llamada.
- SQL dinámico por concatenación (`"... WHERE C_IVA = " & CIva`).
- Código duplicado en múltiples copias de `Facturador_Articulos` (`Luke`, `- copia`, `Copia de`, `Copia (2) de`, Clientes/Sucursal).

**Riesgos potenciales**
- Cambio de alícuota sin vigencia: contamina reportes/recálculos que relean `T110_IVA`.
- Percepción IVA: bloque cond.5 se sobrescribe luego por el bloque RI/excluido (código muerto o error; verificar).
- Alícuota 27 % presente en compras pero la venta solo maneja dos grupos (Neto1/Neto2) + exento.

**Observaciones**
- No hay endpoints API con lógica impositiva propia; la integración fiscal AFIP/CAE y SAP se menciona en documentación existente pero su mapeo de impuestos no es visible en el código analizado (NO DETERMINADO).
