# Esquema de las 10 listas SharePoint

Cada lista corresponde a una colección de datos del ERP. Las columnas marcadas con * son **obligatorias**.

> 💡 **Tip**: Todas las columnas opcionales se marcan como `NotRequired` para que el script PowerShell las cree correctamente.

---

## 1. Clientes

| Columna | Tipo | Notas |
|---|---|---|
| Title* | Texto (256) | Razón social del cliente |
| Ruc | Texto (11) | RUC del cliente |
| Corto | Texto (10) | Código corto (ej: CMP) |
| Sector | Texto (100) | Sector/Rubro |
| Direccion | Texto (255) | |
| Contactos | Multilínea (JSON) | Array de contactos en JSON |
| Areas | Multilínea (JSON) | Array de áreas en JSON |
| LogoBase64 | Multilínea | Imagen del logo en base64 (opcional) |
| Notas | Multilínea | |

---

## 2. Proveedores

| Columna | Tipo | Notas |
|---|---|---|
| Title* | Texto (256) | Razón social del proveedor |
| Ruc | Texto (11) | |
| Rubro | Texto (100) | |
| Contacto | Texto (200) | |
| Telefono | Texto (50) | |
| Email | Texto (100) | |
| Direccion | Texto (255) | |
| Web | Texto (200) | |
| Cuenta | Texto (200) | |
| DiasCredito | Número | |
| FormaPagoDefault | Texto (100) | |
| Calificacion | Texto (10) | A/B/C/D |
| Contactos | Multilínea (JSON) | |
| Notas | Multilínea | |

---

## 3. Cargos

| Columna | Tipo | Notas |
|---|---|---|
| Title* | Texto (256) | Nombre del cargo |
| Categoria | Texto (100) | Supervisión/Ingeniería/Técnico/etc. |
| CostoRealHH | Número (decimal) | Costo real S/ por hora |
| Ganancia | Número (decimal) | Multiplicador de ganancia |
| PrecioHH | Número (decimal) | Precio venta calculado |
| UsaHerramientas | Sí/No | |
| UsaSoftware | Sí/No | |
| UsaEquipos | Sí/No | |
| Notas | Multilínea | |

---

## 4. Personal

| Columna | Tipo | Notas |
|---|---|---|
| Title* | Texto (256) | Nombre completo (campo combinado) |
| Nombres | Texto (200) | |
| Apellidos | Texto (200) | |
| Dni | Texto (15) | |
| Telefono | Texto (50) | |
| Email | Texto (100) | |
| CargoIdSP | Texto (50) | ID de la lista Cargos |
| FechaIngreso | Fecha | |
| Activo | Sí/No | |
| Especialidad | Texto (255) | |
| Notas | Multilínea | |

---

## 5. Gestion (Personal Administrativo para OCs)

| Columna | Tipo | Notas |
|---|---|---|
| Title* | Texto (256) | Nombre completo |
| Nombres | Texto (200) | |
| Apellidos | Texto (200) | |
| Cargo | Texto (100) | Gerente, Gestor Compras, etc. |
| Area | Texto (100) | Gerencia/Logística/Comercial/etc. |
| Email | Texto (100) | |
| Telefono | Texto (50) | |
| RolGestor | Sí/No | |
| RolSolicitante | Sí/No | |
| RolAprobador | Sí/No | |
| RolFirma | Sí/No | |
| Notas | Multilínea | |

---

## 6. Materiales

| Columna | Tipo | Notas |
|---|---|---|
| Title* | Texto (256) | Descripción del material |
| Marca | Texto (100) | |
| Modelo | Texto (200) | |
| Fabricante | Texto (200) | |
| Um | Texto (20) | Unidad de medida |
| CostoReal | Número (decimal) | Costo S/ |
| Ganancia | Número (decimal) | Multiplicador |
| Logistica | Número (decimal) | Multiplicador logística |
| Categoria | Texto (100) | |
| Notas | Multilínea | |

---

## 7. Cotizaciones

| Columna | Tipo | Notas |
|---|---|---|
| Title* | Texto (256) | Código de propuesta (ej: CI26-CMP-OE-001) |
| Tipo | Texto (20) | "servicios" o "suministro" |
| Estado | Texto (20) | borrador/enviada/aprobada/rechazada/facturada/pagada |
| Version | Número | |
| Fecha | Fecha | |
| Servicio | Multilínea | |
| Lugar | Texto (255) | |
| Plazo | Texto (100) | |
| ClienteIdSP | Texto (50) | ID en lista Clientes |
| ClienteRazon | Texto (255) | snapshot |
| ContactoClienteId | Texto (50) | |
| Atencion | Texto (255) | |
| Cargo | Texto (255) | |
| Contacto | Texto (255) | |
| ContactoConsitec | Texto (255) | |
| LogoCliente | Multilínea | Base64 del logo cliente |
| Alcance1 | Multilínea | |
| Alcance2 | Multilínea | |
| Mo1Json | Multilínea | JSON de mano de obra sección 1 |
| Mov1Json | Multilínea | JSON movilización sección 1 |
| Eq1Json | Multilínea | JSON equipos sección 1 |
| Mo2Json | Multilínea | JSON mano de obra sección 2 |
| Mov2Json | Multilínea | JSON movilización sección 2 |
| Eq2Json | Multilínea | JSON equipos sección 2 |
| MatJson | Multilínea | JSON materiales |
| Gg | Número | Gastos generales % |
| Util | Número | Utilidad % |
| Modalidad | Texto (100) | |
| Validez | Texto (50) | |
| FormaPago | Texto (100) | |
| MetodoPago | Texto (100) | |
| NotasJson | Multilínea | |
| ExclJson | Multilínea | |
| RespJson | Multilínea | |
| FichasJson | Multilínea | Para suministros |
| TotalFinal | Número (decimal) | Total cotización |
| Ganancia | Número (decimal) | Ganancia calculada |
| ProyectoIdSP | Texto (50) | Vinculación a proyecto |

---

## 8. OrdenesCompra

| Columna | Tipo | Notas |
|---|---|---|
| Title* | Texto (256) | Número de OC (ej: 202622042) |
| Fecha | Fecha | |
| FechaEntrega | Fecha | |
| FechaVencimiento | Fecha | |
| CentroCompra | Texto (100) | |
| GestorIdSP | Texto (50) | |
| GestorNombre | Texto (200) | |
| GestorEmail | Texto (100) | |
| SolicitanteIdSP | Texto (50) | |
| SolicitanteNombre | Texto (200) | |
| SolicitanteEmail | Texto (100) | |
| AprobadorIdSP | Texto (50) | |
| AprobadorNombre | Texto (200) | |
| AprobadorEmail | Texto (100) | |
| ProveedorIdSP | Texto (50) | |
| ProveedorRazon | Texto (255) | |
| ProveedorRuc | Texto (11) | |
| ProveedorDir | Texto (255) | |
| ProveedorContacto | Texto (200) | |
| ProveedorTel | Texto (50) | |
| ProveedorEmail | Texto (100) | |
| ProyectoIdSP | Texto (50) | |
| ProyectoNombre | Texto (255) | |
| ProyectoTexto | Texto (500) | |
| Almacen | Texto (100) | |
| Pedido | Texto (100) | |
| NroCotizacion | Texto (50) | |
| TratadoCon | Texto (200) | |
| FechaEntregaText | Texto (200) | |
| LugarEntrega | Multilínea | |
| FormaPago | Texto (100) | |
| Moneda | Texto (10) | PEN/USD |
| ItemsJson | Multilínea | JSON de items |
| Igv | Número (decimal) | |
| Descuento | Número (decimal) | |
| Notas | Multilínea | |
| PagosJson | Multilínea | JSON de pagos |
| Estado | Texto (20) | pendiente/pagada_parcial/pagada/vencida |
| Total | Número (decimal) | |
| Pagado | Número (decimal) | |
| Saldo | Número (decimal) | |

---

## 9. Proyectos

| Columna | Tipo | Notas |
|---|---|---|
| Title* | Texto (256) | Nombre del proyecto |
| Codigo | Texto (50) | PRY-26-001 |
| ClienteIdSP | Texto (50) | |
| ClienteRazon | Texto (255) | |
| Lugar | Texto (255) | |
| CotizacionIdSP | Texto (50) | |
| CotizacionCodigo | Texto (50) | |
| CotizacionTipo | Texto (20) | |
| Estado | Texto (20) | planificado/en_curso/completado/pausado/cancelado |
| FechaInicio | Fecha | |
| FechaFin | Fecha | |
| Avance | Número | 0-100 |
| Monto | Número (decimal) | |
| Responsable | Texto (200) | |
| TareasJson | Multilínea | JSON del cronograma Gantt |
| BaselineJson | Multilínea | JSON línea base |
| CalendarioJson | Multilínea | JSON calendario laboral |
| Notas | Multilínea | |

---

## 10. Reportes

| Columna | Tipo | Notas |
|---|---|---|
| Title* | Texto (256) | Código del reporte (RD-001 o RS-001) |
| Tipo | Texto (10) | "diario" o "semanal" |
| ProyectoIdSP | Texto (50) | |
| Fecha | Fecha | |
| FechaIni | Fecha | (solo semanal) |
| FechaFin | Fecha | (solo semanal) |
| ResponsableIdSP | Texto (50) | |
| ResponsableNombre | Texto (200) | |
| Clima | Texto (50) | |
| Turno | Texto (50) | |
| ActividadesJson | Multilínea | |
| PersonalJson | Multilínea | Asistencia |
| FotosJson | Multilínea | Base64 de fotos |
| Incidentes | Multilínea | |
| Observaciones | Multilínea | |
| AvancePorc | Número | |
| Notas | Multilínea | |

---

## ⚠ Importante sobre los JSON

Para datos complejos (arrays de items, contactos múltiples, etc.) usamos **columnas Multilínea** que contienen JSON serializado. La capa de abstracción `storage-adapter.js` se encarga de serializar/deserializar automáticamente.

Esto es porque SharePoint no es una base de datos relacional y las "Listas" no soportan estructuras anidadas. Es el patrón estándar para este tipo de migración.
