# estudio-claude-marketplace

Marketplace de plugins de Claude Code del Estudio BVA. Cada carpeta es un plugin:
un script de Playwright y la skill que lo usa.

| Plugin | Qué hace |
|---|---|
| `setup-bva` | Prepara la computadora: programas, marketplace, dependencias, `.env` y diagnóstico. Avisa al abrir Claude si falta algo. |
| `planes-facilidades-arca` | Releva los planes de facilidades de pago en Mis Facilidades (ARCA). |
| `conciliacion-ncr` | Concilia el Portal IVA contra las NCR de Suizo y Monroe y las saca del Libro Compras. |
| `anticipos-sct` | Releva los anticipos de Ganancias y Bienes Personales en el SCT (ARCA), con capital e intereses. |
| `deudas-ddjj-sct` | Controla en el SCT (ARCA) las deudas que no son anticipos y las DDJJ pendientes de presentación, en un Excel consolidado. |
| `f931-descarga-arca` | Baja los F.931 presentados desde Declaración en Línea (ARCA), un PDF por período, con login automático desde la planilla de claves. |
| `bancos-bva` | Descarga extractos bancarios del home banking de empresas (9 bancos), los renombra y los archiva por sociedad, banco y año. |
| `portal-iva-descarga` | Baja del Portal IVA (ARCA) el Libro IVA Compras y Ventas del período, en Excel, y los archiva en la carpeta IVA del mes. |
| `facturacion-arca` | Emite facturas de servicios en Comprobantes en Línea (ARCA) para cualquier emisor desde la planilla mensual: validar, dry-run y emisión con código. **Escribe en ARCA.** |

## Instalación

**En Windows (el equipo):** doble clic en `Instalar Claude BVA.cmd`, en la unidad
compartida > `0 - ESTUDIO - 00000000000/automatizacion_bva/setup/`. Instala los
programas, el marketplace, todos los plugins y sus dependencias, y arma el `.env`.
Requisitos y paso a paso en [setup-bva/README.md](setup-bva/README.md).

**A mano (Mac o para desarrollar):** el repo es privado, hace falta acceso a
`estudiobva/estudio-claude-marketplace` y estar logueado en GitHub (`gh auth login`).

```
/plugin marketplace add estudiobva/estudio-claude-marketplace
/plugin install setup-bva@estudiobva-skills
/plugin install <plugin>@estudiobva-skills        # uno por cada plugin de la tabla
```

Después, y después de cada actualización: `node scripts/dependencias.js` y
`node scripts/configurar.js` desde la carpeta de `setup-bva` (o pedirle a Claude
"revisá el setup del estudio"). Las credenciales van en `~/.fisco-ar/.env`, fuera
del plugin, así que las actualizaciones no lo tocan.

Para actualizar: `/plugin marketplace update estudiobva-skills` (con `setup-bva`
queda en automático).

## Agregar un plugin

1. Carpeta nueva con `.claude-plugin/plugin.json`, `skills/<nombre>/SKILL.md`,
   `scripts/`, `lib/`, `package.json`, `.env.example` y `README.md`.
2. Sumarlo a `plugins` en `.claude-plugin/marketplace.json` con
   `"source": "./<carpeta>"`.
3. En la skill, correr los scripts con `cd "${CLAUDE_PLUGIN_ROOT:-.}" && ...`.
4. Si necesita otra dependencia de npm o Python, declararla en `package.json` /
   `requirements.txt`: `setup-bva` las instala solo con eso.
5. Subir el cambio en una rama propia y mergear por PR (nunca directo a `main`).
