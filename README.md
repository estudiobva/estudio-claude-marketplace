# estudio-claude-marketplace

Marketplace de plugins de Claude Code del Estudio BVA. Cada carpeta es un plugin:
un script de Playwright y la skill que lo usa.

| Plugin | Qué hace |
|---|---|
| `planes-facilidades-arca` | Releva los planes de facilidades de pago en Mis Facilidades (ARCA). |
| `conciliacion-ncr` | Concilia el Portal IVA contra las NCR de Suizo y Monroe y las saca del Libro Compras. |
| `anticipos-sct` | Releva los anticipos de Ganancias y Bienes Personales en el SCT (ARCA), con capital e intereses. |
| `deudas-ddjj-sct` | Controla en el SCT (ARCA) las deudas que no son anticipos y las DDJJ pendientes de presentación, en un Excel consolidado. |
| `bancos-bva` | Descarga extractos bancarios del home banking de empresas (9 bancos), los renombra y los archiva por sociedad, banco y año. |
| `portal-iva-descarga` | Baja del Portal IVA (ARCA) el Libro IVA Compras y Ventas del período, en Excel, y los archiva en la carpeta IVA del mes. |

## Instalación

El repo es privado: hace falta tener acceso a `estudiobva/estudio-claude-marketplace` y estar logueado en
GitHub (`gh auth login`) en la máquina.

```
/plugin marketplace add estudiobva/estudio-claude-marketplace
/plugin install planes-facilidades-arca@estudiobva-skills
/plugin install conciliacion-ncr@estudiobva-skills
/plugin install anticipos-sct@estudiobva-skills
/plugin install deudas-ddjj-sct@estudiobva-skills
/plugin install bancos-bva@estudiobva-skills
/plugin install portal-iva-descarga@estudiobva-skills
```

Después, una vez por plugin (y después de cada actualización), en la carpeta del plugin:

```bash
npm install
npx playwright install chromium
python3 -m pip install -r requirements.txt
```

Las credenciales van en `~/.fisco-ar/.env` (copiar el `.env.example` del plugin). Ese
archivo queda fuera del plugin, así que las actualizaciones no lo tocan.

Para actualizar: `/plugin marketplace update estudiobva-skills`.

## Agregar un plugin

1. Carpeta nueva con `.claude-plugin/plugin.json`, `skills/<nombre>/SKILL.md`,
   `scripts/`, `lib/`, `package.json`, `.env.example` y `README.md`.
2. Sumarlo a `plugins` en `.claude-plugin/marketplace.json` con
   `"source": "./<carpeta>"`.
3. En la skill, correr los scripts con `cd "${CLAUDE_PLUGIN_ROOT:-.}" && ...`.
4. Subir el cambio en una rama propia y mergear por PR (nunca directo a `main`).
