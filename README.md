# skills

Marketplace de plugins de Claude Code del Estudio BVA. Cada carpeta es un plugin:
un script de Playwright y la skill que lo usa.

| Plugin | Qué hace |
|---|---|
| `planes-facilidades-arca` | Releva los planes de facilidades de pago en Mis Facilidades (ARCA). |
| `conciliacion-ncr` | Concilia el Portal IVA contra las NCR de Suizo y Monroe y las saca del Libro Compras. |

## Instalación

El repo es privado: hace falta tener acceso a `estudiobva/skills` y estar logueado en
GitHub (`gh auth login`) en la máquina.

```
/plugin marketplace add estudiobva/skills
/plugin install planes-facilidades-arca@estudiobva-skills
/plugin install conciliacion-ncr@estudiobva-skills
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
