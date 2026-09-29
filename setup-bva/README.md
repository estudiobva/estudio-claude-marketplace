# setup-bva

Prepara una computadora del Estudio BVA (Windows 10/11) para usar los plugins de
Claude del marketplace `estudiobva-skills`, tal cual se desarrollaron.

## Para el equipo: instalar en tu computadora

**Antes (una sola vez):**

1. Tené Google Drive para escritorio abierto con tu cuenta `@estudiobva.com` y la
   unidad compartida *BVA - Sociedades Farmaceuticas* visible en el Explorador.
2. Creá una cuenta en [github.com](https://github.com) (gratis) y pasale tu usuario a
   Joaquín para que te invite a la organización `estudiobva`. Aceptá la invitación
   que llega por mail.
3. Pedile a Joaquín la contraseña de las planillas de claves del estudio.

**Instalar:**

1. En la unidad compartida, abrí `0 - ESTUDIO - 00000000000 > automatizacion_bva > setup`.
2. Doble clic en **`Instalar Claude BVA.cmd`**.
3. Seguí la ventana. Te va a pedir:
   - permiso de Windows para instalar programas (aceptá);
   - entrar a GitHub con el navegador (copiás un código de 8 letras y tocás *Authorize*);
   - la contraseña de la planilla (no se ve al escribir).
4. Cuando diga **Listo**, cerrá la app de Claude del todo (también desde el ícono junto
   al reloj) y volvela a abrir. Los plugins aparecen en la pestaña **Code**.

Se puede correr las veces que haga falta: lo que ya está hecho lo saltea. Si algo
falla, la ventana dice qué hacer y deja el registro en
`Documentos\BVA-salidas\setup-bva.log` (no tiene contraseñas).

**Después:** los plugins se actualizan solos al abrir Claude. Si alguno deja de
andar, pedile a Claude *"revisá el setup del estudio"*.

## Qué instala

| Paso | Qué hace |
|---|---|
| 1. Programas | Git, Node.js LTS, Python 3.12 y GitHub CLI con `winget`; Claude Code con el instalador oficial; la app de Claude si falta (pregunta). Lo que ya está no se toca. |
| 2. GitHub | `gh auth login --web` y `gh auth setup-git`, y prueba el acceso al repo privado. |
| 3. Marketplace | Agrega `estudiobva/estudio-claude-marketplace` (o lo actualiza) e instala **todos** los plugins que publica, a nivel usuario. |
| 4. Dependencias | `npm install`, Chromium de Playwright y `pip install --user -r requirements.txt` de cada plugin (`scripts/dependencias.js`). |
| 5. Configuración | Arma `%USERPROFILE%\.fisco-ar\.env` con la unidad compartida, la planilla de claves, la contraseña y el Python (`scripts/configurar.js`, deja copia `.env.bak-*` si ya existía). Activa la actualización automática del marketplace y guarda `BVA_UNIDAD_PATH` y `BVA_PYTHON` como variables de usuario. |
| 6. Diagnóstico | `scripts/diagnostico.js`: todo en OK / AVISO / ERROR. |

Además, el plugin trae un hook de inicio: cada vez que se abre Claude revisa (en
menos de un segundo, sin instalar nada) si un plugin actualizado quedó sin
dependencias, si falta el `.env` o si hay plugins nuevos, y en ese caso le avisa a
Claude qué hacer. La skill `setup-bva` hace los arreglos desde el chat.

## Contraseñas

- Las escribe la persona en la ventana del instalador (entrada oculta). Pasan a
  `configurar.js` por variable de entorno, se prueban abriendo la planilla y solo se
  guardan en el `.env` si abre. Nunca van por la línea de comandos ni al registro.
- Claude nunca las pide ni las ve: si falta una, la skill manda a correr el instalador.

## Para Joaquín: publicar el instalador

El instalador que usa el equipo es la copia de la unidad compartida. Después de
cambiar `instalar.ps1` o `Instalar Claude BVA.cmd`, copiar los dos a
`0 - ESTUDIO - 00000000000/automatizacion_bva/setup/`. El resto (dependencias,
configuración, diagnóstico) vive en el plugin y se actualiza con el marketplace.

Para dar de alta a alguien: invitarlo a la organización `estudiobva` en GitHub con
acceso de lectura al repo, y pasarle la contraseña de la planilla.

Opciones del instalador:

```
powershell -ExecutionPolicy Bypass -File instalar.ps1 -SinProgramas
powershell -ExecutionPolicy Bypass -File instalar.ps1 -Unidad "G:\Unidades compartidas\BVA - Sociedades Farmaceuticas"
```
