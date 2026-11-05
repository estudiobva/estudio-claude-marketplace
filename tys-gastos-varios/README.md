# tys-gastos-varios

Baja de **Zetti (T&S Web)** el **Reporte Seguimiento de Valores (4.5.2)** filtrado
por el tipo de valor **56 GAVA - GASTOS VARIOS**, lo parte por mes, lo cruza con
el **EERR** de la sociedad (`0 - ESTUDIO/EERR 2026`) y guarda cada mes en la
carpeta del mes de la sociedad:

```
[Sociedad]/01-Impuestos Mensuales/[AAAA/]AAAAMM/AAAAMM - ZETTI - GASTOS VARIOS.xlsx
```
 Reemplaza a la
habilidad de Cowork `tys-gastos-varios`, que se hacía a mano con Claude in Chrome.

El login es automático: usuario y clave de Zetti en `Claves_Organismos.xlsx`
(grupo de columnas ZETTI, fila "Zetti T&S Web (todas las sociedades)"). La
sociedad se busca por CUIT y se usa la entidad Sociedad, que abarca todos sus
locales. Solo lee: no cambia nada en Zetti ni en el EERR.

## Instalacion

```bash
npm install
npx playwright install chromium
python3 -m pip install -r requirements.txt
cp .env.example .env   # o usar ~/.fisco-ar/.env, el mismo de los demas plugins
```

## Uso

```bash
node scripts/tys-gastos-varios.js --nombre="<sociedad>"            # ejercicio hasta el mes anterior
node scripts/tys-gastos-varios.js --nombre="<sociedad>" --periodo=10/2026
node scripts/tys-gastos-varios.js --todas --periodo=10/2026
node scripts/tys-gastos-varios.js --listar
```

Cada archivo trae las hojas Resumen, Detalle y Cruce EERR del mes. Detalle en
`skills/tys-gastos-varios/SKILL.md`.
