# saldos-farmacias-iva-iibb

Arma los saldos de IVA e IIBB del mes de cada sociedad para la planilla
`FARMACIAS SALDO MM-AAAA` del SharePoint (columnas C a I), leyendo las DDJJ de la unidad
compartida: F.2051 de ARCA, R-606M de ARBA, SICOL de AGIP y CM03 de Convenio Multilateral.
Controla cada DJ, avisa las que faltan y deja las DDJJ listas para subir a la carpeta
`DDJJ IVA-IIBB MM-AAAA` del SharePoint. Solo lee la unidad: no entra a ARCA ni sube nada.

## Instalación

En las máquinas del equipo lo instala `setup-bva`. A mano:

```
/plugin install saldos-farmacias-iva-iibb --marketplace estudiobva/estudio-claude-marketplace
```

```bash
python3 -m pip install -r requirements.txt
cp .env.example .env   # opcional: ruta de la unidad compartida y de salida
```

## Uso

```bash
python3 scripts/saldos-iva-iibb.py --periodo=MM/AAAA
python3 scripts/saldos-iva-iibb.py --periodo=MM/AAAA --valores=valores.json --preparar-subida
```

Deja `SALDOS IVA IIBB MM-AAAA.xlsx` en `~/Documents/BVA-salidas/saldos-iva-iibb/` y un JSON
por stdout. Las DJ que son imagen (y las de AGIP y CM03) salen como `pendientes` con su PNG:
se leen a mano, se cargan en `valores.json` y se vuelve a correr. El formato está en el
encabezado del script y en el SKILL.md.

## Reglas

- IVA: C saldo técnico a favor del contribuyente, D SLD del período, E = min(SLD anterior,
  saldo técnico ARCA − retenciones), F saldo a favor de ARCA. Control: F − D = saldo técnico
  ARCA − SLD anterior − retenciones.
- ARBA: acumulado a favor → G, a pagar → I. AGIP: H a favor, I a pagar. CM03: G Buenos
  Aires, H CABA, I total a pagar.
- Jurisdicción donde no presenta → 0; DJ que falta en la unidad → vacío.
