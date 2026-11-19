"""Controla el comparativo de un config de balance: python3 chk_config.py <config.json>. Todo debe dar 0 (±0,02)."""
import json, sys
c = json.load(open(sys.argv[1])); n = c["comparativo"]; er = n["er"]; f = n["flujo"]; nd = n["nota2"]
a2 = n["anexo2"]; s = [round(sum(v[i] for v in a2.values()), 2) for i in range(3)]
for k, v in a2.items():
    if abs(sum(v[:3]) - v[3]) > 0.02: print("ANEXO2 fila no suma", k, v)
print("Anexo2 adm/com/fin vs ER:", round(s[0] + er[3], 2), round(s[1] + er[4], 2), round(s[2] + er[5], 2))
print("ER: GB", round(er[0] + er[1] - er[2], 2), " ordinaria", round(er[2] + er[3] + er[4] + er[5] - er[6], 2),
      " antes IG", round(er[6] + er[7] + er[8] - er[9], 2), " resultado", round(er[9] + er[10] - er[11], 2))
print("Flujo: op", round(sum(f[3:15]) - f[15], 2), " ini+var=cierre", round(f[0] + f[2] - f[1], 2), " suma flujos=var", round(f[15] + f[17] + f[20] - f[21], 2), " var", round(f[2] - f[21], 2))
pi = n["pn_inicio"]; p5 = n["pn_2025"]
print("PN: pn_inicio (cierre 2025) vs pn_2025[3]", round(sum(pi.values()) - p5[3], 2), " evolución", round(sum(p5[:3]) - p5[3], 2), " resultado", round(p5[2] - er[11], 2))
ei, co, ef = n["anexo3"]; print("Anexo3: CMV", round(ei + co - ef + er[1], 2), " EF vs e.1", round(ef - nd.get("e.1", ef), 2))
banc = sum(b[1] or 0 for b in c.get("bancos", [])); print("Caja+bancos (a):", round(nd.get("a.1", 0) + banc, 2), " FCI (b):", round(sum(c.get("fci_2025", {}).values()), 2))
print("d) créditos fiscales:", round(sum(nd.get(f"d.{i}", 0) for i in range(1, 8)), 2), " i) deudas fiscales:", round(sum(nd.get(f"i.{i}", 0) for i in range(1, 5)) + sum(x[1] for x in n.get("cuentas_extra", {}).get("i", [])), 2),
      " j) financieras:", round(sum(x[1] for x in c.get("deudas_bancarias", [])) + nd.get("j.prestamos", 0), 2))
print("(comparar a, b, c.1, d, e.1, f.2, g.1, h.1, i, j con los totales del ESP; activo - pasivo = PN)")
