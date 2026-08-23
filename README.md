# Portal de Transparencia — Tau Clinical Research

Portal público de consulta para que los alumnos vean el estado editorial de sus manuscritos.
El alumno escribe su nombre y obtiene todos los trabajos en los que participa, con el estado,
la revista, el cuartil, el enlace y las observaciones del equipo.

**Web:** https://jbarbozameca.github.io/publicacionestau/

## Estructura

```
index.html                     ← el portal completo (autocontenido: datos + logo + estilos)
tools/build_data.py            ← convierte el Excel en data.json
tools/build_portal.py          ← ensambla plantilla + datos + logo → index.html
tools/portal_template.html     ← plantilla (diseño y lógica, sin datos)
tools/mark.png                 ← logo Tau
data/TRANSPARENCIA - INVESTIGACIONES Tau 2024-2026.xlsx  ← fuente de datos
```

`index.html` es un único archivo sin dependencias externas (salvo las tipografías de Google Fonts),
así que funciona en GitHub Pages sin configuración adicional.

## Actualizar el portal (cada lunes)

1. Reemplaza el Excel en `data/` por la versión actualizada.
2. Ejecuta:

   ```bash
   pip install openpyxl
   python3 tools/build_portal.py "data/TRANSPARENCIA - INVESTIGACIONES  Tau 2024-2026.xlsx"
   cp portal_tau.html index.html
   ```

3. `git add -A && git commit -m "Actualización semanal" && git push`

La fecha de corte que se muestra en el portal se toma del día en que se ejecuta el script.

## Origen de los datos

Las tres hojas del Excel (`ACEPTADOS PUBLICADOS`, `TRAINING`, `FOCUS`) se leen con estas reglas:

- **Grupo**: se toma de la columna A cuando contiene una fecha (mes de inicio) o una etiqueta
  (`Focus Mayo`, `training 2025`, `grupo 1 UNAC`, `PROTOCOL`). Si la celda solo trae un número
  correlativo o está vacía, el trabajo se muestra como *Sin grupo asignado*.
- **Autores**: se separan por saltos de línea, `;`, `,` y `and`, quitando números de afiliación
  y grados académicos. La búsqueda usa el texto completo de la celda, así que encuentra al autor
  aunque su nombre esté escrito con variantes.
- **Estados**: se normalizan a `Publicado`, `Aceptado`, `En peer review (R1)`, `Enviado`,
  `En proceso`, `En redacción` y `Rechazado`.
