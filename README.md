# Control de Planos

Programa de escritorio (Windows) para registrar a quién se le entrega cada plano y cuándo se devuelve.

## Datos que guarda

| Campo | Cómo se carga |
|---|---|
| Código de pieza | Se escribe |
| Contiene muestra | Sí / No |
| Nombre de la pieza | Se escribe |
| Sector al que se dirige | Se escribe o se elige de los ya usados |
| Operador que pidió el plano | Se elige de la lista de operadores (botón «…» para cargarlos) |
| Fecha de entrega | Automática al presionar **Registrar entrega** |
| Fecha de devolución | Automática al presionar **Marcar como DEVUELTO** (o doble clic en la fila) |

Todo se guarda en una base de datos SQLite (`control_planos.db`) que se crea junto al `.exe`.
Para hacer una copia de seguridad basta con copiar ese archivo.

## Uso

1. Cargar los operadores con el botón **…** al lado de «Operador».
2. Completar el formulario y presionar **Registrar entrega** (o Enter).
3. Cuando vuelve el plano: seleccionar la fila y presionar **Marcar como DEVUELTO**.
4. **Exportar a Excel** genera un `.xlsx` con lo que se está mostrando (respeta la búsqueda y el filtro Todos / Pendientes / Devueltos).

Las filas amarillas son planos sin devolver; las verdes, devueltos.

## Descargar el ejecutable

En la pestaña **Releases** del repositorio → *Control de Planos - última versión* → `ControlPlanos.exe`.
Se recompila automáticamente (GitHub Actions) cada vez que se modifica el código.

Al abrirlo por primera vez Windows puede mostrar «Windows protegió su PC»: elegir **Más información → Ejecutar de todas formas** (el ejecutable no está firmado digitalmente).

## Ejecutar desde el código

```
pip install openpyxl
python ControlPlanos/control_planos.py
```
