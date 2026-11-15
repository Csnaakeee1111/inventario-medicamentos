# Inventario de Medicamentos

Programa de escritorio (Windows/Linux) para controlar medicamentos por lote, con
tabla semanal de lunes a viernes, kardex, cuadres físicos y exportación a Excel/PDF.

## Pestañas
- **Semana (Lun–Vie):** entradas (verde) y salidas (rojo) por día y lote. Cambia sola cada semana.
- **Inventario:** Medicamento · Descripción · Existencias físicas · Kardex · Diferencia · Lote · Vencimiento.
- **Kardex:** historial de cada medicamento con saldo.
- **Cuadre:** conteo físico semanal o mensual contra el kardex; hoja para contar en PDF; ajuste opcional.
- **Vencimientos:** lotes que vencen pronto (rojo vencido, naranja ≤30 días, amarillo ≤90).

## Dónde quedan los datos
- Windows: `%APPDATA%\InventarioMedicamentos\inventario.db` (+ `respaldos\`, uno por día, 30 días).
- Linux: `~/.local/share/InventarioMedicamentos/`.
- `Documentos\Inventario Medicamentos\Inventario_actual.xlsx` se mantiene al día y **se puede editar**:
  al guardarlo, el programa toma los cambios en ~5 s. Lo cambiado se pinta morado (programa y Excel);
  clic en la casilla = quién, cuándo, antes y después. Si el mismo dato cambió también en el programa,
  no se aplica y se avisa. Cambiar «Kardex» crea un ajuste; «Existencias físicas», un conteo.
  Filas nuevas abajo = medicamento/lote nuevo. Historial: hoja «Cambios» y menú Archivo.
- Actualizar el .exe NO toca los datos.

## Probar en Linux
    ./probar.sh          # usa una base de prueba aparte
    QT_QPA_PLATFORM=offscreen .venv/bin/python pruebas/prueba_movimientos.py   # y prueba_pantallas, prueba_excel

## Sacar una versión nueva (puente de actualización)
1. Cambiar el código.
2. `git commit -am "lo que cambió"`
3. `git tag v1.0.1 && git push && git push --tags`
4. GitHub Actions arma `InventarioMedicamentos.exe` y lo publica en Releases (~5 min).
5. En la máquina Windows: Ayuda → Buscar actualizaciones (también revisa solo al abrir).

## Primera instalación en Windows
Descargar `InventarioMedicamentos.exe` de Releases, ponerlo en una carpeta fija
(ej. `C:\Inventario\`) y crear acceso directo en el escritorio.
Windows SmartScreen puede avisar la primera vez: «Más información» → «Ejecutar de todas formas».
