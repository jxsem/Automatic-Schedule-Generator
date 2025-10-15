Generador de Horarios Automático – Google Sheets

Descripción:
Este proyecto consiste en un sistema automático de generación de horarios para empleados, implementado en JavaScript usando Google Apps Script en Google Sheets. Permite organizar turnos de manera eficiente respetando:
- Horas máximas por empleado.
- Descansos fijos y específicos.
- Turnos de mañana y tarde.
- Dobles turnos según necesidades.
- Ajustes personalizados por empleado (ejemplo: horarios fijos de María Ruiz).
- Evita duplicación de horas en el cálculo total.

Tecnologías utilizadas:
- Google Apps Script – Automatización de Google Sheets.
- JavaScript – Lógica de asignación de turnos y cálculos de horas.
- Google Sheets – Interfaz para visualizar y ajustar horarios.

Funcionalidades clave:
1. Asignación de turnos automática
   - Turnos de mañana y tarde.
   - Turnos dobles donde aplica.
   - Exclusión de empleados en días de descanso.

2. Horarios personalizados
   - María Ruiz tiene horarios fijos de lunes a sábado, con ajuste para sumar exactamente 24h.
   - Ajustes de salida y entrada personalizadas (ej. martes 9:00–13:00).

3. Cálculo de horas trabajadas
   - Se calcula la suma total de horas por empleado.
   - Evita duplicaciones y respeta máximos semanales.

4. Compatibilidad con Google Sheets
   - Actualización automática de celdas con ajuste de texto.
   - Compatible con columnas de días y columna de total de horas.

Estructura del código:
- generarHorario() – Función principal que gestiona todo el flujo:
  - Lectura de empleados desde Google Sheets.
  - Configuración de turnos y descansos.
  - Asignación de turnos dobles y roles.
  - Actualización de la hoja y cálculo de horas.

- Configuraciones:
  - maxHoras → Horas máximas por empleado.
  - descansos → Días de descanso fijos.
  - asignacionTurnos → Turnos fijos de mañana/tarde.
  - dobleAsignacion → Días con turnos dobles.
  - horasMaria → Ajuste especial para horas exactas de María Ruiz.

Ejemplo de uso:
1. Crear una hoja de Google Sheets con pestañas: Empleados y Horario.
2. Llenar columna A con los nombres de empleados.
3. Copiar y pegar el script en Editor de Apps Script asociado al Sheet.
4. Ejecutar generarHorario() para actualizar los horarios automáticamente.
5. Ver columna I para el total de horas por empleado.

Beneficios:
- Automatiza la creación de horarios semanales.
- Reduce errores humanos y duplicación de horas.
- Permite personalizar reglas por empleado.
- Ideal para empresas con múltiples empleados y turnos variables.
