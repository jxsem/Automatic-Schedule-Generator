# Generador de Horarios

Script de **Google Apps Script** que genera automáticamente el horario semanal de los empleados en una hoja de Google Sheets y calcula las horas totales de cada uno.

Autor: Jose Manuel Soldado J.

## Qué hace

- Pide por pantalla los datos de cada empleado (nombre, horas, descansos, dobles y turno fijo).
- Rellena la hoja `Empleados` con el turno de cada día, de lunes a sábado.
- Calcula las horas semanales reales y avisa si no coinciden con las horas indicadas.

## Contenido del repositorio

| Archivo | Descripción |
| --- | --- |
| [script-horario.js](script-horario.js) | Código del script (función principal `generarHorario`). |
| `Copia de Creación de Turnos y roles de la semana. CC097.xlsx` | Hoja de ejemplo de turnos y roles. |

## Requisitos

- Una hoja de cálculo de Google Sheets con una pestaña llamada **`Empleados`**.
- La fila 1 con estas cabeceras:

  | Nombre | Lunes | Martes | Miércoles | Jueves | Viernes | Sábado | Horas |
  | --- | --- | --- | --- | --- | --- | --- | --- |

- El script pegado en **Extensiones → Apps Script**.

## Cómo usarlo

1. Abre la hoja y entra en **Extensiones → Apps Script**.
2. Pega el contenido de `script-horario.js`.
3. Selecciona la función `generarHorario` y pulsa **Ejecutar**.
4. Responde a las preguntas que aparecen, en el orden que se describe a continuación.

### 1. Datos de cada empleado

Se repite para cada empleado. Pulsa **Cancelar** en la pregunta del nombre cuando hayas terminado.

| Pregunta | Qué responder |
| --- | --- |
| Nombre | Formato obligatorio: `Apellido2, Apellido1, Nombre`. Ejemplo: `García, López, Ana` se guarda como `Ana López García`. |
| Horas semanales | Número entero entre 1 y 80. |
| Días de descanso | `0`, `1` o `2`. Si es más de 0, indica después qué día(s), separados por coma. |
| Días que dobla | Entre 0 y los días laborables disponibles (6 menos los de descanso). |
| Turno fijo | `mañana`, `tarde` o `ninguno`. |

### 2. Configuración del horario

Se pregunta una sola vez y es común a todos los empleados.

- Hora de apertura y de cierre de la mañana, en formato `HH:MM` (por ejemplo `10:00` y `14:00`).
- Si hay horario de tarde (`si` / `no`). Si lo hay, hora de inicio y de fin de la tarde.

## Lógica de asignación

1. Los **días de descanso** indicados se respetan siempre.
2. Los **turnos dobles** (mañana y tarde el mismo día) se asignan primero, tantos como se hayan indicado. Se respetan aunque el empleado tenga turno fijo.
3. En el resto de días se asigna el **turno fijo** del empleado. Si no tiene, se elige mañana o tarde al azar.
4. Los días se barajan al azar en cada ejecución, así que el horario cambia cada vez que se genera.

> Si no hay horario de tarde, el turno doble equivale a la mañana.

## Avisos y errores

- **Las horas no cuadran:** si las horas calculadas no coinciden con las indicadas, el script muestra un aviso, pero escribe igualmente el horario.
- **Turno fijo de tarde sin horario de tarde:** el script se detiene con un mensaje de error y no genera nada.
- **Datos no válidos** (nombre sin el formato correcto, número fuera de rango, día o hora incorrectos): se muestra el error y la ejecución se detiene. Hay que volver a ejecutar el script.
- **Cancelar** en cualquier pregunta, salvo la del nombre, detiene la ejecución.

> **Atención:** al generar el horario se borra el contenido anterior de las columnas de empleados, turnos y horas de la hoja `Empleados` (desde la fila 2).

## Ejemplo rápido

Configuración:

- Mañana: `10:00-14:00` (4 h)
- Tarde: `17:00-21:00` (4 h)
- Turno doble: 8 h

Empleado con **32 h semanales**, **1 día de descanso** (5 días laborables) y **4 días doblando**:

- 4 días dobles × 8 h = 32 h
- El quinto día laborable se asigna como turno simple (4 h), así que el total real es **36 h**.
- El script avisa de que no coincide con las 32 h indicadas.

Para que cuadre, deja 3 dobles y 2 simples (3 × 8 + 2 × 4 = 32 h).
