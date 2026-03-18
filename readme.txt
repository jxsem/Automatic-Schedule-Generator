
GENERADOR DE HORARIOS
Manual de uso — Google Apps Script
Script realizado por Jose Manuel Soldado J.

========================================

QUE HACE
--------
Genera automaticamente el horario semanal de los empleados en una hoja
de Google Sheets llamada "Empleados". Rellena cada fila con los turnos
de Lunes a Sabado y calcula las horas totales semanales.


REQUISITOS
----------
- Hoja de calculo de Google Sheets con una pestana llamada "Empleados".
- Fila 1 con cabeceras: Nombre | Lunes | Martes | Miercoles | Jueves | Viernes | Sabado | Horas.
- Script pegado en Extensiones -> Apps Script.


COMO USARLO
-----------
Abre Apps Script, selecciona la funcion generarHorario y pulsa Ejecutar.
El script hara las siguientes preguntas por orden:

Por cada empleado:
  - Nombre: formato obligatorio -> Apellido2, Apellido1, Nombre
    Ejemplo: Garcia, Lopez, Ana  ->  resultado: Ana Lopez Garcia
  - Horas semanales: numero entero entre 1 y 80.
  - Dias de descanso: 0, 1 o 2. Si hay descanso, indica que dia(s).
  - Dias que dobla a la semana: entre 0 y los dias laborables disponibles.
    Este numero junto con las horas debe cuadrar exactamente.
  - Turno fijo: escribe "manana", "tarde" o "ninguno".
  - Repite para cada empleado. Pulsa Cancelar cuando hayas terminado.

Configuracion del horario (una sola vez):
  - Hora de apertura y cierre de manana (formato HH:MM, ejemplo: 10:00 - 14:00).
  - Si hay tarde: responde si o no. Si hay tarde, introduce su hora de inicio y fin.


LOGICA DE ASIGNACION
--------------------
- Los dias de descanso fijos nunca se modifican.
- Los turnos dobles se asignan primero, tantos como el usuario indico.
- El resto de dias se cubren con manana o tarde segun el turno fijo,
  o aleatoriamente si no tiene restriccion.
- Los dias se barajan aleatoriamente en cada ejecucion para dar variedad.


ADVERTENCIAS
------------
- Si las horas no cuadran con los dobles y dias disponibles, el script
  avisa pero escribe igualmente lo que pudo calcular.
- Si un empleado tiene turno fijo de tarde pero no se configuro horario
  de tarde, el script aborta con un error.
- Si el nombre no sigue el formato correcto, el script lo rechaza y pide
  que lo vuelvas a introducir.
- Si pulsas Cancelar en cualquier momento, la ejecucion se detiene.


EJEMPLO RAPIDO
--------------
Horario: manana 10:00-14:00 (4h), tarde 17:00-21:00 (4h), doble = 8h.
Empleado con 32h semanales, 1 dia de descanso y 4 dias doblando:
  4 dias doble x 8h = 32h. Correcto.
  El quinto dia laborable queda sin turno asignable y el script lo notifica.

