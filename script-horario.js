function generarHorario() {
  // Obtiene la hoja activa y selecciona la hoja llamada “Empleados”
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hoja = ss.getSheetByName("Empleados");

  // Toma los nombres desde la columna A (empezando en A2 hacia abajo)
  const datos = hoja.getRange(2, 1, hoja.getLastRow() - 1, 1).getValues().flat();

  // Días de la semana que se van a usar como columnas
  const dias = ["Lunes", "Martes", "Miércoles", "Jueves", "Viernes", "Sábado"];

  // Definición de los tres tipos de turnos
  const mañana = "10:00-14:00";
  const tarde = "16:30-20:30";
  const doble = `${mañana} / ${tarde}`; // turno doble (mañana + tarde)

  // Objetos personalizados por empleado -> clave - valor
  // - max: horas semanales máximas
  // - descanso: días libres
  // - solo: “mañana” o “tarde” si solo trabaja en ese turno
  // - doblesMin / doblesMax: mínimo o máximo de días con turno doble
  const reglas = {
    "Adrián Martos": { max: 40, descanso: ["Jueves"], doblesMax: 5, doblesMin: 5},
    "María Ruiz": { max: 20, descanso: ["Jueves"], solo: "mañana" },
    "Jose Peinado": { max: 40, descanso: ["Miércoles"], doblesMax: 5, doblesMin: 5 },
    "Javier Lopez": { max: 32, doblesMax: 2 },
    "Hilario Pastrana": { max: 32, solo: "tarde", doblesMin: 2 },
    "Fran Catena": { max: 32, solo: "mañana", doblesMin: 2 },
    "Fran Cervilla": { max: 32, doblesMax: 3, doblesMin: 3 },
    "Jose Manuel": { max: 32, doblesMax: 2, doblesMin: 2 },
    "Antonio Ferron": { max: 32, doblesMax: 2, doblesMin: 2 },
    "Cinthya Suazo": { max: 28, doblesMax: 1, doblesMin: 2 },
    "Sol Morales": { max: 32, doblesMax: 2, doblesMin: 2 }
  };

  // Limpia las columnas B–H (Lunes–Sábado + Horas) antes de generar el nuevo horario
  hoja.getRange(2, 2, hoja.getLastRow() - 1, dias.length + 1).clearContent();

  // Recorre cada empleado
  datos.forEach((nombre, i) => {
    const regla = reglas[nombre] || { max: 32 }; // si no tiene regla, usa 32h por defecto
    let horas = 0;
    const fila = [];

    // Crea una copia de los días pero barajada aleatoriamente
    const diasMezclados = dias.slice().sort(() => Math.random() - 0.5);

    // Objeto si algún empleado NO tiene dia de descanso asignado
    const asignacion = {};
    // --- ASIGNAR DESCANSO AUTOMÁTICO A FRAN CERVILLA ---
    if (nombre === "Fran Cervilla" && !regla.descanso) {
    const diaAleatorio = dias[Math.floor(Math.random() * dias.length)];
    regla.descanso = [diaAleatorio];
  }

    // --- PRIMERA PASADA: asignación base ---
    diasMezclados.forEach(dia => {
      // Si el día está en su descanso, marca “-” y pasa al siguiente
      if (regla.descanso?.includes(dia)) {
        asignacion[dia] = "DESCANSO";
        return;
      }

      // Cuenta los días que ya tienen turno doble
      const doblesActuales = Object.values(asignacion).filter(x => x.includes("/")).length;
      const puedeDoblar = !regla.doblesMax || doblesActuales < regla.doblesMax;
      // Si dobla un dia necesita un descanso entonces almacenamos el descanso en una variable
      let turno = "DESCANSO";

      // Si tiene restricción de “solo mañana” o “solo tarde”
      if (regla.solo === "mañana") {
        turno = mañana;
        horas += 4;
      } else if (regla.solo === "tarde") {
        turno = tarde;
        horas += 4;
      } else {

        // Asignación aleatoria de turno (doble, mañana o tarde)
        const rand = Math.random(); // Genera un número aleatorio entre 0 y 1 (por ejemplo: 0.12, 0.57, 0.91...).
        if (rand < 0.25 && horas + 8 <= regla.max && puedeDoblar) { 
          turno = doble;
          horas += 8;
        } else if (rand < 0.6 && horas + 4 <= regla.max) {
          turno = mañana;
          horas += 4;
        } else if (horas + 4 <= regla.max) {
          turno = tarde;
          horas += 4;
        }
      }
      //Ese reparto de probabilidades (25 % dobles / 35 % mañanas / 40 % tardes) no es casualidad —> se suele ajustar según cómo esté configurado el resto del personal.
      // Ya que hay como 4 personas con turno fijo de mañanas y habría muchas personas ya de mañanas y pocas de tardes.

      asignacion[dia] = turno;
    });

    //Resultado al final de esta fase:
    //El empleado ya tiene un horario tentativo, pero puede no cumplir aún los mínimos o máximos de dobles.

    // --- SEGUNDA PASADA: forzar mínimos de dobles ---
    if (regla.doblesMin) { // Aquí el script revisa cuántos turnos dobles tiene el empleado y los ajusta hacia arriba si no llega al mínimo.
      let dobles = Object.values(asignacion).filter(x => x.includes("/")).length;// Busca cuantos "/" hay y los guarda en una variable llamada dobles
      const diasDisponibles = diasMezclados.filter(
        d => !asignacion[d].includes("/") && asignacion[d] !== "DESCANSO" // Crea una lista de días disponibles donde no hay turno doble ni descanso. O sea, busca huecos donde podría meter un turno doble.
      );
      while (dobles < regla.doblesMin && diasDisponibles.length) { // Aquí empieza a forzar dobles hasta alcanzar el mínimo obligatorio (doblesMin)
        const dia = diasDisponibles.pop();
        asignacion[dia] = doble; // cambia ese turno normal por un turno doble.
        horas += 4;
        dobles++;
      }
    }
    // En resumen -> Este bloque corrige el azar de la primera pasada

    // --- TERCERA PASADA: reducir dobles si pasa del máximo ---> Lo mismo de arriba pero al revés
    if (regla.doblesMax) {
      let dobles = Object.entries(asignacion)
        .filter(([_, v]) => v.includes("/"))
        .map(([d]) => d); // Basicamente realiza una funcion donde cuenta cuantos "/" hay y si hay mas de los que la variable dobleMax permite se realizara un random de un 50% diciendo si el turno que sobra estará de mañanas o tarde
      while (dobles.length > regla.doblesMax) {
        const dia = dobles.pop();
        // cambia el turno doble por mañana o tarde
        asignacion[dia] = Math.random() < 0.5 ? mañana : tarde;
        horas -= 4;
      }
    }

    // Aquí el código revisa si el empleado tiene más turnos dobles de lo permitido (doblesMax).
    // Si tiene más, va quitando dobles y los convierte en mañana/tarde normales hasta cumplir el límite.

    // Reordena los días a su orden normal (Lunes–Sábado)
    dias.forEach(dia => fila.push(asignacion[dia] || "—"));

    // Añade las horas totales al final
    fila.push(horas);

    // Escribe la fila en la hoja (a partir de la columna B)
    hoja.getRange(i + 2, 2, 1, fila.length).setValues([fila]); // filaInicial, columnaInicial, numFilas, numColumnas
  });

  // Ajusta automáticamente el ancho de columnas para que se vea bien
  hoja.autoResizeColumns(1, dias.length + 2);
}


// SCRIPT REALIZADO POR JOSE MANUEL SOLDADO J. 
