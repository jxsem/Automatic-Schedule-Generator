function generarHorario() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hoja = ss.getSheetByName("Empleados");

  if (!hoja) {
    ui.alert('No encuentro la hoja "Empleados".');
    return;
  }

  let empleadosYReglas;
  try {
    empleadosYReglas = pedirEmpleadosYReglas_(ui);
  } catch (e) {
    ui.alert(String(e && e.message ? e.message : e));
    return;
  }
  const datos = empleadosYReglas.empleados;
  const reglas = empleadosYReglas.reglas;
  if (!datos.length) {
    ui.alert("No se han introducido empleados.");
    return;
  }

  hoja.getRange(2, 1, hoja.getMaxRows() - 1, 1).clearContent();
  hoja.getRange(2, 1, datos.length, 1).setValues(datos.map(n => [n]));
  const lastRow = 1 + datos.length;

  const dias = ["Lunes", "Martes", "Miércoles", "Jueves", "Viernes", "Sábado"];

  let configTurnos;
  try {
    configTurnos = obtenerConfigTurnosDesdeUsuario_(ui);
  } catch (e) {
    ui.alert(String(e && e.message ? e.message : e));
    return;
  }
  const { manana, tarde, doble, horasManana, horasTarde, horasDoble, tieneTarde } = configTurnos;

  if (!tieneTarde) {
    const empleadosSoloTarde = datos.filter(n => (reglas[n] && reglas[n].solo === "tarde"));
    if (empleadosSoloTarde.length) {
      ui.alert(
        'Has indicado que NO hay horario de tarde, pero hay empleados con regla solo:"tarde":\n\n' +
          empleadosSoloTarde.join("\n") +
          '\n\nSolución: vuelve a ejecutar y responde que SÍ hay tarde, o cambia esas reglas.'
      );
      return;
    }
  }

  hoja.getRange(2, 2, lastRow - 1, dias.length + 1).clearContent();

  datos.forEach((nombre, i) => {
    const regla = reglas[nombre] || { max: 32, dobles: 0 };
    const fila = [];
    const asignacion = {};

    // Marca descansos fijos
    dias.forEach(d => {
      if (regla.descanso?.includes(d)) asignacion[d] = "DESCANSO";
    });

    // Días disponibles para trabajar, barajados aleatoriamente
    const diasLibres = dias
      .filter(d => !regla.descanso?.includes(d))
      .sort(() => Math.random() - 0.5);

    const numDobles  = regla.dobles;
    const numSimples = diasLibres.length - numDobles;

    // ── Asignación ────────────────────────────────────────────────────────────
    let doblesRestantes = numDobles;

    diasLibres.forEach(dia => {
      let turnoElegido;

      // PRIMERO: ¿toca doble este día?
      // Se respeta incluso si el empleado tiene turno fijo de mañana o tarde
      if (doblesRestantes > 0) {
        turnoElegido = doble;
        doblesRestantes--;

      // SEGUNDO: si no toca doble, aplica el turno fijo o elige aleatoriamente
      } else if (regla.solo === "mañana") {
        turnoElegido = manana;

      } else if (regla.solo === "tarde") {
        if (!tieneTarde) throw new Error(`"${nombre}": solo tarde pero no hay horario de tarde.`);
        turnoElegido = tarde;

      } else {
        // Sin restricción: alterna mañana/tarde aleatoriamente
        turnoElegido = (tieneTarde && Math.random() < 0.5) ? tarde : manana;
      }

      asignacion[dia] = turnoElegido;
    });

    // ── Calcula las horas reales asignadas ────────────────────────────────────
    let horasReales = 0;
    dias.forEach(d => {
      const t = asignacion[d];
      if (!t || t === "DESCANSO") return;
      if (t === doble)       horasReales += horasDoble;
      else if (t === manana) horasReales += horasManana;
      else if (t === tarde)  horasReales += horasTarde;
    });

    // ── Avisa si las horas reales no coinciden con lo que el usuario pidió ────
    if (Math.abs(horasReales - regla.max) > 0.001) {
      ui.alert(
        `⚠️ "${nombre}": con ${numDobles} doble(s) y ${numSimples} turno(s) simple(s) ` +
        `resultan ${horasReales.toFixed(2)}h, pero se esperaban ${regla.max}h.\n\n` +
        `Ajusta el número de dobles o las horas semanales.`
      );
    }

    dias.forEach(dia => fila.push(asignacion[dia] || "DESCANSO"));
    fila.push(Number(horasReales.toFixed(2)));

    hoja.getRange(i + 2, 2, 1, fila.length).setValues([fila]);
  });

  hoja.autoResizeColumns(1, dias.length + 2);
}


// ─────────────────────────────────────────────
// FUNCIONES AUXILIARES
// ─────────────────────────────────────────────

function pedirEmpleadosYReglas_(ui) {
  const reglas = {};
  const empleados = [];
  const diasValidos = ["Lunes", "Martes", "Miércoles", "Jueves", "Viernes", "Sábado"];

  ui.alert(
    "Vamos a introducir empleados.\n\n" +
      'En cada paso escribe: "Segundo apellido, Primer apellido, Nombre".\n' +
      "Ejemplo: García, López, Ana\n\n" +
      "Pulsa Cancelar para terminar."
  );

  while (true) {
    const resNombre = ui.prompt(
      "Introduce empleado (Apellido2, Apellido1, Nombre)",
      "Ejemplo: García, López, Ana",
      ui.ButtonSet.OK_CANCEL
    );
    if (resNombre.getSelectedButton() !== ui.Button.OK) break;

    const nombre = normalizarNombreEmpleado_(resNombre.getResponseText());
    if (!nombre) continue;

    if (reglas[nombre]) {
      ui.alert(`El empleado "${nombre}" ya existe. Se omitirá.`);
      continue;
    }

    const max = pedirNumeroEnteroEnRango_(
      ui, `Horas semanales de "${nombre}"`, "Introduce un número (ej. 32)", 1, 80
    );

    const descansoCount = pedirNumeroEnteroEnRango_(
      ui,
      `Días de descanso de "${nombre}"`,
      "¿Descansa 0, 1 o 2 días a la semana? (0/1/2)",
      0, 2
    );

    let descanso = [];
    if (descansoCount > 0) {
      descanso = pedirDiasDescanso_(ui, nombre, descansoCount, diasValidos);
    }

    const diasDisponibles = 6 - descansoCount;
    const dobles = pedirNumeroEnteroEnRango_(
      ui,
      `¿Cuántos días doblará "${nombre}" a la semana?`,
      `Tiene ${diasDisponibles} días laborables. Responde entre 0 y ${diasDisponibles}.`,
      0, diasDisponibles
    );

    const solo = pedirSolo_(ui, nombre);

    reglas[nombre] = { max, descanso, dobles, ...(solo ? { solo } : {}) };
    empleados.push(nombre);
  }

  return { empleados, reglas };
}

function normalizarNombreEmpleado_(texto) {
  const raw = String(texto || "").trim();
  if (!raw) return "";
  const partes = raw.includes(",")
    ? raw.split(",").map(s => s.trim()).filter(Boolean)
    : raw.split(/\s+/).map(s => s.trim()).filter(Boolean);

  if (partes.length < 3) {
    throw new Error('Nombre no válido. Usa "Segundo apellido, Primer apellido, Nombre".');
  }

  const apellido2 = partes[0];
  const apellido1 = partes[1];
  const nombre    = partes.slice(2).join(" ");

  return `${nombre} ${apellido1} ${apellido2}`.replace(/\s+/g, " ").trim();
}

function pedirNumeroEnteroEnRango_(ui, titulo, mensaje, min, max) {
  const res = ui.prompt(titulo, mensaje, ui.ButtonSet.OK_CANCEL);
  if (res.getSelectedButton() !== ui.Button.OK) throw new Error("Operación cancelada.");
  const n = Number(String(res.getResponseText() || "").trim());
  if (!Number.isInteger(n) || n < min || n > max) {
    throw new Error(`Valor no válido. Debe ser un entero entre ${min} y ${max}.`);
  }
  return n;
}

function pedirDiasDescanso_(ui, nombreEmpleado, cantidad, diasValidos) {
  const res = ui.prompt(
    `¿Qué día(s) descansa "${nombreEmpleado}"?`,
    `Escribe ${cantidad} día(s) separados por coma. Opciones: ${diasValidos.join(", ")}`,
    ui.ButtonSet.OK_CANCEL
  );
  if (res.getSelectedButton() !== ui.Button.OK) throw new Error("Operación cancelada.");
  const entrada = String(res.getResponseText() || "")
    .split(",")
    .map(s => s.trim())
    .filter(Boolean);

  const normalizados = [...new Set(entrada.map(capitalizarDia_))];
  const invalidos    = normalizados.filter(d => !diasValidos.includes(d));
  if (invalidos.length) {
    throw new Error(`Día(s) no válido(s): ${invalidos.join(", ")}`);
  }
  if (normalizados.length !== cantidad) {
    throw new Error(`Debes indicar exactamente ${cantidad} día(s) de descanso.`);
  }
  return normalizados;
}

function capitalizarDia_(dia) {
  const t = String(dia || "").trim().toLowerCase();
  if (!t) return "";
  if (t === "miercoles") return "Miércoles";
  return t.charAt(0).toUpperCase() + t.slice(1);
}

function pedirSolo_(ui, nombreEmpleado) {
  const res = ui.prompt(
    `Turno fijo de "${nombreEmpleado}"`,
    'Escribe: "mañana", "tarde" o "ninguno"',
    ui.ButtonSet.OK_CANCEL
  );
  if (res.getSelectedButton() !== ui.Button.OK) throw new Error("Operación cancelada.");
  const t = String(res.getResponseText() || "").trim().toLowerCase();
  if (["ninguno", "no", "", "ambos"].includes(t)) return null;
  if (t === "mañana" || t === "manana") return "mañana";
  if (t === "tarde") return "tarde";
  throw new Error('Respuesta no válida. Escribe "mañana", "tarde" o "ninguno".');
}

function obtenerConfigTurnosDesdeUsuario_(ui) {
  const mananaInicio = pedirHora_(ui, "Introduce la HORA DE APERTURA (mañana)", "Ejemplo: 10:00");
  const mananaFin    = pedirHora_(ui, "Introduce la HORA DE CIERRE (mañana)",   "Ejemplo: 14:00");
  validarRango_(mananaInicio, mananaFin, "mañana");

  const respuestaTarde = ui.prompt(
    "¿Tienes horario de tarde?",
    'Responde "si" o "no".',
    ui.ButtonSet.OK_CANCEL
  );
  if (respuestaTarde.getSelectedButton() !== ui.Button.OK) throw new Error("Operación cancelada.");
  const tieneTarde = normalizarSiNo_(respuestaTarde.getResponseText());

  let tardeInicio = null;
  let tardeFin    = null;
  if (tieneTarde) {
    tardeInicio = pedirHora_(ui, "¿Desde qué hora es la TARDE?", "Ejemplo: 16:30");
    tardeFin    = pedirHora_(ui, "¿Hasta qué hora es la TARDE?", "Ejemplo: 20:30");
    validarRango_(tardeInicio, tardeFin, "tarde");
  }

  const manana      = `${formatearHora_(mananaInicio)}-${formatearHora_(mananaFin)}`;
  const horasManana = (mananaFin - mananaInicio) / 60;

  let tarde      = "";
  let horasTarde = 0;
  let doble      = manana;
  let horasDoble = horasManana;

  if (tieneTarde) {
    tarde      = `${formatearHora_(tardeInicio)}-${formatearHora_(tardeFin)}`;
    horasTarde = (tardeFin - tardeInicio) / 60;
    doble      = `${manana} / ${tarde}`;
    horasDoble = horasManana + horasTarde;
  }

  return { manana, tarde, doble, horasManana, horasTarde, horasDoble, tieneTarde };
}

function pedirHora_(ui, titulo, ejemplo) {
  const res = ui.prompt(titulo, ejemplo, ui.ButtonSet.OK_CANCEL);
  if (res.getSelectedButton() !== ui.Button.OK) throw new Error("Operación cancelada.");
  return parsearHoraAMinutos_(res.getResponseText());
}

function normalizarSiNo_(texto) {
  const t = String(texto || "").trim().toLowerCase();
  if (["si", "sí", "s", "y", "yes"].includes(t)) return true;
  if (["no", "n"].includes(t)) return false;
  throw new Error('Respuesta no válida. Escribe "si" o "no".');
}

function parsearHoraAMinutos_(texto) {
  const t = String(texto || "").trim();
  const m = /^(\d{1,2}):(\d{2})$/.exec(t);
  if (!m) throw new Error(`Hora no válida: "${t}". Usa formato HH:MM (ej. 10:00).`);
  const hh = Number(m[1]);
  const mm = Number(m[2]);
  if (Number.isNaN(hh) || Number.isNaN(mm) || hh < 0 || hh > 23 || mm < 0 || mm > 59) {
    throw new Error(`Hora fuera de rango: "${t}".`);
  }
  return hh * 60 + mm;
}

function formatearHora_(minutos) {
  const hh = Math.floor(minutos / 60);
  const mm = minutos % 60;
  return `${String(hh).padStart(2, "0")}:${String(mm).padStart(2, "0")}`;
}

function validarRango_(inicioMin, finMin, nombre) {
  if (finMin <= inicioMin) {
    throw new Error(`Rango de ${nombre} no válido: la hora de cierre debe ser posterior a la de apertura.`);
  }
  const durMin = finMin - inicioMin;
  if (durMin < 30) {
    throw new Error(`Rango de ${nombre} demasiado corto (${durMin} min). Revisa las horas.`);
  }
}