function generarHorario() {
  const hojaEmpleados = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Empleados");
  const hojaHorario = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Horario");

  if (!hojaEmpleados || !hojaHorario) {
    SpreadsheetApp.getUi().alert("Las hojas 'Empleados' o 'Horario' no existen. Revisa los nombres.");
    return;
  }

  const todosEmpleados = hojaEmpleados.getRange(2, 1, hojaEmpleados.getLastRow() - 1).getValues().flat();
  const empleados = todosEmpleados;

  if (empleados.length < 6) {
    SpreadsheetApp.getUi().alert("Debe haber al menos 6 empleados disponibles para asignar.");
    return;
  }

  const dias = ["Lunes", "Martes", "Miércoles", "Jueves", "Viernes", "Sábado"]; // Domingo excluido

  const maxHoras = {
    "Adrián Martos": 40,
    "María Ruiz": 24,
    "Jose Peinado": 40,
    "Javier Lopez": 32,
    "Hilario Pastrana": 32,
    "Fran Catena": 32,
    "Fran Cervilla": 32,
    "Jose Manuel": 32,
    "Antonio Ferron": 32,
    "Cinthya Suazo": 28
  };

  const descansos = {
    "Adrián Martos": [3],
    "María Ruiz": [3],
    "Jose Peinado": [2]
  };

  const asignacionTurnos = {};
  empleados.forEach(e => {
    if (["Fran Catena", "Cinthya Suazo", "María Ruiz"].includes(e)) asignacionTurnos[e] = "mañana";
    else if (e === "Hilario Pastrana") asignacionTurnos[e] = "tarde";
    else asignacionTurnos[e] = null;
  });

  if (!empleados.includes("Fran Cervilla")) {
    SpreadsheetApp.getUi().alert("Falta Fran Cervilla en la lista de empleados.");
    return;
  }
  asignacionTurnos["Fran Cervilla"] = Math.random() < 0.5 ? "mañana" : "tarde";
  const diaDescansoFranC = Math.floor(Math.random() * dias.length);
  descansos["Fran Cervilla"] = [diaDescansoFranC];

  const otrosEmpleados = empleados.filter(e => asignacionTurnos[e] === null);
  for (let i = 0; i < otrosEmpleados.length; i++) {
    asignacionTurnos[otrosEmpleados[i]] = i % 2 === 0 ? "mañana" : "tarde";
  }

  function asignarDiasDoble(cantidad, empleado) {
    if (empleado === "María Ruiz") return [];
    let diasDisponibles = Array.from({ length: dias.length }, (_, i) => i);
    if (descansos[empleado]) diasDisponibles = diasDisponibles.filter(d => !descansos[empleado].includes(d));
    if (["Cinthya Suazo"].includes(empleado)) diasDisponibles = diasDisponibles.filter(d => asignacionTurnos[empleado] === "mañana");

    const seleccionados = [];
    for (let i = 0; i < cantidad; i++) {
      if (diasDisponibles.length === 0) break;
      const idx = Math.floor(Math.random() * diasDisponibles.length);
      seleccionados.push(diasDisponibles[idx]);
      diasDisponibles.splice(idx, 1);
    }
    return seleccionados;
  }

  const dobleAsignacion = {};
  empleados.forEach(e => {
    switch(e) {
      case "Fran Catena": dobleAsignacion[e] = asignarDiasDoble(2, e); break;
      case "Hilario Pastrana": dobleAsignacion[e] = asignarDiasDoble(2, e); break;
      case "Cinthya Suazo": dobleAsignacion[e] = asignarDiasDoble(1, e); break;
      case "María Ruiz": dobleAsignacion[e] = []; break;
      case "Adrián Martos": dobleAsignacion[e] = asignarDiasDoble(5, e); break;
      case "Jose Peinado": dobleAsignacion[e] = [0,1,3,4,5]; break;
      case "Fran Cervilla": dobleAsignacion[e] = asignarDiasDoble(3, e); break;
      default: dobleAsignacion[e] = asignarDiasDoble(2, e);
    }
  });

  const asignacionRoles = {};
  const mezclaRoles = [...empleados].sort(() => Math.random() - 0.5);
  const mitadRoles = Math.floor(empleados.length / 2);
  empleados.forEach(e => {
    if (e === "Cinthya Suazo") asignacionRoles[e] = "Ventas";
    else if (e === "María Ruiz") asignacionRoles[e] = "";
    else asignacionRoles[e] = mezclaRoles.indexOf(e) < mitadRoles ? "Compras" : "Ventas";
  });

  // Horas fijas de María: lunes 5h, martes 4h, miércoles 5h, jueves 0h, viernes 5h, sábado 5h
  const horasMaria = [5, 4, 5, 0, 5, 5];

  const horasTrabajadas = {};
  empleados.forEach(e => horasTrabajadas[e] = 0);

  for (let i = 0; i < dias.length; i++) {
    const mananaEmpleados = empleados.filter(e => asignacionTurnos[e] === "mañana" && e !== "María Ruiz");
    const tardeEmpleados = empleados.filter(e => asignacionTurnos[e] === "tarde");

    empleados.forEach(e => {
      if (descansos[e] && descansos[e].includes(i)) {
        const idxM = mananaEmpleados.indexOf(e);
        if(idxM > -1) mananaEmpleados.splice(idxM, 1);
        const idxT = tardeEmpleados.indexOf(e);
        if(idxT > -1) tardeEmpleados.splice(idxT, 1);
      }
    });

    empleados.forEach((e, idxE) => {
      let textoTurno = "";

      if (e === "María Ruiz" && !descansos[e].includes(i)) {
        textoTurno = i === 1 ? "9:00 - 13:00" : "9:00 - 14:00";
      } else if (dobleAsignacion[e].includes(i)) {
        textoTurno = "10:00 - 14:00\n16:30 - 20:30";
      } else if (mananaEmpleados.includes(e)) {
        textoTurno = "10:00 - 14:00";
      } else if (tardeEmpleados.includes(e)) {
        textoTurno = "16:30 - 20:30";
      }

      hojaEmpleados.getRange(idxE+2, i+2).setValue(textoTurno);
      hojaEmpleados.getRange(idxE+2, i+2).setWrap(true);
    });

    // Acumular horas
    empleados.forEach(e => {
      let horasDia = 0;
      if (e === "María Ruiz") horasDia = horasMaria[i]; // solo su arreglo fijo
      else if (dobleAsignacion[e].includes(i)) horasDia = Math.min(8, maxHoras[e] - horasTrabajadas[e]);
      else if (mananaEmpleados.includes(e) || tardeEmpleados.includes(e)) horasDia = Math.min(4, maxHoras[e] - horasTrabajadas[e]);

      horasTrabajadas[e] += horasDia;
    });
  }

  empleados.forEach((e, idx) => hojaEmpleados.getRange(idx+2,9).setValue(horasTrabajadas[e]));

  SpreadsheetApp.getUi().alert("¡Horario generado y horas calculadas con éxito!");
}