window.onload = function () {
    listarPersonas();
}

function listarPersonas() {
    pintar({
        divPintado: 'divPersona',
        url: 'Persona/listaPersonas',
        cabeceras: ['Nombre', 'Apellido Paterno', 'Apellido Materno'],
        propiedades: ['nombre', 'appaterno', 'apmaterno'],
        propiedadId: 'iidpersona'
    })
}

function exportarExcel() {
    fetchGet('Persona/generarReporte', 'text', function (data) {
        var a = document.createElement("a");
        a.href = "data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64," + data;
        a.click();
    })
}