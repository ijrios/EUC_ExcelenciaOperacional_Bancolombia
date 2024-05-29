$(document).ready(function () {
    // Manejar el clic en el botón "Registrar"
    $("#btnRegistrar").click(function () {
        // Redirigir a la página deseada
        window.location.href = '/Home/Registro';

    });

    var mensajeRechazo = '@Request.QueryString["mensajeRechazo"]';

    if (mensajeRechazo) {
        Swal.fire({
            icon: 'Error',
            title: 'Sin autorización!',
            text: mensajeRechazo,
            confirmButtonText: 'OK'
        });
    }

    // Manejar el clic en el botón "Consultar"
    $("#btnConsultar").click(function () {
        // Redirigir a la página deseada
        window.location.href = '/Home/Consultas';
    });

    // Manejar el clic en el botón "Administrar"
    $("#btnAdministrar").click(function () {
        // Redirigir a la página deseada
        window.location.href = '/Home/Index';
    });
});
