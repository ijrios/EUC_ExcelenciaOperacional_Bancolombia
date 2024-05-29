       $(document).ready(function () {
            $('#tbllist').DataTable({
                pageLength: 3,
                lengthChange: false,
                pagingType: 'simple_numbers'
            });

            $.ajax({
                url: '/Home/ObtenerUsuario',
                type: 'GET',
                success: function (data) {
                    if (data && data.length > 0) {
                        $('#tbllist').DataTable().clear().destroy(); // Limpia y destruye la tabla actual
                        $('#tbllist').DataTable({
                            data: data,
                            columns: [
                                { data: 'Usuario' },
                                { data: 'Nombre' },
                                { data: 'Correo' },
                                { data: 'Perfil' },
                                { // Columna de acción con botones de eliminar y editar
                                    data: null,
                                    defaultContent: `
                                                <div style="display: inline-block; align-items: center;">
                                                    <button class="btn-delete" type="button" style="display: inline-block;  width: 1px; height: 1px; margin-right: 20px; margin-top:2px">
                                                        <i class="fas fa-trash" style="color: black;"></i>
                                                    </button>
                                                    <button class="btn-edit"  type="button" style="display: inline-block; width: 1px; height: 1px;margin-top:2px">
                                                        <i class="fas fa-edit" style="color: black;"></i>
                                                    </button>
                                                </div>`
                                }
                            ],
                            pageLength: 3,
                            lengthChange: false,
                            pagingType: 'simple_numbers',
                            initComplete: function () {

                                $('#tbllist tbody').on('click', '.btn-delete', function () {
                                    var data = $('#tbllist').DataTable().row($(this).parents('tr')).data();
                                    var id = data.Id;

                                    // Confirmación antes de eliminar
                                    Swal.fire({
                                        title: '¿Estás seguro?',
                                        text: "¡No podrás revertir esto!",
                                        icon: 'warning',
                                        showCancelButton: true,
                                        confirmButtonColor: '#9820EC',
                                        cancelButtonColor: '#7900D3',
                                        confirmButtonText: 'Sí, eliminarlo!'
                                    }).then((result) => {
                                        if (result.isConfirmed) {
                                            $.ajax({
                                                url: '/Home/EliminarUsuario',
                                                type: 'POST',
                                                data: {
                                                    id: id
                                                },
                                                success: function (response) {
                                                    console.log(response);
                                                    Swal.fire({
                                                        icon: 'success',
                                                        title: '¡Eliminado!',
                                                        text: 'El registro ha sido eliminado correctamente.',
                                                        confirmButtonText: 'OK'
                                                    });
                                                },
                                                error: function (error) {
                                                    console.log(error);
                                                    Swal.fire({
                                                        icon: 'error',
                                                        title: 'Error al eliminar',
                                                        text: 'Ha ocurrido un error al intentar eliminar el registro.',
                                                        confirmButtonText: 'OK'
                                                    });
                                                }
                                            });
                                        }
                                    });
                                });


                                $('#tbllist tbody').on('click', '.btn-edit', function () {
                                    var data = $('#tbllist').DataTable().row($(this).parents('tr')).data();

                                    $('#usuario').val(data.Usuario);
                                    $('#correo').val(data.Correo);
                                    $('#nombre').val(data.Nombre);
                                    // Asignar la opción correspondiente al select según el valor de data.Perfil
                                    if (data.Perfil === "1111") {
                                        $('#perfil select').val("Administrador");
                                    } else if (data.Perfil === "1001") {
                                        $('#perfil select').val("Ingreso, informes");
                                    } else if (data.Perfil === "1100") {
                                        $('#perfil select').val("Ingreso, consultas");
                                    } else if (data.Perfil === "1101") {
                                        $('#perfil select').val("Ingreso, Consultas, Informes");
                                    } else if (data.Perfil === "0100") {
                                        $('#perfil select').val("Consultas");
                                    } else {
                                        $('#perfil select').val(""); // O alguna otra acción si el perfil no coincide
                                    }

                                });


                            }
                        });
                    } else {
                        console.log('No se encontraron datos');
                        Swal.fire({
                            icon: 'Error',
                            title: 'Error',
                            text: 'No se encontraron datos',
                            confirmButtonText: 'OK'
                        });
                    }
                },
                error: function (error) {
                    console.log(error);
                }
            });
        });
