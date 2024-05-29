 $(document).ready(function () {
            $('#tbllist').DataTable({
                pageLength: 3,
                lengthChange: false,
                pagingType: 'simple_numbers'
            });

            $.ajax({
                url: '/Home/ObtenerReprocesos',
                type: 'GET',
                success: function (data) {
                    if (data && data.length > 0) {
                        $('#tbllist').DataTable().clear().destroy(); // Limpia y destruye la tabla actual
                        $('#tbllist').DataTable({
                            data: data,
                            columns: [
                                { data: 'FechaOp' },
                                { data: 'FechaReg' },
                                { data: 'Consecutivo' },
                                { data: 'Moneda' },
                                { data: 'Valor' },
                                { data: 'Nit' },
                                { // Columna de acción con botón de eliminar
                                    data: null,
                                    defaultContent: `
              <div style="display: inline-block;align-items: center;">
                <button class="btn-delete" type="button" style="display: inline-block; width: 1px; height: 1px; margin-right: 20px; margin-top:2px">
                  <i class="fas fa-trash" style="color: black;"></i>
                </button>
                <button class="btn-edit" type="button" style="display: inline-block; width: 1px; height: 1px;margin-top:2px">
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
                                                url: '/Home/EliminarReproceso',
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

                                        const form = document.querySelector("form")
                                        form.classList.add('secActive');

                                        $('#id').val(data.Id);
                                        $('#consecutivo').val(data.Consecutivo);
                                        $('#fecha').val(data.FechaOp);
                                        $('#nit').val(data.Nit);
                                        $('#valor').val(data.Valor);
                                        $('#moneda').val(data.Moneda);
                                        $('#cliente').val(data.Cliente);
                                        $('#responsable').val(data.Responsable);
                                        $('#producto').val(data.ProdEvento);
                                        $('#area').val(data.Area);
                                        $('#tipoError').val(data.TipoError);
                                        $('#perdida').val(data.Perdida);
                                        $('#impacto').val(data.Impacto);
                                        $('#causa').val(data.Causa);
                                        $('#descripcion').val(data.Descripcion);
                                    });
                             

                            }
                        });

                    } else {
                        console.log('No se encontraron datos');
                    }
                },
                error: function (error) {
                    console.log(error);
                }
            });

     


        $('#btnBuscar').on('click', function () {
            var fechaini = $('#fechaIni').val();
            var fechafin = $('#fechaFin').val();
            var area = $('#area').val();

            $.ajax({
                url: '/Home/ObtenerReprocesosFechaArea',
                type: 'GET',
                data: {
                    fechaini: fechaini,
                    fechafin: fechafin,
                    area: area
                },
                success: function (data) {
                    if (data && data.length > 0) {
                        $('#tbllist').DataTable().clear().destroy(); // Limpia y destruye la tabla actual
                        $('#tbllist').DataTable({
                            data: data,
                            columns: [
                                { data: 'FechaOp' },
                                { data: 'FechaReg' },
                                { data: 'Consecutivo' },
                                { data: 'Moneda' },
                                { data: 'Valor' },
                                { data: 'Nit' },
                                { // Columna de acción con botón de eliminar
                                    data: null,
                                    defaultContent: `
              <div style="display: inline-block;align-items: center;">
                <button class="btn-delete" type="button" style="display: inline-block; width: 1px; height: 1px; margin-right: 20px; margin-top:2px">
                  <i class="fas fa-trash" style="color: black;"></i>
                </button>
                <button class="btn-edit" type="button" style="display: inline-block; width: 1px; height: 1px;margin-top:2px">
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
                                                url: '/Home/EliminarReproceso',
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

                                    const form = document.querySelector("form")
                                    form.classList.add('secActive');

                                    $('#id').val(data.Id);
                                    $('#consecutivo').val(data.Consecutivo);
                                    $('#fecha').val(data.FechaOp);
                                    $('#nit').val(data.Nit);
                                    $('#valor').val(data.Valor);
                                    $('#moneda').val(data.Moneda);
                                    $('#cliente').val(data.Cliente);
                                    $('#responsable').val(data.Responsable);
                                    $('#producto').val(data.ProdEvento);
                                    $('#area').val(data.Area);
                                    $('#tipoError').val(data.TipoError);
                                    $('#perdida').val(data.Perdida);
                                    $('#impacto').val(data.Impacto);
                                    $('#causa').val(data.Causa);
                                    $('#descripcion').val(data.Descripcion);
                                });
                            }
                        });
                    } else {
                        console.log('No se encontraron datos');
                    }
                },
                error: function (error) {
                    console.log(error);
                }
            });
        });




        $('#tbllist tbody').on('click', 'tr', function (e) {
            if (!$(e.target).hasClass('btn-delete') && !$(e.target).closest('.btn-delete').length && !$(e.target).hasClass('btn-edit') && !$(e.target).closest('.btn-edit').length) {
                var rowData = $('#tbllist').DataTable().row(this).data();
                // Aquí puedes usar rowData para mostrar todos los detalles en algún lugar
                mostrarDetallesCompletos(rowData);
            }
        });

        function mostrarDetallesCompletos(detalle) {
            // Construir una cadena HTML con los detalles
            var detalleHtml = '<div style="font-size: 12px; text-align: left;">';
            for (var key in detalle) {
                if (detalle.hasOwnProperty(key)) {
                    detalleHtml += '<div><strong>' + key + ':</strong> ' + detalle[key] + '</div>';
                }
            }
            detalleHtml += '</div>';

            Swal.fire({
                icon: 'success',
                title: '¡Datos de reproceso!',
                html: detalleHtml,
                confirmButtonColor: '#9820EC',
                cancelButtonColor: '#7900D3',
                confirmButtonText: 'OK'
            });
            console.log(detalle); // Ejemplo de impresión en consola
        }
        });

        document.addEventListener('DOMContentLoaded', function () {
            document.getElementById('btnCerrarSesion').addEventListener('click', function () {
                window.location.href = '/Home/Index';
            });

        });

        const form = document.querySelector("form"),
            backBtn = form.querySelector(".backBtn");
        backBtn.addEventListener("click", () => form.classList.remove('secActive'));

    </script>

    <script>
        $(document).ready(function () {
            $('#btnGuardarDatos').click(function () {
                // Obtener los valores de los campos del formulario
                var usuario = "@ViewBag.WindowsUsername";
                var fechaOp = $('#fecha').val().toString();
                var fechaReg = obtenerFechaActual().toString();
                var consecutivo = $('#consecutivo').val().toString();
                var nit = $('#nit').val().toString();
                var moneda = $('#moneda').val();
                var valor = parseFloat($('#valor').val());
                var cliente = $('#cliente').val().toString();
                var producto = $('#producto').val().toString();
                var responsable = $('#responsable').val().toString();
                var usuarioReproceso =usuario.toString();
                var perdida = $('#perdida').val().toString();
                var impacto = $('#impacto').val().toString();
                var causa = $('#causa').val().toString();
                var descripcion = $('#descripcion').val().toString();
                var area = $('#area').val().toString();
                var tipoError = $('#tipoError').val().toString();
                var Id = $('#id').val();

                // Convertir la fecha a objeto Date y obtener el nombre del mes
                var fecha = new Date($('#fecha').val());
                var mes = fecha.toLocaleString('es-ES', { month: 'long' });


                // Enviar los datos al controlador mediante AJAX
                $.ajax({
                    url: '/Home/EditarReproceso',
                    type: 'POST',
                    data: {
                        Id: Id,
                        FechaOp: fechaOp,
                        FechaReg: fechaReg,
                        Consecutivo: consecutivo,
                        Nit: nit,
                        Moneda: moneda,
                        Valor: valor,
                        Cliente: cliente,
                        ProdEvento: producto,
                        Responsable: responsable,
                        UsuarioReproceso: usuarioReproceso,
                        Perdida: perdida,
                        Impacto: impacto,
                        Causa: causa,
                        Descripcion: descripcion,
                        Area: area,
                        TipoError: tipoError,
                        Mes: mes
                    },
                    success: function (response) {
                        // Manejar la respuesta del servidor si es necesario
                        console.log(response);
                        Swal.fire({
                            icon: 'success',
                            title: '¡Datos editados!',
                            text: 'Los datos se han guardado correctamente, en la vista de consultas lo puede visualizar',
                            confirmButtonText: 'OK'
                        });
                    },
                    error: function (error) {
                        // Manejar errores de AJAX si es necesario
                        console.log(error);
                        Swal.fire({
                            icon: 'Error',
                            title: '¡Datos no se han guardado!',
                            text: error,
                            confirmButtonText: 'OK'
                        });
                    }
                });
            });
        });

        // Función para obtener la fecha actual en el formato deseado
        function obtenerFechaActual() {
            var fecha = new Date();
            var year = fecha.getFullYear();
            var month = ("0" + (fecha.getMonth() + 1)).slice(-2);
            var day = ("0" + fecha.getDate()).slice(-2);
            return year + "-" + month + "-" + day;
        }



        $('#btnExportarExcel').click(function () {
            var tableData = $('#tbllist').DataTable().rows().data().toArray();

            // Convertir los datos de la tabla a JSON
            var jsonData = JSON.stringify(tableData);

            // Realizar la solicitud AJAX para enviar el JSON al controlador
            $.ajax({
                url: '/Home/LlenarReprocesosInformes',
                type: 'POST',
                contentType: 'application/json',
                data: JSON.stringify({ jsonData: jsonData }),
                success: function (response) {
                    if (response.success) {
                        // Manejar la respuesta del servidor si es necesario
                        Swal.fire({
                            icon: 'success',
                            title: '¡Informe de calidad generado!',
                            html: 'Se ha generado correctamente el informe',
                            confirmButtonText: 'OK'
                        });
                    }
                  
                },
                error: function (error) {
                    // Manejar errores de AJAX si es necesario
                    console.log(error);
                    Swal.fire({
                        icon: 'error',
                        title: '¡Informe de calidad no fue generado!',
                        html: 'No se ha generado correctamente el informe',
                        confirmButtonText: 'OK'
                    });
                }
            });
        });


