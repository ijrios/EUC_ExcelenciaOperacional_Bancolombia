
        $(document).ready(function () {
            $('#tbllist').DataTable();


            $('#btnConsulta').click(function () {

                const chartType = document.getElementById("tipo").value;

                if (chartType === "eco") {
                    $.ajax({
                        url: '/Home/ObtenerReprocesosPerdidas',
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
                    <button class="btn-delete" style="display: inline-block; width: 1px; height: 1px; margin-right: 20px; margin-top:2px">
                      <i class="fas fa-trash" style="color: black;"></i>
                    </button>
                    <button class="btn-edit" style="display: inline-block; width: 1px; height: 1px;margin-top:2px">
                      <i class="fas fa-edit" style="color: black;"></i>
                    </button>
                  </div>`
                                        }
                                    ],
                                    pageLength: 3,
                                    lengthChange: false,
                                    "language": {
                                        "search": "Buscar"
                                    },
                                    pagingType: 'simple_numbers',
                                    initComplete: function () {



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
                } else if (chartType === "noeco") {
                    $.ajax({
                        url: '/Home/ObtenerReprocesosNoPerdidas',
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
                    <button class="btn-delete" style="display: inline-block; width: 1px; height: 1px; margin-right: 20px; margin-top:2px">
                      <i class="fas fa-trash" style="color: black;"></i>
                    </button>
                    <button class="btn-edit" style="display: inline-block; width: 1px; height: 1px;margin-top:2px">
                      <i class="fas fa-edit" style="color: black;"></i>
                    </button>
                  </div>`
                                        }
                                    ],
                                    pageLength: 3,
                                    "language": {
                                        "search": "Buscar"
                                    },
                                    lengthChange: false,
                                    pagingType: 'simple_numbers',
                                    initComplete: function () {



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


        $('#btnExportarExcel').click(function () {
            var tableData = $('#tbllist').DataTable().rows().data().toArray();

            // Convertir los datos de la tabla a JSON
            var jsonData = JSON.stringify(tableData);

            // Realizar la solicitud AJAX para enviar el JSON al controlador
            $.ajax({
                url: '/Home/LlenarReprocesos',
                type: 'POST',
                contentType: 'application/json',
                data: JSON.stringify({ jsonData: jsonData }),
                success: function (response) {
                    if (response.success) {
                        // Manejar la respuesta del servidor si es necesario
                        Swal.fire({
                            icon: 'success',
                            title: '¡Informe de perdidas generado!',
                            html: 'Se ha gerado correctamente el informe',
                            confirmButtonColor: '#00c288',
                            cancelButtonColor: '#7900D3',
                            confirmButtonText: 'OK'
                        });
                    }
                },
                error: function (error) {
                    // Manejar errores de AJAX si es necesario
                    console.log(error);
                    Swal.fire({
                        icon: 'error',
                        title: '¡Informe de perdidas no fue generado!',
                        html: 'No se ha gerado correctamente el informe',
                        confirmButtonColor: '#00c288',
                        cancelButtonColor: '#7900D3',
                        confirmButtonText: 'OK'
                    });
                }
            });
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
                confirmButtonColor: '#00c288',
                cancelButtonColor: '#7900D3',
                confirmButtonText: 'OK'
            });
            console.log(detalle); // Ejemplo de impresión en consola
        }

        document.addEventListener('DOMContentLoaded', function () {
            document.getElementById('btnCerrarSesion').addEventListener('click', function () {
                window.location.href = '/Home/Informes';
            });

        });

        const form = document.querySelector("form"),
            backBtn = form.querySelector(".backBtn");
        backBtn.addEventListener("click", () => form.classList.remove('secActive'));
