
        $(document).ready(function () {
            // Inicializa la tabla vacía
            $('#tbllist').DataTable({
                "language": {
                    "search": "Buscar"
                },
                "pageLength": 3,
                "lengthChange": false,
                "pagingType": "simple_numbers"
            });
            // Realiza la solicitud AJAX para obtener los datos de UsuariosErrores
            $.ajax({
                url: '/Home/ObtenerError',
                type: 'GET',
                success: function (data) {
                    if (data && data.length > 0) {
                        $('#tbllist').DataTable().clear().destroy(); // Limpia y destruye la tabla actual
                        $('#tbllist').DataTable({
                            data: data,
                            columns: [
                                { data: 'Usuario' },
                                { data: 'Errores' },
                                { data: 'Porcentaje' },
                                { // Columna de acción con botones de eliminar y editar
                                    data: null,
                                    defaultContent: `
                                        <div style="display: inline-block; align-items: center;">
                                            <button class="btn-delete" style="display: inline-block;  width: 1px; height: 1px; margin-right: 20px; margin-top:2px">
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
                                "search": "Buscar",
                                "decimal": "",
                                "emptyTable": "No hay datos disponibles en la tabla",
                                "info": "Mostrando _START_ a _END_ de _TOTAL_ entradas",
                                "infoEmpty": "Mostrando 0 a 0 de 0 entradas",
                                "infoFiltered": "(filtrado de _MAX_ entradas totales)",
                                "infoPostFix": "",
                                "thousands": ",",
                                "lengthMenu": "Mostrar _MENU_ entradas",
                                "loadingRecords": "Cargando...",
                                "processing": "Procesando...",
                                "search": "Buscar:",
                                "zeroRecords": "No se encontraron registros coincidentes",
                                "paginate": {
                                    "first": "Primero",
                                    "last": "Último",
                                    "next": "Siguiente",
                                    "previous": "Anterior"
                                },
                                "aria": {
                                    "sortAscending": ": activar para ordenar la columna de manera ascendente",
                                    "sortDescending": ": activar para ordenar la columna de manera descendente"
                                }

                            },
                            pagingType: 'simple_numbers',
                            createdRow: function (row, data, dataIndex) {
                                if (parseInt(data.Errores) > 16) {
                                    $(row).addClass('row-red');
                                } else if (parseInt(data.Errores) < 16) {
                                    $(row).addClass('row-orange');
                                }
                            },
                            initComplete: function () {
                                // Lógica adicional después de inicializar la tabla (si es necesario)
                                $('#tbllist_filter').css('margin-bottom', '20px');
                            }
                        });
                    } else {
                        console.log('No se encontraron datos');
                    }
                },
                error: function (error) {
                    console.log('Error:', error);
                }
            });
        });
    
