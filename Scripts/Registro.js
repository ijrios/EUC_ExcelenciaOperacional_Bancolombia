
        $(document).ready(function () {

            // Manejar el clic del botón de cerrar sesión
            $('#btnCerrarSesion').click(function () {
                // Redirigir al usuario a la acción de cerrar sesión en el controlador
                window.location.href = '/Home/Index';
            });

            $('.deshabilitarBtn').click(function (event) {
                event.preventDefault(); // Evita que se envíe el formulario

                var valorConsecutivo = $('#consecutivo').val();
                if (!isNaN(valorConsecutivo) && valorConsecutivo !== '') {
                    $('#consecutivo').val(null);
                }
                // Colocar 'Genérico' como placeholder
                $('#consecutivo').attr('placeholder', 'Genérico');
                // Deshabilitar el campo
                $('#consecutivo').prop('disabled', true);

                Swal.fire({
                    icon: 'success',
                    title: '¡Registro genérico sin consecutivo!',
                    text: 'Bienvenido, puedes proceder a generar tu reproceso sin consecutivo',
                    confirmButtonColor: '#00c288',
                    confirmButtonText: 'OK'
                });



            });

        });


        document.addEventListener('DOMContentLoaded', function () {

            const areaSelect = document.getElementById('area');
            if (areaSelect) {
                areaSelect.addEventListener('change', function () {
                    const chartArea = areaSelect.value;
                    $.ajax({
                        url: '/Home/GetTipoErrores',
                        type: 'GET',
                        data: { area: chartArea },
                        success: function (data) {
                            const tipoSelect = $('#tipoError');
                            tipoSelect.empty();  // Vaciar las opciones actuales
                            tipoSelect.append(new Option("Selecciona el tipo", "", true, true));  // Agregar la opción por defecto
                            data.forEach(function (tipoError) {
                                tipoSelect.append(new Option(tipoError.Tipo));
                            });
                        },
                        error: function (error) {
                            console.log("Error fetching causas:", error);
                        }
                    });

                });

            }

            const tipoerrorSelect = document.getElementById('tipoError');
            if (areaSelect && tipoerrorSelect) {
                tipoerrorSelect.addEventListener('change', function () {
                    const chartTipo = tipoerrorSelect.value;
                    const chartArea = areaSelect.value;
                    $.ajax({
                        url: '/Home/GetCausaErrores',
                        type: 'GET',
                        data: { tipoerror: chartTipo, area: chartArea },
                        success: function (data) {
                            const causaSelect = $('#causa');
                            causaSelect.empty();  // Vaciar las opciones actuales
                            causaSelect.append(new Option("Selecciona la causa", "", true, true));  // Agregar la opción por defecto
                            data.forEach(function (causa) {
                                causaSelect.append(new Option(causa.Causa));
                            });
                        },
                        error: function (error) {
                            console.log("Error fetching causas:", error);
                        }
                    });
                });
            }

           
            $.ajax({
                url: '/Home/GetMoneda',
                type: 'GET',
                success: function (data) {
                    const monedaSelect = $('#moneda');
                    monedaSelect.empty();  // Vaciar las opciones actuales
                    monedaSelect.append(new Option("Selecciona la moneda", "", true, true));  // Agregar la opción por defecto
                    data.forEach(function (moneda) {
                        monedaSelect.append(new Option(moneda.Monedas, moneda.Id));
                    });
                },
                error: function (error) {
                    console.log("Error fetching moeda:", error);
                }
            });

        });


        $(document).ready(function () {

            // Función para el botón de búsqueda
            $('.findBtn').click(function (e) {
                e.preventDefault();
                var consecutivo = $('#consecutivo').val();

                var fecha = consecutivo.substring(0, 8); // Los primeros 8 caracteres representan la fecha
                var code = consecutivo.substring(8, 11); // Los caracteres del 9 al 11 representan el código
                var consecutivofinal = consecutivo.substring(11); // El resto representa el consecutivo final

                fecha = fecha.replace(/-/g, ''); // Eliminar los guione



                $.ajax({
                    url: '/Home/Buscar',
                    type: 'POST',
                    data: { consecutivo: consecutivofinal, fecha, code },
                    success: function (data) {
                        // Aquí puedes manejar los datos recibidos y mostrarlos en los campos correspondientes del formulario
                        // Por ejemplo:
                        if (data === null || data.length === 0) {
                            console.log(data); // Verifica los datos recibidos en la consola del navegador
                            Swal.fire({
                                icon: 'Error',
                                title: '¡Datos no Encontrados!',
                                text: 'Los datos se no han encontrado correctamente, no puede seguir con la operación a menos que asigne 0000 en el consecutivo.',
                                confirmButtonColor: '#00c288',
                                cancelButtonColor: '#7900D3',
                                confirmButtonText: 'OK'
                            });
                        } else {
                            console.log(data); // Verifica los datos recibidos en la consola del navegador
                            Swal.fire({
                                icon: 'success',
                                title: '¡Encontrados!',
                                text: 'Los datos se han encontrado correctamente.',
                                confirmButtonColor: '#00c288',
                                cancelButtonColor: '#7900D3',
                                confirmButtonText: 'OK'
                            });

                            var moneda = '';
                            switch (data[0].MONEDA) {
                                case 1:
                                    moneda = 'USD'
                                    break;
                                case 3:
                                    moneda = 'GBP'
                                    break;
                                case 5:
                                    moneda = 'CHF'
                                    break;
                                case 7:
                                    moneda = 'SEK'
                                    break;
                                case 8:
                                    moneda = 'DKK'
                                    break;
                                case 11:
                                    moneda = 'JPY'
                                    break;
                                case 13:
                                    moneda = 'VEB'
                                    break;
                                case 14:
                                    moneda = 'ATS'
                                    break;
                                case 15:
                                    moneda = 'CAD'
                                    break;
                                case 16:
                                    moneda = 'EUR'
                                    break;
                                case 17:
                                    moneda = 'PTE'
                                    break;
                                case 18:
                                    moneda = 'AUD'
                                    break;
                                case 19:
                                    moneda = 'MXN'
                                    break;
                                case 21:
                                    moneda = 'NZD'
                                    break;
                                case 22:
                                    moneda = 'BRL'
                                    break;
                            }



                            $('#nit').val(data[0].NIT);
                            $('#valor').val(data[0].VALOR);
                            $('#moneda').val(moneda);
                            $('#cliente').val(data[0].NOMBRE);
                            

                             $.ajax({
                                 url: '/Home/BuscarProductoEvento',
                                 type: 'POST',
                                 data: { producto: data[0].PRODUCTO, evento: data[0].EVENTO},
                                        success: function (datae) {
                                            if (datae !== null) {
                                                $('#producto').val(datae[0].descripcion);
                                            } else {
                                                $('#producto').val(data[0].PRODUCTO + "-" + data[0].EVENTO);
                                            }
                                        },
                                        error: function (error) {
                                            console.log("Error fetching producto:", error);
                                            $('#producto').val(data[0].PRODUCTO + "-"+data[0].EVENTO);
                                        }
                                    });


                            var fecha = (data[0].FECREC);
                            fecha = fecha.toString()

                            // Separar la cadena en año, mes y día
                            var anio = fecha.slice(0, 4);
                            var mes = fecha.slice(4, 6);
                            var dia = fecha.slice(6, 8);

                            // Construir la cadena de fecha en formato yyyy-MM-dd
                            var fechaFormateada = anio + '-' + mes + '-' + dia;

                            // Asignar la fecha formateada al campo de fecha en el formulario
                            $('#fecha').val(fechaFormateada);
                        }



                    },
                    error: function () {

                        if (consecutivo === "null") {

                            Swal.fire({
                                icon: 'success',
                                title: '¡Registro genérico sin consecutivo!',
                                text: 'Bienvenido, puedes proceder a generar tu reproceso sin consecutivo',
                                confirmButtonColor: '#00c288',
                                confirmButtonText: 'OK'
                            });


                        } else if (consecutivo !== null) {
                            Swal.fire({
                                icon: 'error',
                                title: '¡Consecutivo no existe!',
                                text: 'Los datos no se han encontrado correctamente, no puedes seguir con el registro del reproceso.',
                                confirmButtonColor: '#00c288',
                                confirmButtonText: 'OK'
                            });
                            $('#consecutivo').val(null);
                            // Deshabilitar todas las casillas de entrada en el formulario
                            $('form input').prop('disabled', true);
                            $('form select').prop('disabled', true);
                        }

                    }
                });
            });

        });
        $(document).ready(function () {
            $('#btnGuardarDatos').click(function (event) {
                event.preventDefault(); 
                // Obtener los valores de los campos del formulario
                var consecutivo = $('#consecutivo').val();


                if (consecutivo === '' || consecutivo === null) {

                var usuario = "@ViewBag.WindowsUsername";
                var fechaOp = $('#fecha').val().toString();
                var fechaReg = obtenerFechaActual().toString();
                var consecutivoFin = "000000000000000";
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
                // Obtener el mes numérico y convertirlo al nombre del mes
                var mesNumerico = obtenerMesNumerico(fechaOp);
                var nombreMes = obtenerNombreMes(mesNumerico);
                var fecha = new Date();
                var año = fecha.getFullYear();
                var queja = $('#queja').val().toString();

                    if (fechaOp === '' || nit === '' || moneda === '' || isNaN(valor) || cliente === '' || producto === '' || responsable === '' || perdida === '' || impacto === '' || causa === '' || descripcion === '' || area === '' || tipoError === '' || queja === '') {
                        Swal.fire({
                            icon: 'error',
                            title: '¡Campos Vaciós!',
                            confirmButtonColor: '#00c288',
                            cancelButtonColor: '#7900D3',
                            text: 'Todos los campos deben estar diligenciados.',
                            confirmButtonText: 'OK'
                        });
                    } else {

                        // Enviar los datos al controlador mediante AJAX
                        $.ajax({
                            url: '/Home/InsertarReproceso',
                            type: 'POST',
                            data: {
                                FechaOp: fechaOp,
                                FechaReg: fechaReg,
                                Consecutivo: consecutivoFin,
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
                                Mes: nombreMes,
                                Año: año,
                                QuejaCliente: queja
                            },
                            success: function (response) {
                 
                                console.log(response); // Asegúrate de que la respuesta sea un objeto JSON válido en la consola
                                if (response.success === true) {
                                    Swal.fire({
                                        icon: 'success',
                                        title: '¡Datos guardados!',
                                        text: 'Los datos se han guardado correctamente, en la vista de consultas lo puede visualizar',
                                        confirmButtonColor: '#00c288',
                                        cancelButtonColor: '#7900D3',
                                        confirmButtonText: 'OK'
                                    });

                                    $('#usuario').val('');
                                    $('#fechaOp').val('');
                                    $('#fechaReg').val('');
                                    $('#consecutivo').val('');
                                    $('#nit').val('');
                                    $('#moneda').val('');
                                    $('#valor').val('');
                                    $('#cliente').val('');
                                    $('#producto').val('');
                                    $('#responsable').val('');
                                    $('#perdida').val('');
                                    $('#impacto').val('');
                                    $('#causa').val('');
                                    $('#descripcion').val('');
                                    $('#area').val('');
                                    $('#tipoError').val('');
                                    $('#queja').val('');



                                } else {
                                    Swal.fire({
                                        icon: 'error',
                                        title: '¡Datos no guardados!',
                                        text: 'Los datos no se han guardado correctamente',
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
                                    title: '¡Datos no guardados!',
                                    confirmButtonColor: '#00c288',
                                    cancelButtonColor: '#7900D3',
                                    text: error,
                                    confirmButtonText: 'OK'
                                });
                            }
                        });
                    }

                } else {

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
                // Obtener el mes numérico y convertirlo al nombre del mes
                var mesNumerico = obtenerMesNumerico(fechaOp);
                var nombreMes = obtenerNombreMes(mesNumerico);
                var fecha = new Date();
                var año = fecha.getFullYear();
                var queja = $('#queja').val().toString();

                    if (consecutivo === '' || fechaOp === '' || nit === '' || moneda === '' || isNaN(valor) || cliente === '' || producto === '' || responsable === '' || perdida === '' || impacto === '' || causa === '' || descripcion === '' || area === '' || tipoError === '' || queja === '') {
                        Swal.fire({
                            icon: 'error',
                            title: '¡Campos Vaciós!',
                            confirmButtonColor: '#00c288',
                            cancelButtonColor: '#7900D3',
                            text: 'Todos los campos deben estar diligenciados.',
                            confirmButtonText: 'OK'
                        });
                    } else {
                        // Enviar los datos al controlador mediante AJAX
                        $.ajax({
                            url: '/Home/InsertarReproceso',
                            type: 'POST',
                            data: {
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
                                Mes: nombreMes,
                                Año: año,
                                QuejaCliente: queja
                            },
                            success: function (response) {
                                console.log(response); // Asegúrate de que la respuesta sea un objeto JSON válido en la consola
                                if (response.success === true) {
                                    Swal.fire({
                                        icon: 'success',
                                        title: '¡Datos guardados!',
                                        text: 'Los datos se han guardado correctamente, en la vista de consultas lo puede visualizar',
                                        confirmButtonColor: '#00c288',
                                        cancelButtonColor: '#7900D3',
                                        confirmButtonText: 'OK'
                                    });

                                    $('#usuario').val('');
                                    $('#fechaOp').val('');
                                    $('#fechaReg').val('');
                                    $('#consecutivo').val('');
                                    $('#nit').val('');
                                    $('#moneda').val('');
                                    $('#valor').val('');
                                    $('#cliente').val('');
                                    $('#producto').val('');
                                    $('#responsable').val('');
                                    $('#perdida').val('');
                                    $('#impacto').val('');
                                    $('#causa').val('');
                                    $('#descripcion').val('');
                                    $('#area').val('');
                                    $('#tipoError').val('');
                                    $('#queja').val('');

                                } else {
                                    Swal.fire({
                                        icon: 'error',
                                        title: '¡Datos no guardados!',
                                        text: 'Los datos no se han guardado correctamente',
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
                                    title: '¡Datos no guardados!',
                                    confirmButtonColor: '#00c288',
                                    cancelButtonColor: '#7900D3',
                                    text: error,
                                    confirmButtonText: 'OK'
                                });
                            }
                        });
                    }
                }

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

        // Función para obtener el mes numérico de una fecha en formato YYYY-MM-DD
        function obtenerMesNumerico(fecha) {
            var mes = fecha.substring(5, 7); // Extraer el mes (MM)
            return parseInt(mes, 10); // Convertir a número entero
        }

        // Función para obtener el nombre del mes a partir de su número (1 a 12)
        function obtenerNombreMes(numeroMes) {
            var nombresMeses = [
                "Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio",
                "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"
            ];
            return nombresMeses[numeroMes - 1]; // El índice comienza en 0, por eso se resta 1
        }

        $(document).ready(function () {
            // Función para obtener la lista de usuarios desde la API
            function obtenerUsuarios() {
                return $.ajax({
                    url: '/Home/ObtenerUsuario', // URL de tu API para obtener usuarios
                    dataType: 'json'
                });
            }

            // Inicializar Autocomplete en el campo de entrada
            $('#responsable').autocomplete({
                source: function (request, response) {
                    obtenerUsuarios().done(function (data) {
                        // Filtrar usuarios que coincidan con el término de búsqueda
                        var term = request.term.toLowerCase();
                        var usuariosFiltrados = data.filter(function (usuario) {
                            return usuario.Usuario.toLowerCase().includes(term);
                        });
                        // Mapear los nombres de usuario para el autocompletar
                        var nombresUsuarios = usuariosFiltrados.map(function (usuario) {
                            return usuario.Usuario;
                        });
                        response(nombresUsuarios);
                    });
                },
                minLength: 1, // Número mínimo de caracteres antes de hacer la búsqueda
                select: function (event, ui) {
                    // Manejar la selección del usuario
                    $('#responsable').val(ui.item.value);
                    return false; // Evitar que jQuery UI establezca el valor automáticamente
                }
            });


         
        });

        document.addEventListener("DOMContentLoaded", function () {
            var container = document.querySelector('.container');
            setTimeout(function () {
                container.classList.add('opened');
            }, 100);
        });
