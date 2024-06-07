
        document.addEventListener('DOMContentLoaded', function () {

            cargaraños();

            // Cargo el gráfico de Google
            google.charts.load('current', {
                'packages': ['corechart']
            });

            function cargaraños() {
                $.ajax({
                    url: '/Home/GetAño',
                    type: 'GET',
                    success: function (data) {
                        const monedaSelect = $('#año');
                        monedaSelect.empty();  // Vaciar las opciones actuales
                        monedaSelect.append(new Option("Selecciona el año", "", true, true));  // Agregar la opción por defecto
                        data.forEach(function (anus) {
                            monedaSelect.append(new Option(anus.Año));
                        });
                    },
                    error: function (error) {
                        console.log("Error fetching moeda:", error);
                    }
                });
            }
       

            google.charts.setOnLoadCallback(function () {
                var fecha = new Date();
                var año = fecha.getFullYear();
                cargarGraf(7, año);

            });

            //Función que cargar el gràfico de Google
            function cargarGraf(code, año) {
                
                google.charts.setOnLoadCallback(draw(code, año));
            }

            function draw(code, año) {

                const mesesOrden = ['Enero', 'Febrero', 'Marzo', 'Abril', 'Mayo', 'Junio', 'Julio', 'Agosto', 'Septiembre', 'Octubre', 'Noviembre', 'Diciembre'];

                fetch(`/Home/ObtenerErrorMes?code=${code}&año=${año}`)
                    .then(response => response.json())
                    .then(data => {
                        // Verificar la estructura de los datos obtenidos
                        console.log('Datos obtenidos:', data);

                        // Obtener los meses únicos y ordenarlos según mesesOrden
                        const meses = [...new Set(data.map(item => item.Mes))].sort((a, b) => mesesOrden.indexOf(a) - mesesOrden.indexOf(b));
                        console.log('Meses únicos ordenados:', meses);

                        // Obtener las áreas únicas
                        const areas = [...new Set(data.map(item => item.Area))];
                        console.log('Áreas únicas:', areas);

                        // Crear la estructura de datos para Google Charts
                        const arregloDatos = [['Mes', ...areas]];

                        // Agrupar los datos por mes y área
                        meses.forEach(mes => {
                            const row = [mes];
                            areas.forEach(area => {
                                const item = data.find(d => d.Mes === mes && d.Area === area);
                                row.push(item ? parseInt(item.Errores, 10) : 0);
                            });
                            arregloDatos.push(row);
                        });

                        console.log('Estructura de datos para Google Charts:', arregloDatos);

                        const dataTable = google.visualization.arrayToDataTable(arregloDatos);

                        const options = {
                            width: 670,
                            height: 260,
                            colors: ['#863AA5', '#00c288', '#4285F4', '#EA4335'], // Personaliza los colores
                            title: 'Consolidado de errores por Mes y Área',
                            legend: { position: 'right' },
                            hAxis: {
                                title: 'Mes'
                            },
                            vAxis: {
                                title: 'Errores'
                            }
                        };

                        const chart = new google.visualization.LineChart(document.getElementById('piechart'));
                        chart.draw(dataTable, options);
                        console.log("Datos recibidos y gráfico dibujado");
                    })
                    .catch(error => {
                        console.error('Error fetching data:', error);
                    });
            }

        });
        $(document).ready(function () {
            // Manejar el clic del botón de cerrar sesión
            $('#btnCerrarSesion').click(function () {
                // Redirigir al usuario a la acción de cerrar sesión en el controlador
                window.location.href = '/Home/Informes';
            });

            $('#findBtn').click(function () {
                const chartArea = document.getElementById("area").value;
                const chartType = document.getElementById("tipo").value;
                const selectedFecha = document.getElementById("meses").value;
                const año = document.getElementById("año").value;

                if (chartArea === "todas") {
                    cargarGrafico(7, selectedFecha, chartType, año);
                } else if (chartArea === "trade") {
                    cargarGrafico(1, selectedFecha, chartType, año);
                } else if (chartArea === "cartera") {
                    cargarGrafico(2, selectedFecha, chartType, año);
                } else if (chartArea === "balanza") {
                    cargarGrafico(3, selectedFecha, chartType, año);
                } else if (chartArea === "cv") {
                    cargarGrafico(4, selectedFecha, chartType, año);
                } else if (chartArea === "swift") {
                    cargarGrafico(6, selectedFecha, chartType, año);
                }

            });

            //Función que cargar el gràfico de Google
            function cargarGrafico(code, mes, tipo, año) {
                // Cargo el gráfico de Google
                google.charts.load('current', {
                    'packages': ['corechart']
                });
                google.charts.setOnLoadCallback(drawChart(code, mes, tipo, año));
            }


            function drawChart(code, mes, tipo, año) {

                if (tipo === "tipoerror") {
                    fetch(`/Home/ObtenerErrorTipo?code=${code}&fecha=${mes}&año=${año}`)
                        .then(response => response.json())
                        .then(data => {
                            const arregloDatos = [['Tipo', 'Errores', { role: 'style' }]];
                            const colores = ['#00c288']; // Definimos colores diferentes para cada tipo de error

                            data.forEach((item, index) => {
                                const errores = parseInt(item.Errores, 10);
                                if (!isNaN(errores)) {
                                    arregloDatos.push([item.Tipo, errores, colores[index % colores.length]]);
                                }
                            });

                            const dataTable = google.visualization.arrayToDataTable(arregloDatos);

                            const options = {
                                width: 670,
                                height: 260,
                                title: 'Consolidado de errores por Tipo',
                                legend: { position: 'none'},
                                tooltip: { isHtml: true },
                                hAxis: {
                                    title: 'Tipos de errores',
                                    ticks: arregloDatos.map(row => row[0]), // Usamos los nombres de los tipos como etiquetas del eje horizontal
                                },
                            };

                            let chart;
                            chart = new google.visualization.ColumnChart(document.getElementById('piechart'));

                            chart.draw(dataTable, options);
                            console.log("Datos recibidos y gráfico dibujado")
                        })
                        .catch(error => {
                            console.error('Error fetching data:', error);
                        });
                } else if (tipo === "errores") {

                    const mesesOrden = ['Enero', 'Febrero', 'Marzo', 'Abril', 'Mayo', 'Junio', 'Julio', 'Agosto', 'Septiembre', 'Octubre', 'Noviembre', 'Diciembre'];
                    fetch(`/Home/ObtenerErrorMes?code=${code}&año=${año}`)
                        .then(response => response.json())
                        .then(data => {
                            // Verificar la estructura de los datos obtenidos
                            console.log('Datos obtenidos:', data);

                            // Obtener los meses únicos y ordenarlos según mesesOrden
                            const meses = [...new Set(data.map(item => item.Mes))].sort((a, b) => mesesOrden.indexOf(a) - mesesOrden.indexOf(b));
                            console.log('Meses únicos ordenados:', meses);

                            // Obtener las áreas únicas
                            const areas = [...new Set(data.map(item => item.Area))];
                            console.log('Áreas únicas:', areas);

                            // Crear la estructura de datos para Google Charts
                            const arregloDatos = [['Mes', ...areas]];

                            // Agrupar los datos por mes y área
                            meses.forEach(mes => {
                                const row = [mes];
                                areas.forEach(area => {
                                    const item = data.find(d => d.Mes === mes && d.Area === area);
                                    row.push(item ? parseInt(item.Errores, 10) : 0);
                                });
                                arregloDatos.push(row);
                            });

                            console.log('Estructura de datos para Google Charts:', arregloDatos);

                            const dataTable = google.visualization.arrayToDataTable(arregloDatos);

                            const options = {
                                width: 670,
                                height: 260,
                                colors: ['#863AA5', '#00c288','#4285F4', '#EA4335'], // Personaliza los colores
                                title: 'Consolidado de errores por Mes y Área',
                                legend: { position: 'right' },
                                hAxis: {
                                    title: 'Mes'
                                },
                                vAxis: {
                                    title: 'Errores'
                                }
                            };

                            const chart = new google.visualization.LineChart(document.getElementById('piechart'));
                            chart.draw(dataTable, options);
                            console.log("Datos recibidos y gráfico dibujado");
                        })
                        .catch(error => {
                            console.error('Error fetching data:', error);
                        });
                } else if (tipo === "impacto") {
                    fetch(`/Home/ObtenerErrorImpacto?code=${code}&mes=${mes}&año=${año}`)
                        .then(response => response.json())
                        .then(data => {
                            const arregloDatos = [['Impacto', 'Errores', { role: 'style' }]];
                            const colores = ['#00c288']; // Definimos colores diferentes para cada tipo de error

                            data.forEach((item, index) => {
                                const errores = parseInt(item.Errores, 10);
                                if (!isNaN(errores)) {
                                    arregloDatos.push([item.Impacto, errores, colores[index % colores.length]]);
                                }
                            });

                            const dataTable = google.visualization.arrayToDataTable(arregloDatos);

                            const options = {
                                width: 670,
                                height: 260,
                                title: 'Consolidado de errores por Impacto',
                                legend: { position: 'none' },
                                tooltip: { isHtml: true },
                                hAxis: {
                                    title: 'Errores por Impacto',
                                    ticks: arregloDatos.map(row => row[0]), // Usamos los nombres de los tipos como etiquetas del eje horizontal
                                },
                            };

                            let chart;
                            chart = new google.visualization.ColumnChart(document.getElementById('piechart'));

                            chart.draw(dataTable, options);
                            console.log("Datos recibidos y gráfico dibujado")
                        })
                        .catch(error => {
                            console.error('Error fetching data:', error);
                        });
                } else if (tipo === "perdida") {
                    fetch(`/Home/ObtenerErrorPerdida?code=${code}&mes=${mes}&año=${año}`)
                        .then(response => response.json())
                        .then(data => {
                            const arregloDatos = [['Perdida', 'Errores', { role: 'style' }]];
                            const colores = ['#00c288']; // Definimos colores diferentes para cada tipo de error

                            data.forEach((item, index) => {
                                const errores = parseInt(item.Errores, 10);
                                if (!isNaN(errores)) {
                                    arregloDatos.push([item.Perdida, errores, colores[index % colores.length]]);
                                }
                            });

                            const dataTable = google.visualization.arrayToDataTable(arregloDatos);

                            const options = {
                                width: 670,
                                height: 260,
                                title: 'Consolidado de errores por Pérdidas',
                                legend: { position: 'none' },
                                tooltip: { isHtml: true },
                                hAxis: {
                                    title: 'Pérdidas económicas',
                                    ticks: arregloDatos.map(row => row[0]), // Usamos los nombres de los tipos como etiquetas del eje horizontal
                                },
                               
                            };

                            let chart;
                            chart = new google.visualization.ColumnChart(document.getElementById('piechart'));

                            chart.draw(dataTable, options);
                            console.log("Datos recibidos y gráfico dibujado")
                        })
                        .catch(error => {
                            console.error('Error fetching data:', error);
                        });
                } else if (tipo === "causas") {
                    fetch(`/Home/ObtenerErrorCausa?code=${code}&mes=${mes}&año=${año}`)
                        .then(response => response.json()) // Convertimos la respuesta a JSON
                        .then(data => {
                            // Seleccionamos el elemento donde se insertará la tabla
                            const tableContainer = document.getElementById('piechart');

                            // Creamos la tabla y el encabezado con un estilo CSS adicional
                            let tableHTML = '<table class="styled-table">';
                            tableHTML += '<thead><tr><th>Causa</th><th>Errores</th></tr></thead>';
                            tableHTML += '<tbody>';

                            // Iteramos sobre los datos recibidos y generamos las filas de la tabla
                            data.forEach(item => {
                                const errores = parseInt(item.Errores, 10);
                                if (!isNaN(errores)) {
                                    tableHTML += `<tr><td>${item.Causa}</td><td>${errores}</td></tr>`;
                                }
                            });

                            tableHTML += '</tbody></table>';

                            // Insertamos la tabla en el contenedor
                            tableContainer.innerHTML = tableHTML;

                            console.log("Datos recibidos y tabla generada");
                        })
                        .catch(error => {
                            // Manejamos cualquier error en la solicitud
                            console.error('Error fetching data:', error);
                        });
                } else if (tipo === "queja") {
                    fetch(`/Home/ObtenerErrorQueja?code=${code}&mes=${mes}&año=${año}`)
                        .then(response => response.json())
                        .then(data => {
                            const arregloDatos = [['Queja', 'Errores', { role: 'style' }]];
                            const colores = ['#00c288']; // Definimos colores diferentes para cada tipo de error

                            data.forEach((item, index) => {
                                const errores = parseInt(item.Errores, 10);
                                if (!isNaN(errores)) {
                                    arregloDatos.push([item.Queja, errores, colores[index % colores.length]]);

                                }
                            });

                            const dataTable = google.visualization.arrayToDataTable(arregloDatos);
                            const options = {
                                width: 670,
                                height: 260,
                                title: 'Consolidado de Quejas cliente',
                                legend: { position: 'none' },
                                tooltip: { isHtml: true },
                                hAxis: {
                                    title: 'Queja cliente',
                                    ticks: arregloDatos.map(row => row[0]), // Usamos los nombres de los tipos como etiquetas del eje horizontal
                                },

                            };

                           

                            let chart;
                            chart = new google.visualization.ColumnChart(document.getElementById('piechart'));

                            chart.draw(dataTable, options);
                            console.log("Datos recibidos y gráfico dibujado")
                        })
                        .catch(error => {
                            console.error('Error fetching data:', error);
                        });
                } else if (tipo === "tipoerrorDuo") {
                    fetch(`/Home/ObtenerErrorTipo?code=${code}&fecha=${mes}&año=${año}`)
                        .then(response => response.json()) // Convertimos la respuesta a JSON
                        .then(data => {
                            // Seleccionamos el elemento donde se insertará la tabla
                            const tableContainer = document.getElementById('piechart');

                            // Creamos la tabla y el encabezado con un estilo CSS adicional
                            let tableHTML = '<table class="styled-table">';
                            tableHTML += '<thead><tr><th>Tipo Error</th><th>Errores</th></tr></thead>';
                            tableHTML += '<tbody>';

                            // Iteramos sobre los datos recibidos y generamos las filas de la tabla
                            data.forEach(item => {
                                const errores = parseInt(item.Errores, 10);
                                if (!isNaN(errores)) {
                                    tableHTML += `<tr><td>${item.Tipo}</td><td>${errores}</td></tr>`;
                                }
                            });

                            tableHTML += '</tbody></table>';

                            // Insertamos la tabla en el contenedor
                            tableContainer.innerHTML = tableHTML;

                            console.log("Datos recibidos y tabla generada");
                        })
                        .catch(error => {
                            // Manejamos cualquier error en la solicitud
                            console.error('Error fetching data:', error);
                        });
                }

            }

        });
  
