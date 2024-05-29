$(document).ready(function () {
	// Función para el botón de búsqueda
	$('.findBtn').click(function (e) {
		e.preventDefault(); // Evitar que el formulario se envíe automáticamente
		var consecutivo = $('#consecutivo').val();

		// Extraer las partes del consecutivo
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
						confirmButtonText: 'OK'
					});
				} else {
					console.log(data); // Verifica los datos recibidos en la consola del navegador
					Swal.fire({
						icon: 'success',
						title: '¡Encontrados!',
						text: 'Los datos se han encontrado correctamente.',
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
					$('#producto').val(data[0].PRODUCTO + "-" + data[0].EVENTO);

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

				if (consecutivo === "0000") {

					Swal.fire({
						icon: 'success',
						title: '¡Registro genérico sin consecutivo!',
						text: 'Bienvenido, puedes proceder a generar tu reproceso sin consecutivo',
						confirmButtonText: 'OK'
					});


				} else if (consecutivo !== null) {
					Swal.fire({
						icon: 'error',
						title: '¡Consecutivo no existe!',
						text: 'Los datos no se han encontrado correctamente, no puedes seguir con el registro del reproceso.',
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
