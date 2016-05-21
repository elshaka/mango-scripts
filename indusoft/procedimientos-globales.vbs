Sub IniciarSesion(http, usuario, contrasena)
    Dim url, parameters
    $Trace("Iniciando sesión en " & $server_url & " - " & Now)
    url = $server_url & "sessions/"
    parameters = "<user><login>" & usuario & "</login><password>" & contrasena & "</password></user>"
    http.Open "POST", url, False
    http.setRequestHeader "Content-Type", "application/xml"
    http.setRequestHeader "Accept", "application/xml"
    http.Send(parameters)
End Sub

Sub GuardarConsumo(http, codigo_orden, num_batch, id_balanza, num_tolva, cantidad)
    Dim url, parameters, amount
    $Trace("Guardando consumo " & Now)
    url = $server_url & "orders/generate_consumption"
    amount = $Format("###.##", cantidad)
    parameters = "<consumption><order_code>"&codigo_orden&"</order_code>" & _
	         "<batch_number>"&num_batch&"</batch_number><scale_id>"&id_balanza&"</scale_id>" & _
	         "<hopper_number>"&num_tolva&"</hopper_number><amount>"&amount&"</amount></consumption>"
    http.Open "POST", url, False
    http.setRequestHeader "Content-Type", "application/xml"
    http.setRequestHeader "Accept", "application/xml"
    http.Send(parameters)	
End Sub

Function GuardarNoPesados(http, codigo_orden, num_batch)
    Dim url, parameters
    $Trace("Guardando no pesados" & Now)
    url = $server_url & "orders/generate_not_weighed_consumptions"
    parameters = "<order_batch><order_code>"&codigo_orden&"</order_code>" & _
	         "<batch_number>"&num_batch&"</batch_number></order_batch>"
    http.Open "POST", url, False
    http.setRequestHeader "Content-Type", "application/xml"
    http.setRequestHeader "Accept", "application/xml"
    http.Send(parameters)
    GuardarNoPesados = Eval(http.responseXML.getElementsByTagName("success")(0).text)	
End Function

Function GuardoConsumo(http, codigo_orden, num_batch, id_balanza, num_tolva)
    Dim url, parameters
    $Trace("Verificando consumo" & Now)
    url = $server_url & "orders/consumption_exists?"
    parameters = "order_code="&codigo_orden&"&" & _
                 "batch_number="&num_batch&"&" & _
                 "scale_id="&id_balanza&"&" & _
                 "hopper_number=" & num_tolva    
    http.Open "GET", url & parameters, False
    http.setRequestHeader "Content-Type", "application/xml"
    http.setRequestHeader "Accept", "application/xml"
    http.Send
    GuardoConsumo = Eval(http.responseXML.getElementsByTagName("exists")(0).text)
End Function

Function CerrarOrden(http, codigo_orden)
    Dim url, parameters
    $Trace("Cerrando orden " & Now)
    url = $server_url & "orders/close"
    parameters = "<order_code>"&codigo_orden&"</order_code>"
    http.Open "POST", url, False
    http.setRequestHeader "Content-Type", "application/xml"
    http.setRequestHeader "Accept", "application/xml"
    http.Send(parameters)
    CerrarOrden = Eval(http.responseXML.getElementsByTagName("closed")(0).text)
End Function

Sub ListarTolvas(http)
    Dim url, parameters
    '$Trace("Listando contenido de tolvas" & " - " & Now)
    url = $server_url & "scales/hoppers_ingredients/"
    http.Open "GET", url, False
    http.setRequestHeader "Content-Type", "application/xml"
    http.setRequestHeader "Accept", "application/xml"
    http.Send
  
    Dim scale, scale_id, hopper, numero
    For Each scale In http.responseXML.SelectNodes("scales/scale")
    	scale_id = CInt(scale.SelectSingleNode("id").text)
	
    	$balanza[scale_id-1].Peso_min = scale.SelectSingleNode("minimum-weight").text
	   	$balanza[scale_id-1].Peso_max = scale.SelectSingleNode("maximum-weight").text
	  	$balanza[scale_id-1].Nombre = scale.SelectSingleNode("name").text

		For Each hopper In scale.SelectNodes("hoppers/hopper")
			numero = CInt(hopper.SelectsingleNode("number").text)

	    	Select Case scale_id
    		Case 1
		    	$Parametro_Dosf[numero-1].Nombre = hopper.SelectsingleNode("name").text
		        $Parametro_Dosf[numero-1].Codigo_Mat_Prim = hopper.SelectsingleNode("ingredient-code").text
		        $Parametro_Dosf[numero-1].Nombre_Mat_Prim = hopper.SelectsingleNode("ingredient-name").text
		    Case 2
		    	$Parametro_Liq[numero-1].Nombre = hopper.SelectsingleNode("name").text
		        $Parametro_Liq[numero-1].Codigo_Mat_Prim = hopper.SelectsingleNode("ingredient-code").text
		        $Parametro_Liq[numero-1].Nombre_Mat_Prim = hopper.SelectsingleNode("ingredient-name").text
		    End Select
		Next
    Next
End Sub


Sub ListarOrdenes(http)
    Dim url, parameters
    $Trace("Listando ordenes de producción" & " - " & Now)
    url = $server_url & "orders/open"
    http.Open "GET", url, False
    http.setRequestHeader "Content-Type", "application/xml"
    http.setRequestHeader "Accept", "application/xml"
    http.Send
  
    Dim order, index
    index = 0
    For Each order In http.responseXML.SelectNodes("orders/order")
		$Orden_Produccion[Index].Nro_orden = order.SelectsingleNode("code").text
		$Orden_Produccion[Index].cliente = order.SelectsingleNode("client-name").text
		$Orden_Produccion[Index].Cod_Receta = order.SelectsingleNode("recipe-code").text
		$Orden_Produccion[Index].Nombre_Receta = order.SelectsingleNode("recipe-name").text
		$Orden_Produccion[Index].B_Prog = order.SelectsingleNode("prog-batches").text

		Index = Index + 1
		If Index + 1 > $Tag_general.cantidad_ordenes Then
			Exit For
		End If
    Next
End Sub

Sub ValidarOrden(http, order_code)
    Dim url, parameters
    $Trace("Validando orden " & order_code & " - " & Now)
    url = $server_url & "orders/validate?order_code=" & order_code
    http.Open "GET", url, False
    http.setRequestHeader "Content-Type", "application/xml"
    http.setRequestHeader "Accept", "application/xml"
    http.Send

  	Dim validation, msg, scale, scale_id, index, numero, hopper, parameter, parameter_type
  	
  	For index = 0 To 21
  		$Parametro_Dosf[index].Dosis_Aux = 0
  	Next

  	For index = 0 To 6
  		$Parametro_Liq[index].Dosis_Aux = 0
  	Next
  	  	  	
  	Set validation = http.responseXML.SelectsingleNode("order-validation")

  	$tag_general.orden_valida = Eval(validation.SelectsingleNode("valid").text)
  	
  	If Not $Tag_general.orden_valida Then
  		msg = "Las siguientes materias primas no fueron encontradas en ninguna tolva:"
  		Dim ingredient
  		For Each ingredient In validation.SelectNodes("missing-ingredient-names/missing-ingredient-name")
  			msg = msg & vbCrLf & "- " & ingredient.text
  		Next
  		MsgBox msg
  	Else
  		For Each scale In validation.selectNodes("scale-amounts/scale-amount")
  			scale_id = CInt(scale.SelectSingleNode("scale-id").text)

  			For Each hopper In scale.SelectNodes("hoppers/hopper")
				numero = CInt(hopper.SelectsingleNode("number").text)
	
		    	Select Case scale_id
	    		Case 1
			    	$Parametro_Dosf[numero-1].Dosis_Aux = hopper.SelectsingleNode("amount").text
			    Case 2
			    	$Parametro_Liq[numero-1].Dosis_Aux = hopper.SelectsingleNode("amount").text
			    End Select
			Next
  		Next
  		
  		For Each parameter In validation.selectNodes("parameters/parameter")
  			parameter_type = CInt(parameter.SelectSingleNode("type").text)
  			
  			$Parametros_Aux[parameter_type - 1] = parameter.selectsingleNode("value").text
  		
  		Next
  	End If
End Sub
