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
