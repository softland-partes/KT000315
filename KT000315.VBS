'* Seccion de declaracion de variables *
Dim sNameFile, sPath, sFileOrigen, sFileUri
Dim oFso
Dim vNroLot
'***************
Function On_Initialization_Transaction(sErrorMessage)
  vNroLot = transaccion.numerodelote

    With oWizard.Steps(1).Table
        sPath = .Fields("VIRT_TXTPAT").Value
        sNameFile = .Fields("VIRT_TXTNAM").Value
    End With
End Function

Function On_Finish_Transaction(sErrorMessage)

  'Variable para fso
  Dim sNewNameFile, oFileDestino, sFileUriDestino
  Dim sFilestreamOrigen, sFilestreamDestino
  sContador = 0

  Set oFso = CreateObject("scripting.filesystemObject")


  If not oFso.FileExists(sPath & sNameFile) Then
    Set sFileUri = oFso.CreateTextFile (sPath & sNameFile, true)
  End If

  Set sFileUri = oFso.GetFile(sPath & sNameFile)

  sPos = InStr(1, sNameFile, ".")
  If sPos = 0 Then
    sPos = Len(sNameFile) + 1
  End If

  sNewFile = Mid(sNameFile, 1, sPos - 1) & "_" & vNroLot & Mid(sNameFile, sPos, 100)

  If oFso.FileExists(sPath & sNewFile) Then
    oFso.DeleteFile sPath & sNewFile
  End If

  sFileOrigen = sNameFile
  Set oFso = Nothing
  Set oFso = CreateObject("scripting.filesystemObject")
  sNameFileDestino = sPath & sNewFile
  Set oFileDestino = oFso.CreateTextFile (sNameFileDestino, true)
  Set oFileDestino = oFso.GetFile(sNameFileDestino)

  Set sFilestreamOrigen = sFileUri.OpenAsTextStream(1,-2)
  Set sFilestreamDestino = oFileDestino.OpenAsTextStream (2,-2)

  Do While Not sFilestreamOrigen.AtEndOfStream
    sLineaActual = sFilestreamOrigen.ReadLine  'Lee la linea
    sNumeroDeLinea = Right ("000000" & sContador,6) 'Concatena ceros y recorta
    sLineaActual = Replace (cstr(sLineaActual), "@@@@@@", sNumeroDeLinea) 'reemplaza @ por ceros
    sFilestreamDestino.writeLine (sLineaActual)
    sContador = sContador + 1
  Loop
  sFilestreamDestino.close
  sFilestreamOrigen.close

  'Borro archivo original
  sFileUri.Delete


  Set oFso = Nothing
  Set sFileUri = Nothing
  Set sFileUriDestino = Nothing
End Function

Function On_Finish_Insert(sErrorMessage)
End Function
