Attribute VB_Name = "Servicio"
Sub crearQr()
    
    Dim http As New MSXML2.ServerXMLHTTP60
    Dim url As String
    Dim arr() As Byte
    
    url = "https://chart.googleapis.com/chart?chs=150x150&cht=qr&chl=mi_contenido"
    
    With http
        .Open "GET", url, False
        .send
        
'        Debug.Print .responseBody
        arr = .responseBody
        
    End With
    
    
    Rem ---------------------------
    
    Dim numFile As Byte
    Dim pathImage As String
    
    numFile = FreeFile
    
    pathImage = "C:\Users\JHONY\Desktop\mi_imagen.png"
    
    Open pathImage For Binary Access Write As #numFile
        Put #numFile, 1, arr
    Close #numFile
    
    MsgBox "Se ha descargado el qr"
    
End Sub
