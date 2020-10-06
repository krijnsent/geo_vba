Attribute VB_Name = "ModImgFunctions"
'Reference to Microsoft Windows Image Acquisition Library 2.0

Sub RegisterIMGFunctions()
    
    Application.MacroOptions _
        Macro:="geo_img_data", _
        Description:="Extract the coordinates from an image if they are present.", _
        Category:="Geo_vba formulas", _
        ArgumentDescriptions:=Array( _
            "Filename, including path, e.g. C:\temp\my_img.jpg", _
            "Default is latitude&longitude. Extra file info to retreive from WIA, comma separated. E.g. DateTime,EquipMake,EquipModel,ExifPixXDim,ExifPixYDim")

End Sub

Function geo_img_data(FileNm As String, Optional ResCols As String = "latlon") As Variant()
Attribute geo_img_data.VB_Description = "Extract the coordinates from an image if they are present."
Attribute geo_img_data.VB_ProcData.VB_Invoke_Func = " \n21"

    Dim fileName As Variant
    'TotFile = Dir(FileNm)
    Dim resArr() As Variant
    ReDim resArr(0, 0)
    
    If Dir(FileNm) = "" Then
        resArr(0, 0) = "ERROR, cannot find file"
    Else
        If GetAttr(FileNm) = vbDirectory Then
            resArr(0, 0) = "ERROR, input is a directory"
        Else
            iDateTime = ""
            LatDec = 0
            LngDec = 0
    
            On Error Resume Next
            Set ImgFile = New WIA.ImageFile
            ImgFile.LoadFile (FileNm)
            On Error GoTo 0
            
            Set iLat = Nothing
            iLatRef = ""
            Set iLng = Nothing
            iLngRef = ""
            On Error Resume Next
            Set iLat = ImgFile.Properties("GpsLatitude").value
            Set iLng = ImgFile.Properties("GpsLongitude").value
            iLatRef = ImgFile.Properties("GpsLatitudeRef").value
            iLngRef = ImgFile.Properties("GpsLongitudeRef").value
            On Error GoTo 0

            If Not iLat Is Nothing Then
                LatDec = iLat(1) + iLat(2) / 60 + iLat(3) / 3600
                If iLatRef = "S" Then LatDec = LatDec * -1
            Else
                LatDec = 0
            End If
            If Not iLng Is Nothing Then
                LngDec = iLng(1) + iLng(2) / 60 + iLng(3) / 3600
                If iLngRef = "W" Then LngDec = LngDec * -1
            Else
                LngDec = 0
            End If
            
            ReDim resArr(0, 1)
            resArr(0, 0) = LatDec
            resArr(0, 1) = LngDec
            
            If ResCols <> "latlon" Then
                tempArr = Split(ResCols, ",")
                ReDim Preserve resArr(0, 2 + UBound(tempArr))
                For i = 0 To UBound(tempArr)
                    PrpVal = "-"
                    On Error Resume Next
                    PrpVal = ImgFile.Properties(tempArr(i)).value
                    On Error GoTo 0
                    resArr(0, 2 + i) = PrpVal
                Next i
            End If
            
        End If
    End If
    'GetAttr(FileNm)
        
    geo_img_data = resArr

'    If TotFile = "" Then
'        ReDim resArr(0, 0)
'        resArr(0, 0) = "No/unknown file"
'        geo_img_data = resArr()
'    Else
'        iDateTime = ""
'        LatDec = 0
'        LngDec = 0
'
'        On Error Resume Next
'        Set ImgFile = New WIA.ImageFile
'        ImgFile.LoadFile (FileNm)
'        iDateTime = ImgFile.Properties("DateTime")
'        iLat = ImgFile.Properties("GpsLatitude")
'        iLatRef = ImgFile.Properties("GpsLatitudeRef")
'        iLng = ImgFile.Properties("GpsLongitude")
'        iLngRef = ImgFile.Properties("GpsLongitudeRef")
'        On Error GoTo 0
'        If Not IsEmpty(iLat) Then
'            LatDec = iLat(1) + iLat(2) / 60 + iLat(3) / 3600
'            If iLatRef = "S" Then LatDec = LatDec * -1
'        Else
'            LatDec = 0
'        End If
'        If Not IsEmpty(iLng) Then
'            LngDec = iLng(1) + iLng(2) / 60 + iLng(3) / 3600
'            If iLngRef = "W" Then LngDec = LngDec * -1
'        Else
'            LngDec = 0
'        End If
'
'
'
'
'    End If


End Function


Sub GetGPSData(FileNm As String)


    Dim fileName As Variant
    TotFile = Dir(FileNm, vbNormal)
    Dim resArr() As Variant
    
    'Reference to Microsoft Windows Image Acquisition Library 2.0
    Set ImgFile = New WIA.ImageFile
    ImgFile.LoadFile (TotFile)

'        Rw = Cells(Rows.Count, 1).End(xlUp).Offset(1, 0).Row
'        For Each P In ImgFile.Properties
'            Debug.Print P.Name
'        Next P
'
'        Worksheets("Src").Cells(Rw, 1).value = strFolder
'        Worksheets("Src").Cells(Rw, 2).value = fileName
'
'        On Error Resume Next    ' some of the pictures do not have this data
'        Worksheets("Src").Cells(Rw, 3).value = ImgFile.Properties("DateTime")
'        On Error GoTo 0
'
'        If UCase(right(fileName, 3)) = "JPG" Then
'            'Images only
'            On Error Resume Next
'            iLat = ImgFile.Properties("GpsLatitude")
'            iLatRef = ImgFile.Properties("GpsLatitudeRef")
'            iLng = ImgFile.Properties("GpsLongitude")
'            iLngRef = ImgFile.Properties("GpsLongitudeRef")
'            On Error GoTo 0
'            If Not IsEmpty(iLat) Then
'                LatDec = iLat(1) + iLat(2) / 60 + iLat(3) / 3600
'                If iLatRef = "S" Then LatDec = LatDec * -1
'            Else
'                LatDec = 0
'            End If
'            If Not IsEmpty(iLng) Then
'                LngDec = iLng(1) + iLng(2) / 60 + iLng(3) / 3600
'                If iLngRef = "W" Then LngDec = LngDec * -1
'            Else
'                LngDec = 0
'            End If
'            Worksheets("Src").Cells(Rw, 4).value = LatDec
'            Worksheets("Src").Cells(Rw, 5).value = LngDec
'        End If
'
'
'        'Set the fileName to the next file
'        fileName = Dir
'    Wend

End Sub

