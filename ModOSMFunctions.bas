Attribute VB_Name = "ModOSMFunctions"
'Set a reference to Microsoft XML, v6.0

'Two variables for caching, so the formulas don't update every recalculation
Public Const CCCacheSeconds = 120   'Nr of seconds cache, default >= 60
Public CCDict As New Scripting.Dictionary

Sub RegisterOSMFunctions()
    
    Application.MacroOptions _
        Macro:="geo_nom_reverse", _
        Description:="OSM Nominatim reverse lookup: search for city, street etc of the given latitude & longitude.", _
        Category:="Geo_vba formulas", _
        ArgumentDescriptions:=Array( _
            "Latitude (north-south) of coordinate 1, from -90 to +90", _
            "Longitude (east-west) of coordinate 1, from -180 to +180", _
            "Default: return values in one cell, False: return every element in a different cell", _
            "Comma separated list of result columns. Default: all. Other options e.g. house_number,road,neighbourhood,city,state,country,country_code")


    Application.MacroOptions _
        Macro:="geo_nom_search", _
        Description:="OSM Nominatim search: search for e.g. latitude & longitude for the street, city, etc.", _
        Category:="Geo_vba formulas", _
        ArgumentDescriptions:=Array( _
            "Address input, e.g. main street or Amsterdam,nl", _
            "Number of results, now fixed to 1", _
            "Default: return values in one cell, False: return every element in a different cell", _
            "Comma separated list of result columns. Default: all. Other options e.g. lat,lon,display_name,place_id,osm_type,osm_id,place_rank,boundingbox,class,type,importance")

End Sub

Function geo_nom_reverse(lat As Double, lng As Double, Optional OneCellResult As Boolean = True, Optional ResCols As String = "all") As Variant()
Attribute geo_nom_reverse.VB_Description = "OSM Nominatim reverse lookup: search for city, street etc of the given latitude & longitude."
Attribute geo_nom_reverse.VB_ProcData.VB_Invoke_Func = " \n20"

    Dim resArr() As Variant
    Dim Url As String
    Dim xDoc1 As New MSXML2.DOMDocument60
    'Url, correct for , as decimal setting
    Url = "https://nominatim.openstreetmap.org/reverse?format=xml&addressdetails=1&lat=" + CStr(lat) + "&lon=" + CStr(lng)
    Url = Replace(Url, ",", ".")
    Set xDoc1 = get_xml(Url)
    
    If xDoc1.parseError.ErrorCode <> 0 Then
        ReDim resArr(0, 0)
        resArr(0, 0) = xDoc1.parseError.reason
    Else
        xDoc1.SetProperty "SelectionLanguage", "XPath"
        Dim loc As MSXML2.IXMLDOMElement
        Set loc = xDoc1.SelectSingleNode("/reversegeocode/result")
        Set locDet = xDoc1.SelectSingleNode("/reversegeocode/addressparts")
        
        If loc Is Nothing Then
            'No result, check for error code, otherwise return complete XML
            Set locErr = xDoc1.SelectSingleNode("/reversegeocode/error")
            ReDim resArr(0, 0)
            If locErr Is Nothing Then
                resArr(0, 0) = xDoc.XML
            Else
                resArr(0, 0) = locErr.Text
            End If
        Else
            'OK result
            If LCase(ResCols) = "all" Then
                'All columns
                If OneCellResult Then
                    ReDim resArr(0, 0)
                    resArr(0, 0) = loc.Text
                Else
                    tempArr = Split(loc.Text, ",")
                    ReDim resArr(0, UBound(tempArr))
                    For i = 0 To UBound(tempArr)
                        resArr(0, i) = Trim(tempArr(i))
                    Next i
                End If
            Else
                'Filter returned data
                tempArr = Split(ResCols, ",")
                ReDim resArr(0, 0)
                If OneCellResult = False Then
                    ReDim resArr(0, UBound(tempArr))
                End If
                
                'locDet
                For i = 0 To UBound(tempArr)
                    Set ch = locDet.SelectSingleNode(tempArr(i))
                    
                    If OneCellResult = False Then
                        If Not ch Is Nothing Then
                            resArr(0, i) = ch.Text
                        End If
                    Else
                        If Not ch Is Nothing Then
                            resArr(0, 0) = resArr(0, 0) & ch.Text
                        End If
                        If i < UBound(tempArr) Then
                            resArr(0, 0) = resArr(0, 0) & ","
                        End If
                    End If
                Next i
            End If
        End If
    End If

    geo_nom_reverse = resArr

End Function


Function geo_nom_search(address As String, Optional NumberOfResults As Integer = 1, Optional OneCellResult As Boolean = True, Optional ResCols As String = "default") As Variant()
Attribute geo_nom_search.VB_Description = "OSM Nominatim search: search for e.g. latitude & longitude for the street, city, etc."
Attribute geo_nom_search.VB_ProcData.VB_Invoke_Func = " \n20"

    Dim resArr() As Variant
    Dim Url As String
    Dim xDoc1 As New MSXML2.DOMDocument60
    
    Url = "https://nominatim.openstreetmap.org/search?format=xml&addressdetails=1&q=" & address
    Url = Replace(Url, ",", ".")
    Set xDoc1 = get_xml(Url)
    
    If xDoc1.parseError.ErrorCode <> 0 Then
        ReDim resArr(0, 0)
        resArr(0, 0) = xDoc1.parseError.reason
    Else
        xDoc1.SetProperty "SelectionLanguage", "XPath"
        Dim loc As MSXML2.IXMLDOMElement
        
        Set locs = xDoc1.SelectNodes("/searchresults/place")
        
        If locs Is Nothing Then
            'No results, check for error code, otherwise return complete XML
            Set locErr = xDoc1.SelectSingleNode("/reversegeocode/error")
            ReDim resArr(0, 0)
            If locErr Is Nothing Then
                resArr(0, 0) = xDoc.XML
            Else
                resArr(0, 0) = locErr.Text
            End If
        Else
            'OK result
            If LCase(ResCols) = "default" Then
                'All columns
                If OneCellResult Then
                    ReDim resArr(0, 0)
                    resArr(0, 0) = locs(0).Attributes.getNamedItem("lat").Text & "," & locs(0).Attributes.getNamedItem("lon").Text
                Else
                    ReDim resArr(0, 1)
                    resArr(0, 0) = locs(0).Attributes.getNamedItem("lat").Text
                    resArr(0, 1) = locs(0).Attributes.getNamedItem("lon").Text
                End If
            Else
                'Filter returned data
                tempArr = Split(ResCols, ",")
                ReDim resArr(0, 0)
                If OneCellResult = False Then
                    ReDim resArr(0, UBound(tempArr))
                End If
                
                'locDet
                For i = 0 To UBound(tempArr)
                    Set locAtt = locs(0).Attributes.getNamedItem(tempArr(i))
                    
                    If OneCellResult = False Then
                        If Not locAtt Is Nothing Then
                            resArr(0, i) = locAtt.Text
                        End If
                    Else
                        If Not locAtt Is Nothing Then
                            resArr(0, 0) = resArr(0, 0) & locAtt.Text
                        End If
                        If i < UBound(tempArr) Then
                            resArr(0, 0) = resArr(0, 0) & ","
                        End If
                    End If
                Next i
            End If
        End If
    End If

    geo_nom_search = resArr
    
End Function

Private Function get_xml(strUrl As String) As MSXML2.DOMDocument60

'Set a reference to Microsoft XML, v6.0
Dim xDoc As New MSXML2.DOMDocument60
xDoc.async = False

'For caching, check if data already exists
IsInDict = CCDict.Exists(strUrl)
GetNewData = False
If IsInDict = True Then
    'In dictionary, check time
    CheckFailed = False
    If CCDict(strUrl) + TimeSerial(0, 0, CCCacheSeconds) < Now() Then
        CheckFailed = True
    ElseIf InStr(LCase((CCDict("DATA-" & strUrl))), "Error") > 0 Then
        CheckFailed = True
    Else
        xDoc.Load (CCDict("DATA-" & strUrl))
        If xDoc.parseError.ErrorCode <> 0 Then
            CheckFailed = True
        End If
    End If
    
    If CheckFailed Then
        'Has not been updated recently and/or forced no caching, update now
        CCDict.Remove strUrl
        CCDict.Add strUrl, Now()
        If CCDict.Exists("DATA-" & strUrl) Then CCDict.Remove "DATA-" & strUrl
        GetNewData = True
    End If
Else
    CCDict.Add strUrl, Now()
    GetNewData = True
End If

If GetNewData = True Then
    xDoc.Load (strUrl)
    CCDict.Add "DATA-" & strUrl, TempData
Else
    xDoc.Load (CCDict("DATA-" & strUrl))
End If

Set get_xml = xDoc

End Function
