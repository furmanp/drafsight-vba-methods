Attribute VB_Name = "DraftSight_Functions"
'Module containing functions that to automate DraftSight operations based on MS Excel input
'Coordinates shall be provided in form of an array (X,Y)

Option Explicit
Option Base 1     'Set initial index of all arrays to 1

Sub xls_ds_draw_polyline(coordinates() As Double, layer_name As String, layer_colour As Integer, polyline_closed As Boolean)

    Dim dsApp As DraftSight.Application
    Dim dsDoc As DraftSight.Document
    Dim dsModel As DraftSight.Model
    Dim dsSketchManager As DraftSight.SketchManager
    Dim dsPolyline As DraftSight.Polyline
    
    
    Set dsApp = GetObject(, "DraftSight.Application")       'Connect to DraftSight
    dsApp.AbortRunningCommand                               'Abort commands currently running in DraftSight to avoid nested commands
    Set dsDoc = dsApp.GetActiveDocument()                   'Get active document
    
    If Not dsDoc Is Nothing Then
        Set dsModel = dsDoc.GetModel()                      'Get model space
        Set dsSketchManager = dsModel.GetSketchManager()    'Get Sketch Manager
        Call ds_create_layer(layer_name, layer_colour, dsDoc)
        Set dsPolyline = dsSketchManager.InsertPolyline2D(coordinates, polyline_closed)
    Else
        MsgBox "There are no open documents in DraftSight."
    End If
    
End Sub

Sub ds_draw_curve(txt_curve As String, dsDoc As DraftSight.Document, curveID As Long)
    'Locally create objects that are needed to execute the function. Only the active DraftSight Document is provided.
    'Curve coordinates shall be provided in form of four consecutive numbers X1, Y1, X2, Y2 separated with commas.
    Dim dsModel As DraftSight.Model
    Dim dsSketchManager As DraftSight.SketchManager
    Dim dsLine As DraftSight.line
    

    Set dsModel = dsDoc.GetModel()
    Set dsSketchManager = dsModel.GetSketchManager()
    
    Dim a_temp As Variant
    
    a_temp = Split(txt_curve, ",")
    
    Dim x1 As Double, x2 As Double, y1 As Double, y2 As Double
        
    If UBound(a_temp) = 5 Then                   'checking it is two points
    
        x1 = a_temp(0)
        y1 = a_temp(2)
        
        x2 = a_temp(3)
        y2 = a_temp(5)
        
        Set dsLine = dsSketchManager.InsertLine(x1, y1, 0#, x2, y2, 0#)
        Call ds_add_custom_data(dsLine, curveID)
              
    End If
End Sub

Sub ds_create_layer(layerName As String, layerColor As Integer, dsDoc As DraftSight.Document)
'Create DraftSight layers defining its name, color and add it to currently opened document in DraftSight.

    Dim dsLayerManager As DraftSight.LayerManager
    Dim dsColor As DraftSight.Color
    Dim dsLayer As DraftSight.Layer
    
    Set dsLayerManager = dsDoc.GetLayerManager
    dsLayerManager.CreateLayer layerName, dsLayer, dsCreateObjectResult_Success
    
    Set dsColor = dsLayer.Color
    dsColor.SetNamedColor (layerColor)
    dsLayer.Color = dsColor
    dsLayer.Activate                    'make newly created layer active

End Sub

Sub ds_add_custom_data(ByVal dsLine As DraftSight.line, curveID As Long)
        
    Dim applicationName As String
    applicationName = "DataApp"
    
    Dim markerForString As Long
    Dim markerForInt16 As Long
    markerForString = 1000
    markerForInt16 = 1070
    
    'Get custom data for the Curve
    Dim dsCustomData As DraftSight.CustomData
    Set dsCustomData = dsLine.GetCustomData(applicationName)
    'Clear existing custom data
    dsCustomData.Delete
    'Get the index
    Dim index As Long
    index = dsCustomData.GetDataCount()
    
    'Add a description of the Curve as a string value to the custom data
    dsCustomData.InsertStringData index, markerForString, "Line"
    'Get the next index
    index = dsCustomData.GetDataCount()
    
    'Add custom data section to custom data
    Dim dsInnerCustomData As DraftSight.CustomData
    Set dsInnerCustomData = dsCustomData.InsertCustomData(index)
    'Get the next index
    index = dsInnerCustomData.GetDataCount()
    
    'Add the layer name of Curve as layer name data to custom data
    dsInnerCustomData.InsertLayerName index, dsLine.Layer
    'Get the next index
    index = dsInnerCustomData.GetDataCount()
    
    
    'Add the ID number of the Curve to custom data
    dsInnerCustomData.InsertInteger16Data index, markerForInt16, curveID
    'Set custom data
    dsLine.SetCustomData applicationName, dsCustomData
    
    PrintCustomDataInfo dsLine.GetCustomData(applicationName)
End Sub

Sub PrintCustomDataInfo(ByVal dsCustomData As CustomData)
    'Get custom data count
    Dim count As Long
    count = dsCustomData.GetDataCount()
    Dim index As Long
    For index = 0 To count - 1
        'Get custom data type
        Dim dataType As dsCustomDataType_e
        dsCustomData.GetDataType index, dataType
        'Get custom data marker
        Dim marker As Long
        dsCustomData.GetDataMarker index, marker
        Select Case dataType
        Case dsCustomDataType_e.dsCustomDataType_BinaryData
            If True Then
                'Get binary data from custom data
                Dim binaryArray As Variant
                binaryArray = dsCustomData.GetByteData(index)
                       
                Dim binaryDataContent As String
                binaryDataContent = ""
                     
                If IsEmpty(binaryArray) Then
                    binaryDataContent = "Empty"
                Else
                    Dim j As Long
                    For j = LBound(binaryArray) To UBound(binaryArray)
                        binaryDataContent = binaryDataContent + CStr(binaryArray(j)) & ","
                    Next j
                End If
                'Print custom data index, data type, marker, and binary value
                PrintCustomDataElement index, dataType, marker, binaryDataContent
            End If
        Case dsCustomDataType_e.dsCustomDataType_CustomData
            If True Then
                'Get inner custom data
                Dim dsGetCustomData As DraftSight.CustomData
                Set dsGetCustomData = Nothing
                dsCustomData.GetCustomData index, dsGetCustomData
                PrintCustomDataInfo dsGetCustomData
            End If
        Case dsCustomDataType_e.dsCustomDataType_Double
            If True Then
                'Get double value from custom data
                Dim doubleValue As Double
                dsCustomData.GetDoubleData index, doubleValue
                'Print custom data index, data type, marker and double value
                PrintCustomDataElement index, dataType, marker, doubleValue
            End If
        Case dsCustomDataType_e.dsCustomDataType_Handle
            If True Then
                'Get handle value from custom data
                Dim handle As String
                handle = dsCustomData.GetHandleData(index)
                'Print custom data index, data type, marker, and handle value
                PrintCustomDataElement index, dataType, marker, handle
            End If
        Case dsCustomDataType_e.dsCustomDataType_Integer16
            If True Then
                Dim int16Value As Long
                dsCustomData.GetInteger16Data index, int16Value
                'Print custom data index, data type, marker, and Int16 value
                PrintCustomDataElement index, dataType, marker, int16Value
            End If
        Case dsCustomDataType_e.dsCustomDataType_Integer32
            If True Then
                Dim int32Value As Long
                dsCustomData.GetInteger32Data index, int32Value
                'Print custom data index, data type, marker, and Int32 value
                PrintCustomDataElement index, dataType, marker, int32Value
            End If
        Case dsCustomDataType_e.dsCustomDataType_LayerName
            If True Then
                'Get layer name from custom data
                Dim layerName As String
                dsCustomData.GetLayerName index, layerName
                'Print custom data index, data type, marker, and layer name value
                PrintCustomDataElement index, dataType, marker, layerName
            End If
        Case dsCustomDataType_e.dsCustomDataType_Point
            If True Then
                'Get point coordinates from custom data
                Dim x As Double, y As Double, z As Double
                dsCustomData.GetPointData index, x, y, z
                'Print custom data index, data type, marker, and point values
                Dim pointCoordinates As String
                pointCoordinates = x & "," & y & "," & z
                PrintCustomDataElement index, dataType, marker, pointCoordinates
            End If
        Case dsCustomDataType_e.dsCustomDataType_String
            If True Then
                'Get string value from custom data
                Dim stringValue As String
                dsCustomData.GetStringData index, stringValue
                'Print custom data index, data type, marker, and string value
                PrintCustomDataElement index, dataType, marker, stringValue
            End If
        Case dsCustomDataType_e.dsCustomDataType_Unknown
            If True Then
                'Print custom data index, data type, marker and value
                PrintCustomDataElement index, dataType, marker, "Unknown value"
            End If
        Case Else
        End Select
    Next
    Debug.Print ("")
End Sub


