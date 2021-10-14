Attribute VB_Name = "ds_product_definition"
Option Explicit
Option Base 1
Sub ds_pd_draw_main()
'Subroutine imports Data from Product Definition sheet, and Cross section coordinates from Geometry Sheet. Offsets the curves by specified value
'and calls in a function that draws these polylines in DraftSight
    Dim i As Integer, j As Integer, k As Integer, s_txt As String, a_temp As Variant
    Dim curve_ID As Integer
    Dim layer_name As String
    Dim layer_colour As Integer
    Dim offset_val As Double
    Dim section_id As Integer
    Dim a_data As Variant
    Dim a_coordinates As Variant
    Dim offset_coordinates As Variant
    Dim a1D_coordinates() As Double
    Dim curve_start As Double
    Dim curve_end As Variant
    Dim polyline_closed As Boolean
    Dim use_input As Integer
    polyline_closed = True
    Sheets("Product Definition").Activate
    
    a_data = a_make(".pd", 1, 1, 0, 0)
    a_data = a_row0(a_data)

    For i = 1 To UBound(a_data)
        curve_ID = a_look("Alignment", a_data, i) * 10000 + a_look("Geometry", a_data, i) * 1000 + a_look("Section", a_data, i) * 10 + a_look("Segment", a_data, i)
        layer_name = a_look("Name", a_data, i)
        offset_val = a_look("offset", a_data, i)
        layer_colour = a_look("Colour", a_data, i)
        section_id = a_look("Section", a_data, i)
        use_input = a_look("Use", a_data, i)
        
        If IsEmpty(a_look("L1", a_data, i)) Then
            curve_start = 0
        Else
            curve_start = a_look("L1", a_data, i)
        End If
            
        If IsEmpty(a_look("L2", a_data, i)) Then
            curve_end = "_Full_Length"
        Else
            curve_end = a_look("L2", a_data, i)
        End If
                
        s_txt = ".c" & curve_ID
        
        Sheets("Geometry").Activate
        a_coordinates = a_make(s_txt, 1, 1, 0, 3)
        If use_input = 1 Then
            If Not ((a_coordinates(1, 1) = a_coordinates(UBound(a_coordinates, 1), 1)) _
                And (a_coordinates(1, 2) = a_coordinates(UBound(a_coordinates, 1), 2))) Then
                polyline_closed = False
            End If
            
            offset_coordinates = ds_offset_polyline(a_coordinates, -offset_val, curve_start, curve_end)
                    
            'Redimensioning 2D array into 1D Array and parsing them from Variant to Double
            ReDim a1D_coordinates(1 To UBound(offset_coordinates) * 2)
            For j = 1 To UBound(offset_coordinates)
                a1D_coordinates((j - 1) * 2 + 1) = CDbl(offset_coordinates(j, 1))
                a1D_coordinates((j - 1) * 2 + 2) = CDbl(offset_coordinates(j, 2))
            Next
               
            Call xls_ds_draw_polyline(a1D_coordinates, layer_name, layer_colour, polyline_closed)
            polyline_closed = True
        End If
        
    Next
        Sheets("Product Definition").Activate
End Sub

