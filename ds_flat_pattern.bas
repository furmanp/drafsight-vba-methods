Attribute VB_Name = "ds_flat_pattern"
Option Explicit
Option Base 1

Sub ds_fp_draw_main()
    'Subroutine imports Data from Product Definition sheet, and Cross section coordinates from Geometry Sheet. Creates flat pattern based on provided data
    Dim i As Integer, j As Integer, k As Integer, s_txt As String, a_temp As Variant
    Dim curve_ID As String
    Dim layer_name As String
    Dim layer_colour As Integer
    Dim offset_val As Double
    Dim section_id As Integer
    Dim a_data As Variant
    Dim a_coordinates As Variant
    Dim updated_coordinates As Variant
    Dim a1D_coordinates() As Double
    Dim curve_start As Double
    Dim curve_end As Variant
    Dim polyline_closed As Boolean
    Dim use_input As Integer
    polyline_closed = False
    
    Sheets("Product Definition").Activate
    
    a_data = a_make(".pd", 1, 1, 0, 0)
    a_data = a_row0(a_data)

    For i = 1 To UBound(a_data)
        curve_ID = a_look("Alignment", a_data, i) * 10000 + a_look("Geometry", a_data, i) * 1000 + a_look("Section", a_data, i) * 10 + a_look("Segment", a_data, i)
        layer_name = a_look("Name", a_data, i)
        layer_colour = a_look("Colour", a_data, i)
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
            If layer_name = "Mould" Then
                updated_coordinates = new_polyline(a_coordinates, curve_start, curve_end)
            
                If a_look("Alignment", a_data, i) = 1 Then
                    updated_coordinates = straigthen_polyline_midpoint(updated_coordinates, "V")
                ElseIf a_look("Alignment", a_data, i) = 2 Then
                    updated_coordinates = straigthen_polyline_midpoint(updated_coordinates, "U")
                End If
                                
                'Redimensioning 2D array into 1D Array and parsing them from Variant to Double
                ReDim a1D_coordinates(1 To UBound(updated_coordinates) * 2)
                For j = 1 To UBound(updated_coordinates)
                    If a_look("Alignment", a_data, i) = 1 Then
                        a1D_coordinates((j - 1) * 2 + 1) = CDbl(updated_coordinates(j, 1))
                        a1D_coordinates((j - 1) * 2 + 2) = CDbl(a_coordinates(1, 3))
                    ElseIf a_look("Alignment", a_data, i) = 2 Then
                        a1D_coordinates((j - 1) * 2 + 1) = CDbl(a_coordinates(1, 3))
                        a1D_coordinates((j - 1) * 2 + 2) = CDbl(updated_coordinates(j, 2))
                    End If
                Next
                Call xls_ds_draw_polyline(a1D_coordinates, layer_name, layer_colour, polyline_closed)

            End If
        End If
    Next
    Sheets("Product Definition").Activate
End Sub

