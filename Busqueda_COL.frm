VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Busqueda_COL 
   Caption         =   "Proyeccion_Demanda"
   ClientHeight    =   7110
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   10020
   OleObjectBlob   =   "Busqueda_COL.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "Busqueda_COL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False




Public Encontrado As Double

Private Sub Aceptar_Click()
 
Dim Crange As Range
Dim IRange As Range
Dim ORange As Range
Dim PRange, PRange2, PRange3, PRange4, wer5 As Range

Dim Flag As Integer
Dim Selecciones  As Integer
Dim CostValue_Max As Long


Dim Revision As String
Dim Palabra As String
Dim Shp As Shape



Dim nombreHoja As String

nombreHoja = "RESULTADO"
    
     If (BuscarHoja1(nombreHoja)) Then
     Application.DisplayAlerts = False
       Sheets(nombreHoja).Select
     ActiveWindow.SelectedSheets.Delete
      
    End If
    
nombreHoja = "Proyección"
    
     If (BuscarHoja1(nombreHoja)) Then
     Application.DisplayAlerts = False
       Sheets(nombreHoja).Select
     ActiveWindow.SelectedSheets.Delete
      
    End If

Worksheets.Add.Name = "RESULTADO"
Worksheets.Add.Name = "Proyección"



Worksheets("Hoja1").Activate

StartTime = Timer

NextCol = Cells(1, Columns.Count).End(xlToLeft).Column
LastRow = Cells(Rows.Count, 80).End(xlUp).Row
Cells(1, 50).Resize(LastRow, 148).ClearContents

NextCcol = 70
NextTCol = 60

'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXX
'XXXXXX    EL VALOR DEL 7 ES POR LA CANTIDAD DE OBJETOS
'XXXXXX    DEBEMOS HACER 6 SI EXCLUIMOS LA SELECCION DE PAIS
'XXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX

For j = 1 To 6   '  ESTE VALOR DE 6 DEBE SER 7 SI TENEMOS EL SELECTOR DE PAISES

    Select Case j
        Case 1
            MyControl = "lbSubcat" 'COLOR
            MyColumn = 5
        Case 2
            MyControl = "LMarca" ' HORMA
            MyColumn = 6
        Case 3
            MyControl = "LForma" 'TOBILLO
            MyColumn = 7
        Case 4
            MyControl = "LColor" 'PLANTA
            MyColumn = 10
        Case 5
        MyControl = "LPiedra" 'TACO
            MyColumn = 11
        Case 6
        MyControl = "LLinea" ' TALLA
            MyColumn = 12
        Case 7
        MyControl = "Pais"
        MyColumn = 1
     
     End Select

NextRow = 2

                For i = 0 To Me.Controls(MyControl).ListCount - 1
                    If Me.Controls(MyControl).Selected(i) = True Then
                        'Cells(NextRow, NextTCol).Value = _
                        'Me.Controls(MyControl).List(i)
                        NextRow = NextRow + 1
                    End If
                Next i


    If NextRow > 2 Then

    Selecciones = Selecciones + 1
    
    Else
    
    Selecciones = Selecciones + 0
    
    End If

Next j

'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX

If Me.Controls("OptionButton4").Value = True And Selecciones >= 1 Then

    'Hay que filtrar por listbox
    
                NextCol = Cells(1, Columns.Count).End(xlToLeft).Column
                LastRow = Cells(Rows.Count, 80).End(xlUp).Row
                Cells(1, 50).Resize(LastRow, 128).ClearContents
                
                NextCcol = 70
                NextTCol = 60
                
                For j = 1 To 6
                
                    Select Case j
                         Case 1
                            MyControl = "lbSubcat"
                            MyColumn = 5
                        Case 2
                            MyControl = "LMarca"
                            MyColumn = 6
                        Case 3
                            MyControl = "LForma"
                            MyColumn = 7
                        Case 4
                            MyControl = "LColor"
                            MyColumn = 10
                        Case 5
                        MyControl = "LPiedra"
                            MyColumn = 11
                        Case 6
                            MyControl = "LLinea"
                            MyColumn = 12
                        Case 7
                            MyControl = "Pais"
                            MyColumn = 1
                     End Select
                
                     NextRow = 2
                
                            For i = 0 To Me.Controls(MyControl).ListCount - 1
                                If Me.Controls(MyControl).Selected(i) = True Then
                                    Cells(NextRow, NextTCol).Value = _
                                    Me.Controls(MyControl).List(i)
                                    NextRow = NextRow + 1
                                End If
                            Next i
                
                
                            If NextRow > 2 Then
                            
                                MyFormula = "=NOT(ISNA(MATCH(RC" & MyColumn & ", R2C" & NextTCol & ":R" & NextRow - 1 & "C" & NextTCol & ",0)))"
                                'Cells(2, NextCcol).FormulaR1C1 = MyFormula
                                Cells(2, NextCcol).FormulaR1C1 = MyFormula
                                NextTCol = NextTCol + 1
                                NextCcol = NextCcol + 1
                            End If
                
                Next j
                
                Unload Me
                
                If NextCcol > 70 Then
                    
                                       
                    Set Crange = Range(Cells(1, 70), Cells(2, NextCcol - 1))
                    Set IRange = Range("A1").CurrentRegion
                    Set ORange = Cells(1, 80)
                    
                    
                    IRange.AdvancedFilter xlFilterCopy, Crange, ORange
                    
                    Cells(1, 70).Resize(1, 10).EntireColumn.Clear
                
                End If
                
                
                Set WSD = Worksheets("Hoja1")
                FinalRow10 = WSD.Cells(Rows.Count, 80).End(xlUp).Row
                FinalCol11 = WSD.Cells(1, Columns.Count).End(xlToLeft).Column
                
                canti = FinalRow10 - 1
                
                If canti > 0 Then
                
               
                Range("cb1").AutoFilter , Field:=4, Criteria1:="*" & Me.TextBox4.Text & "*"
                'Unload Me
        
                 
                
                End If
                
                
                FinalRow10 = WSD.Cells(Rows.Count, 80).End(xlUp).Row
                FinalCol11 = WSD.Cells(1, Columns.Count).End(xlToLeft).Column
                
                canti = FinalRow10 - 1
                
                If canti > 0 Then
                
                Continua = 0
                               
                Else
                Continua = 1
                
                         
                End If
                
                
                
                Select Case Continua
                Case 1:
                
                    Worksheets("RESULTADO").Activate
                   
                    MsgBox "La búsqueda realizada no encuentra registros, intente otra selección."
                
                Exit Sub
                Case Else
                
                Set PRange4 = WSD.Cells(1, 80).Resize(FinalRow10, FinalCol11)
                    PRange4.Select
                    
                    Selection.Copy
                    
                 
                    
                     Worksheets("RESULTADO").Activate
                    Range("A1").Select
                    ActiveSheet.Paste

                    GoTo Formato


                    
                    campi = Worksheets("Resultado").Cells(14, 2).Value
                    
                    
                    
                    Exit Sub
                    
                End Select

    
    
    
'xxxxxxxxxxxxxxxxxxx  FIN DE NOMBRE y ELECCION DE SUBCATEGORIA
End If


'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX

    
                If Me.Controls("OptionButton4").Value = True And Selecciones = 0 Then  ' NOMBRE DE PRODUCTO
          
                  Puntaje4 = 1    'MsgBox "Checked"
                               
                End If
                
                If Puntaje4 <> 0 Then
                 
                    Aviso = 1
                Else
                
                    Aviso = 0
                
                End If
                    
                Select Case Aviso
                Case 1:
                
                ' ACABA FILTRA EN SITIO Y HACE TODA LA RUTINA y TERMINA EL PROCEDURE
                Range("A1").AutoFilter , Field:=4, Criteria1:="*" & Me.TextBox4.Text & "*"
                Unload Me
            
                    Set WSD = Worksheets("Hoja1")
                    FinalRow10 = WSD.Cells(Rows.Count, 1).End(xlUp).Row
                    FinalCol11 = WSD.Cells(1, Columns.Count).End(xlToLeft).Column
                    
                    canti = FinalRow10 - 1
                    
                    If canti > 0 Then
                    Continua = 0
                    Else
                    Continua = 1
                    End If
            
                       Select Case Continua
                            Case 1:
                                ActiveSheet.ShowAllData
                                Worksheets("RESULTADO").Activate
                              
                                MsgBox "La búsqueda realizada no encuentra registros, intente otra selección."
                                 Exit Sub
                            
                            Case Else
                            
                            Set PRange4 = WSD.Cells(1, 1).Resize(FinalRow10, FinalCol11)
                            PRange4.Select
                            
                            Selection.Copy
                            
                        
                    Worksheets("RESULTADO").Activate
                    Range("A1").Select
                    ActiveSheet.Paste
                    
                    GoTo Formato
                                   
                    campi = Worksheets("Resultado").Cells(14, 2).Value
                            
                    
                    

                    Exit Sub
                            
                            End Select

                            Worksheets("RESULTADO").Activate
                        
                                        
                             
                                Range("A2").Select
                                Exit Sub
    
                    Case 0:
                    
                
                    End Select
    
    
    Palabra = Trim(Me.TextBox4.Value)
         
Flag = 777

NextCol = Cells(1, Columns.Count).End(xlToLeft).Column
LastRow = Cells(Rows.Count, 80).End(xlUp).Row
Cells(1, 50).Resize(LastRow, 128).ClearContents

NextCcol = 70
NextTCol = 60

For j = 1 To 6   ' TENER EN CUENTA QUE EL VALOR DE 6 ES PORQUE NO HAY SELECTOR DE PAIS
    Select Case j
        Case 1
            MyControl = "lbSubcat"
            MyColumn = 5
        Case 2
            MyControl = "LMarca"
            MyColumn = 6
        Case 3
            MyControl = "LForma"
            MyColumn = 7
        Case 4
            MyControl = "LColor"
            MyColumn = 10
        Case 5
        MyControl = "LPiedra"
            MyColumn = 11
        Case 6
        MyControl = "LLinea"
            MyColumn = 12
        
        Case 7
        MyControl = "Pais"
        MyColumn = 1
     End Select

NextRow = 2

For i = 0 To Me.Controls(MyControl).ListCount - 1
    If Me.Controls(MyControl).Selected(i) = True Then
        Cells(NextRow, NextTCol).Value = _
        Me.Controls(MyControl).List(i)
        NextRow = NextRow + 1
    End If
Next i


If NextRow > 2 Then

    MyFormula = "=NOT(ISNA(MATCH(RC" & MyColumn & ", R2C" & NextTCol & ":R" & NextRow - 1 & "C" & NextTCol & ",0)))"
    'Cells(2, NextCcol).FormulaR1C1 = MyFormula
    Cells(2, NextCcol).FormulaR1C1 = MyFormula
    NextTCol = NextTCol + 1
    NextCcol = NextCcol + 1
End If

Next j

Unload Me

If NextCcol > 70 Then

    Set Crange = Range(Cells(1, 70), Cells(2, NextCcol - 1))
    
    Set IRange = Range("A1").CurrentRegion
    
    
    
    Set ORange = Cells(1, 80)
    IRange.AdvancedFilter xlFilterCopy, Crange, ORange
    
    'Cells(1, 70).Resize(1, 10).EntireColumn.Clear

End If


Set WSD = Worksheets("Hoja1")
FinalRow = WSD.Cells(Rows.Count, 80).End(xlUp).Row
FinalCol = WSD.Cells(1, Columns.Count).End(xlToLeft).Column

canti = FinalRow - 1

If canti > 0 Then
Continua = 0
Else
Continua = 1
End If

Select Case Continua
Case 1:

    Worksheets("RESULTADO").Activate
   
    MsgBox "La búsqueda realizada no encuentra registros, intente otra selección."

Exit Sub
Case Else

End Select


FinalRow = WSD.Cells(Rows.Count, 80).End(xlUp).Row
FinalCol = WSD.Cells(1, Columns.Count).End(xlToLeft).Column


Set PRange = WSD.Cells(1, 80).Resize(FinalRow, FinalCol + 1)
PRange.Select

'TTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTT
'TTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTT
'TTTTTTTTTTT                                                          TTTTTTTTT
'TTTTTTTTTTT  DESARROLLO DE FILTRO DE CAMPAÑAS. COMBOBOX1             TTTTTTTTT
'TTTTTTTTTTT                                                          TTTTTTTTT
'TTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTT
'TTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTT

    
    FinalRow2 = WSD.Cells(Rows.Count, 80).End(xlUp).Row
     
    canti = FinalRow2 - (FinalRow + 7)
    
    If canti = 0 Then
    Aviso = 1
    Else
    Aviso = 0
    End If
    
    Select Case Aviso
    Case 1:
    Encontrado = Cells(2, 80).Value
    
     Worksheets("RESULTADO").Activate
     
     MsgBox "La búsqueda realizada no encuentra registros en el rango seleccionado pero intente desde " & Encontrado & " en adelante y encontrará lo que busca."

    
    Exit Sub
    Case 0:
        
     End Select
    

'TTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTT
'TTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTT
'TTTTTTTTTTT                                                          TTTTTTTTT
'TTTTTTTTTTT  DESARROLLO DE FILTRO DE CAMPAÑAS. COMBOBOX2             TTTTTTTTT
'TTTTTTTTTTT                                                          TTTTTTTTT
'TTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTT
'TTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTT


FinalRow2 = WSD.Cells(Rows.Count, 80).End(xlUp).Row
FinalCol2 = WSD.Cells(1, Columns.Count).End(xlToLeft).Column
Set PRange2 = WSD.Cells(FinalRow + 7, 80).Resize(FinalRow2, FinalCol2 + 1)
PRange2.Select

    'ActiveSheet.Range(Cells(1, 80), Cells(1, FinalCol)).Select
    'Selection.Copy
    'ActiveSheet.Cells(FinalRow2 + 3, 80).Select
    'ActiveSheet.Paste
    'ActiveSheet.Cells(FinalRow2 + 4, 81).Value = "<=" & ComboBox2.Value
    
    

    'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
    'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
    'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
    'XXXXXXXXXXXXXXXXXXXXXXXX  AQUI SE COLOCA LA SINTAXIS DE ESTIMACION   XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
    'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
    'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
    
    ActiveSheet.Range(Cells(FinalRow2 + 3, 80), Cells(FinalRow2 + 4, 128)).Name = "kari9" ' CRITERIOS DE FILTRADO EN COMBOBOX2
    'Range("kari9").Select
    
    ActiveSheet.Cells(FinalRow2 + 7, 80).Name = "kari8" 'UBICACION DEL RESULTADO DEL FILTRO AVANZADO
    'Range("kari8").Select
    
    PRange2. _
        AdvancedFilter Action:=xlFilterCopy, CriteriaRange:=Range("kari9") _
        , CopyToRange:=Range("kari8"), Unique:=True

'TTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTT
'TTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTT
'TTTTTTTTTTT                                                          TTTTTTTTT
'TTTTTTTTTTT  SELECCION DE xlCaptioContains en Autofiltro             TTTTTTTTT
'TTTTTTTTTTT                                                          TTTTTTTTT
'TTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTT
'TTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTT


FinalRow3 = WSD.Cells(Rows.Count, 80).End(xlUp).Row
FinalRow4 = FinalRow2 + 7

Canti2 = FinalRow3 - FinalRow4

'Maxim_val_find = ActiveWorkbo
'LR5 = ActiveSheet.Cells(Rows.Count, 1).End(xlUp).Row  'Identifica cantidad de filas
Set wer44 = ActiveSheet.Range(Cells(FinalRow4 + 1, 123), Cells(FinalRow3, 123))
wer44.Select



FinalRow3 = WSD.Cells(Rows.Count, 80).End(xlUp).Row
   
    canti = FinalRow3 - (FinalRow2 + 7)
    
    If canti = 0 And Canti2 <> 0 Then
    Aviso = 1
    Else
    Aviso = 0
    End If
    
    Select Case Aviso
    Case 1:
    Worksheets("Hoja1").Activate
    Encontrado = Worksheets("Hoja1").Cells(FinalRow + 8, 81).Value
        
    Worksheets("RESULTADO").Activate
    
    
    Exit Sub
        
    Case 0:
      
    
   End Select
    
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXX   COPIA RANGO DE RESULTADOS PARA REPORTE FINAL    XXXXXX
'XXXXXXXXXXXXXXXX                                                   XXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX


    
FinalRow3 = WSD.Cells(Rows.Count, 80).End(xlUp).Row
FinalCol2 = WSD.Cells(1, Columns.Count).End(xlToLeft).Column
Set PRange3 = WSD.Cells(FinalRow2 + 7, 80).Resize(FinalRow3, FinalCol2 - 79)
'PRange3.Select
    
    Selection.Copy

Range("EH1").Select
ActiveSheet.Paste

    

'TTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTT
'TTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTT
'TTTTTTTTTTT                                                          TTTTTTTTT
'TTTTTTTTTTT                   REVISA LOS CHECKBOX                    TTTTTTTTT
'TTTTTTTTTTT                                                          TTTTTTTTT
'TTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTT
'TTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTT

    
    'If ActiveSheet.Shapes("CheckBox1").Value = xlOn Then
     
    If Me.Controls("OptionButton1").Value = True Then
    
      Puntaje1 = 1    'MsgBox "Checked"
     Else
       Puntaje1 = 0    'MsgBox "Unchecked"
        
     End If

 If Me.Controls("OptionButton2").Value = True Then
      
      Puntaje2 = 1    'MsgBox "Checked"
     Else
       Puntaje2 = 0    'MsgBox "Unchecked"
        
     End If
     
     If Me.Controls("OptionButton3").Value = True Then  ' FORMA
      
      Puntaje3 = 1    'MsgBox "Checked"
     Else
       Puntaje3 = 0    'MsgBox "Unchecked"
        
     End If
       
       
  If Me.Controls("OptionButton4").Value = True Then  ' NOMBRE DE PRODUCTO
      
      Puntaje4 = 1    'MsgBox "Checked"
     
     Else
       
       Puntaje4 = 0    'MsgBox "Unchecked"
        
    End If
     
  
'_____________________________________________________________________________________

'VERIFICA QUE SOLO 1 CHECK BOX ESTE ACTIVADO
 
    


    If Puntaje1 + Puntaje2 + Puntaje3 + Puntaje4 > 1 Then
     Aviso = 1
        Else
        Aviso = 0
        End If
        
    Select Case Aviso
    Case 1:
    
    
    MsgBox "Si va a buscar con palabras active una sola casilla. Ingrese nuevamente su selección"
    
    Exit Sub
    
    Case 0:
    
    End Select
    
    
    Palabra = Me.TextBox4.Value
    
    'MsgBox Palabra

  Select Case Puntaje1  ' COLOR
    Case 1:
    Range("E1").AutoFilter , Field:=5, Criteria1:="*" & Me.TextBox4.Text & "*"
    Case 0:
  End Select
  
  
  Select Case Puntaje2    'PIEDRA
    Case 1:
    Range("F1").AutoFilter , Field:=6, Criteria1:="*" & Me.TextBox4.Text & "*"
    Case 0:
  End Select

Select Case Puntaje3   ' ESTILO
    Case 1:
    Range("I1").AutoFilter , Field:=9, Criteria1:="*" & Me.TextBox4.Text & "*"
    Case 0:
  End Select
  
  Select Case Puntaje4  'NOMBRE
    Case 1:
   
    
    Range("D1").AutoFilter , Field:=4, Criteria1:="*" & Me.TextBox4.Text & "*"
    
    Case 0:
  End Select


FinalRow10 = WSD.Cells(Rows.Count, 91).End(xlUp).Row
FinalCol11 = WSD.Cells(1, Columns.Count).End(xlToLeft).Column

WSD.Cells(FinalRow10, FinalCol11).Select

Set PRange4 = WSD.Range(Cells(1, 80), Cells(FinalRow10, FinalCol11))
PRange4.Select

If FinalRow10 <> 1 Then
Mayi = 0
Else
Mayi = 1

End If

Select Case Mayi

Case 1:
    
    Worksheets("RESULTADO").Activate
   
    MsgBox " No se encuentran resultados con la selección ingresada. Intente nuevamente."

     
    Exit Sub

Case Else




End Select

Selection.Copy

'Worksheets.Add.Name = "RESULTADO"

  Worksheets("RESULTADO").Activate
                    Range("A1").Select
                    ActiveSheet.Paste
                    
Formato:

            lcolumn = Cells(1, Columns.Count).End(xlToLeft).Column
            lrow = Cells(Rows.Count, "A").End(xlUp).Row

            Cells.Select
            
               With Selection.Interior
                   .Pattern = xlNone
                   .TintAndShade = 0
                   .PatternTintAndShade = 0
               End With
   

canti = lrow - 1

   

    Range("A1").Select
    
    Call Estimacion
    

    EndTime = Timer
    
   kind1 = EndTime - StartTime
   kind2 = WorksheetFunction.Round(kind1 / 60, 2)
   
    MsgBox "¡Proceso terminado en " & kind1 & " segundos! ", vbInformation, xMensaje
   Debug.Print "Execution time in seconds: ", EndTime - StartTime

End Sub


Private Sub Cancel_Button2_Click()

Unload Me

End Sub

Private Sub CheckBox1_Click()

End Sub

Private Sub Clear_Click()

For i = 0 To lbSubcat.ListCount - 1
    Me.lbSubcat.Selected(i) = False
Next i

End Sub
Private Sub ComboBox1_Change()

Cinicio = ComboBox1.Value

If ComboBox1.Value > ComboBox2.Value Then

'MsgBox " Verifique que la campaña final sea mayor a la campaña inicial"

Else
'MsgBox " Verifique que la campaña final sea mayor a la campaña inicial"

End If


End Sub


Private Sub ComboBox2_Change()

CFinal = ComboBox2.Value

If ComboBox2.Value >= ComboBox1.Value Then

'MsgBox "La selección va desde " & ComboBox1.Value & " a " & ComboBox2.Value

Else
'MsgBox " Verifique que la campaña final sea mayor a la campaña inicial"

End If

End Sub



Private Sub CommandButton10_Click()
For i = 0 To LLinea.ListCount - 1
Me.LLinea.Selected(i) = True

Next i
End Sub

Private Sub CommandButton11_Click()
For i = 0 To LLinea.ListCount - 1
    Me.LLinea.Selected(i) = False
Next i
End Sub

Private Sub CommandButton12_Click()
For i = 0 To LPiedra.ListCount - 1
Me.LPiedra.Selected(i) = True

Next i
End Sub

Private Sub CommandButton13_Click()
For i = 0 To LPiedra.ListCount - 1
    Me.LPiedra.Selected(i) = False
Next i
End Sub

Private Sub CommandButton14_Click()
For i = 0 To Pais.ListCount - 1
Me.Pais.Selected(i) = True

Next i



End Sub

Private Sub CommandButton15_Click()
For i = 0 To Pais.ListCount - 1
    Me.Pais.Selected(i) = False
Next i

End Sub

Private Sub CommandButton3_Click()
For i = 0 To lbSubcat.ListCount - 1
Me.lbSubcat.Selected(i) = True

Next i

End Sub



Private Sub Label1_Click()

End Sub

Private Sub CommandButton4_Click()
For i = 0 To LMarca.ListCount - 1
Me.LMarca.Selected(i) = True

Next i
End Sub

Private Sub CommandButton5_Click()
For i = 0 To LMarca.ListCount - 1
    Me.LMarca.Selected(i) = False
Next i
End Sub

Private Sub CommandButton6_Click()
For i = 0 To LForma.ListCount - 1
Me.LForma.Selected(i) = True

Next i
End Sub

Private Sub CommandButton7_Click()
For i = 0 To LForma.ListCount - 1
    Me.LForma.Selected(i) = False
Next i
End Sub

Private Sub CommandButton8_Click()
For i = 0 To LColor.ListCount - 1
Me.LColor.Selected(i) = True
Next i

End Sub

Private Sub CommandButton9_Click()
For i = 0 To LColor.ListCount - 1
    Me.LColor.Selected(i) = False
Next i
End Sub

Private Sub Label7_Click()

End Sub

Private Sub lbSubcat_Click()

End Sub

Private Sub LColor_Click()

For i = 0 To LColor.ListCount - 1
Me.LColor.Selected(i) = True
Next i

End Sub

Private Sub OptionButton2_Click()

End Sub

Private Sub OptionButton3_Click()
MsgBox "Seleccione tambien una sub categoría para realizar la búsqueda."
End Sub

Private Sub OptionButton4_Click()

End Sub

Private Sub TextBox4_Change()

End Sub

Private Sub TextBox5_Change()

End Sub

Private Sub UserForm_Initialize()
Dim IRange As Range
Dim ORange As Range



'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX

Worksheets("Hoja1").Activate

Worksheets("Hoja1").Range(Cells(1, 14), Cells(300, 300)).ClearContents




'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX

FinalRow = Cells(Rows.Count, 1).End(xlUp).Row
NextCol = Cells(1, Columns.Count).End(xlToLeft).Column + 2

Set IRange = Range("A1").Resize(FinalRow, NextCol - 2)
IRange.Select


Cells(1, 5).Copy Destination:=Cells(1, NextCol)  'COLOR

Set ORange = Cells(1, NextCol)

IRange.AdvancedFilter Action:=xlFilterCopy, CopyToRange:=ORange, Unique:=True

LastRow = Cells(Rows.Count, NextCol).End(xlUp).Row

Cells(1, NextCol).Resize(LastRow, 1).Sort Key1:=Cells(1, NextCol), _
Order1:=xlAscending, Header:=xlYes


With Me.lbSubcat
.RowSource = ""
.List = Cells(2, NextCol).Resize(LastRow - 1, 1).Value
End With

Cells(1, NextCol).Resize(LastRow, 1).Clear


'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX

FinalRow = Cells(Rows.Count, 1).End(xlUp).Row
NextCol = Cells(1, Columns.Count).End(xlToLeft).Column + 2

Cells(1, 6).Copy Destination:=Cells(1, NextCol) 'HORMA

Set ORange = Cells(1, NextCol)
Set IRange = Range("A1").Resize(FinalRow, NextCol - 2)

IRange.AdvancedFilter Action:=xlFilterCopy, CopyToRange:=ORange, Unique:=True

LastRow = Cells(Rows.Count, NextCol).End(xlUp).Row

'Cells(1, NextCol).Resize(LastRow, 1).Sort Key1:=Cells(1, NextCol), _
'Order1:=xlAscending, Key2:=Cells(1, NextCol + 1), Order2:=xlAscending, Header:=xlYes

Cells(1, NextCol).Resize(LastRow, 1).Sort Key1:=Cells(1, NextCol), _
Order1:=xlAscending, Header:=xlYes



With Me.LMarca 'HORMA
.RowSource = ""
.List = Cells(2, NextCol).Resize(LastRow - 1, 1).Value
End With

Cells(1, NextCol).Resize(LastRow, 1).Clear


'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX

FinalRow = Cells(Rows.Count, 1).End(xlUp).Row
NextCol = Cells(1, Columns.Count).End(xlToLeft).Column + 2

Cells(1, 7).Copy Destination:=Cells(1, NextCol)  'TOBILLO

Set ORange = Cells(1, NextCol)

IRange.AdvancedFilter Action:=xlFilterCopy, CopyToRange:=ORange, Unique:=True

LastRow = Cells(Rows.Count, NextCol).End(xlUp).Row

'Cells(1, NextCol).Resize(LastRow, 1).Sort Key1:=Cells(1, NextCol), _
'Order1:=xlAscending, Key2:=Cells(1, NextCol + 1), Order2:=xlAscending, Header:=xlYes

Cells(1, NextCol).Resize(LastRow, 1).Sort Key1:=Cells(1, NextCol), _
Order1:=xlAscending, Header:=xlYes

With Me.LForma
.RowSource = ""
.List = Cells(2, NextCol).Resize(LastRow - 1, 1).Value
End With

Cells(1, NextCol).Resize(LastRow, 1).Clear


'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX

FinalRow = Cells(Rows.Count, 1).End(xlUp).Row
NextCol = Cells(1, Columns.Count).End(xlToLeft).Column + 2

Cells(1, 10).Copy Destination:=Cells(1, NextCol) 'PLANTA

Set ORange = Cells(1, NextCol)

IRange.AdvancedFilter Action:=xlFilterCopy, CopyToRange:=ORange, Unique:=True

LastRow = Cells(Rows.Count, NextCol).End(xlUp).Row

'Cells(1, NextCol).Resize(LastRow, 1).Sort Key1:=Cells(1, NextCol), _
'Order1:=xlAscending, Key2:=Cells(1, NextCol + 1), Order2:=xlAscending, Header:=xlYes

Cells(1, NextCol).Resize(LastRow, 1).Sort Key1:=Cells(1, NextCol), _
Order1:=xlAscending, Header:=xlYes

With Me.LColor ' PLANTA
.RowSource = ""
.List = Cells(2, NextCol).Resize(LastRow - 1, 1).Value
End With

Cells(1, NextCol).Resize(LastRow, 1).Clear


'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX


FinalRow = Cells(Rows.Count, 1).End(xlUp).Row
NextCol = Cells(1, Columns.Count).End(xlToLeft).Column + 2

Set IRange = Range("A1").Resize(FinalRow, NextCol - 2)

Cells(1, 12).Copy Destination:=Cells(1, NextCol) ' TACO

Set ORange = Cells(1, NextCol)

IRange.AdvancedFilter Action:=xlFilterCopy, CopyToRange:=ORange, Unique:=True

LastRow = Cells(Rows.Count, NextCol).End(xlUp).Row

Cells(1, NextCol).Resize(LastRow, 1).Sort Key1:=Cells(1, NextCol), _
Order1:=xlAscending, Header:=xlYes


With Me.LLinea 'TACO
.RowSource = ""
.List = Cells(2, NextCol).Resize(LastRow - 1, 1).Value
End With

Cells(1, NextCol).Resize(LastRow, 1).Clear


'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX


FinalRow = Cells(Rows.Count, 1).End(xlUp).Row
NextCol = Cells(1, Columns.Count).End(xlToLeft).Column + 2

Set IRange = Range("A1").Resize(FinalRow, NextCol - 2)

Cells(1, 11).Copy Destination:=Cells(1, NextCol) '

Set ORange = Cells(1, NextCol)

IRange.AdvancedFilter Action:=xlFilterCopy, CopyToRange:=ORange, Unique:=True

LastRow = Cells(Rows.Count, NextCol).End(xlUp).Row

Cells(1, NextCol).Resize(LastRow, 1).Sort Key1:=Cells(1, NextCol), _
Order1:=xlAscending, Header:=xlYes


With Me.LPiedra ' NUMERO TACO
.RowSource = ""
.List = Cells(2, NextCol).Resize(LastRow - 1, 1).Value
End With

Cells(1, NextCol).Resize(LastRow, 1).Clear

'Worksheets("Hoja3").Activate
'Worksheets("Hoja3").Select


End Sub
Function BuscarHoja1(nombreHoja As String) As Boolean
 
    For i = 1 To Worksheets.Count
        If Worksheets(i).Name = nombreHoja Then
            BuscarHoja1 = True
            Exit Function
        End If
    Next
     
    BuscarHoja1 = False
 
End Function

