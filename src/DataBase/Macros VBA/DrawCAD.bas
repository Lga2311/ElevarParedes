Attribute VB_Name = "DrawCAD"
Sub DrawGABARITO()

'Variáveis do Cad para selecionar o programa
Dim myapp As Object
Dim mydwg As AcadDocument

'Verificação de Erro pra alguma coisa (pesquisar)
Set myapp = GetObject(, "Autocad.application")
ERRORHANDLER:
    If Err.Description <> "" Then
        Err.Clear
        Set myapp = CreateObject("Autocad.application")
    End If

'Selecionar cad aberto
myapp.Visible = True

'Variáveis pra desenhar os círculo
Dim POLYLINE As AcadPolyline  'CAD

    'Variáveis: Pontos e Ciclos do FOR
    Dim pts(0 To 11) As Double
    Dim CicloLim As Single
    
    'Localizando o limite do FOR
    Sheets("Banco de Dados").Select
    CicloLim = Application.WorksheetFunction.Max(Range("BC:BC").Value)
    
    'Selecionando primeiro X
    Range("Be6").Select
    
    'Desenhar GABARITO
    For Ciclo = 0 To CicloLim - 1

        'Pesquisar pontos no Excel
        For i = 0 To 11
                    
            'Linha
            lin = Application.WorksheetFunction.RoundDown(i / 3, 0)
            
            'Coluna
            col = i Mod 3
            
            'Valor da Célula
            'Col = 2 >>> ponto Z que é igual a 0
            If col = 2 Then pts(i) = 0 Else pts(i) = ActiveCell.Offset(lin, col)
            
        Next
        
        'Desenhando no cad e atualizando visualização
        Set POLYLINE = AutoCAD.Application.ActiveDocument.ModelSpace.AddPolyline(pts)
        POLYLINE.Layer = Application.WorksheetFunction.XLookup("DrawGABARITO", Worksheets("Layers").Range("C:C"), Worksheets("Layers").Range("A:A"))
        AutoCAD.Application.Update
        
        ActiveCell.Offset(5, 0).Select
        
    Next

End Sub

Sub DrawLAJE()

'Variáveis do Cad para selecionar o programa
Dim myapp As Object
Dim mydwg As AcadDocument

'Verificação de Erro pra alguma coisa (pesquisar)
Set myapp = GetObject(, "Autocad.application")
ERRORHANDLER:
    If Err.Description <> "" Then
        Err.Clear
        Set myapp = CreateObject("Autocad.application")
    End If

'Selecionar cad aberto
myapp.Visible = True

'Variáveis pra desenhar os círculo
Dim POLYLINE As AcadPolyline  'CAD

    'Variáveis: Pontos e Ciclos do FOR
    Dim pts(0 To 11) As Double
    Dim CicloLim As Single
    
    'Localizando o limite do FOR
    Sheets("Banco de Dados").Select
    CicloLim = Application.WorksheetFunction.Max(Range("Bi:Bi").Value)
    
    'Selecionando primeiro X
    Range("Bk6").Select
    
    'Desenhar GABARITO
    For Ciclo = 0 To CicloLim - 1
        
        'Pesquisar pontos no Excel
        For i = 0 To 11
                    
            'Linha
            lin = Application.WorksheetFunction.RoundDown(i / 3, 0)
            
            'Coluna
            col = i Mod 3
            
            'Valor da Célula
            Select Case True
            
                'Col = 2 >>> ponto Z que é igual a 0
                Case col = 2
                    pts(i) = 0
                
                Case Not IsError(ActiveCell.Offset(lin, col)) And col <> 2
                    pts(i) = ActiveCell.Offset(lin, col)
                
                'Tratando paredes com Viga acima
                Case IsError(ActiveCell.Offset(lin, col))
                    pts(i) = 10 ^ 12
                    
            End Select
            
        Next
        
        'Desenhando no cad e atualizando visualização
        If pts(0) <> 10 ^ 12 And pts(1) <> 10 ^ 12 And pts(2) <> 10 ^ 12 Then
            Set POLYLINE = AutoCAD.Application.ActiveDocument.ModelSpace.AddPolyline(pts)
            POLYLINE.Layer = Application.WorksheetFunction.XLookup("DrawLAJE", Worksheets("Layers").Range("C:C"), Worksheets("Layers").Range("A:A"))
            POLYLINE.Closed = True
            AutoCAD.Application.Update
            
            'Proxima parede
            ActiveCell.Offset(6, 0).Select
        Else
            'Proxima parede
            ActiveCell.Offset(6, 0).Select
        End If
        
    Next

End Sub

Sub DrawESQ()

'Variáveis do Cad para selecionar o programa
Dim myapp As Object
Dim mydwg As AcadDocument

'Verificação de Erro pra alguma coisa (pesquisar)
Set myapp = GetObject(, "Autocad.application")
ERRORHANDLER:
    If Err.Description <> "" Then
        Err.Clear
        Set myapp = CreateObject("Autocad.application")
    End If

'Selecionar cad aberto
myapp.Visible = True

'Variáveis pra desenhar os círculo
Dim POLYLINE As AcadPolyline  'CAD

    'Variáveis: Pontos e Ciclos do FOR
    Dim pts(0 To 11) As Double
    Dim CicloLim As Single
    
    'Localizando o limite do FOR
    Sheets("Banco de Dados").Select
    CicloLim = Application.WorksheetFunction.Max(Range("Bo:Bo").Value)
    
    'Selecionando primeiro X
    Range("Bq6").Select
    
    'Desenhar GABARITO
    For Ciclo = 0 To CicloLim - 1

        'Pesquisar pontos no Excel
        For i = 0 To 11
                    
            'Linha
            lin = Application.WorksheetFunction.RoundDown(i / 3, 0)
            
            'Coluna
            col = i Mod 3
            
            'Valor da Célula
            'Col = 2 >>> ponto Z que é igual a 0
            If col = 2 Then pts(i) = 0 Else pts(i) = ActiveCell.Offset(lin, col)
            
        Next
        
        'Desenhando no cad e atualizando visualização
        Set POLYLINE = AutoCAD.Application.ActiveDocument.ModelSpace.AddPolyline(pts)
        POLYLINE.Layer = Application.WorksheetFunction.XLookup("DrawESQ", Worksheets("Layers").Range("C:C"), Worksheets("Layers").Range("A:A"))
        POLYLINE.Closed = True
        AutoCAD.Application.Update
        
        ActiveCell.Offset(6, 0).Select
        
    Next

End Sub

Sub DrawGROUTE()

'Variáveis do Cad para selecionar o programa
Dim myapp As Object
Dim mydwg As AcadDocument

'Verificação de Erro pra alguma coisa (pesquisar)
Set myapp = GetObject(, "Autocad.application")
ERRORHANDLER:
    If Err.Description <> "" Then
        Err.Clear
        Set myapp = CreateObject("Autocad.application")
    End If

'Selecionar cad aberto
myapp.Visible = True

'Variáveis pra desenhar os círculo
Dim LINE As AcadLine  'CAD

    'Variáveis: Pontos e Ciclos do FOR
    Dim Inicio(0 To 2) As Double
    Dim Fim(0 To 2) As Double
    Dim CicloLim As Single
    
    'Localizando o limite do FOR
    Sheets("Banco de Dados").Select
    CicloLim = Range("BX1").Value
    
    'Selecionando primeiro X
    Range("CA6").Select
    
    'Desenhar GABARITO
    For Ciclo = 0 To CicloLim - 1

        'Pesquisar pontos no Excel
        Inicio(0) = ActiveCell.Offset(0, 0)
        Inicio(1) = ActiveCell.Offset(0, 1)
        Inicio(2) = 0
        Fim(0) = ActiveCell.Offset(1, 0)
        Fim(1) = ActiveCell.Offset(1, 1)
        Fim(2) = 0
        
        'Desenhando no cad e atualizando visualização
        Set LINE = AutoCAD.Application.ActiveDocument.ModelSpace.AddLine(Inicio, Fim)
        LINE.Layer = Application.WorksheetFunction.XLookup("DrawGROUTE", Worksheets("Layers").Range("C:C"), Worksheets("Layers").Range("A:A"))

        AutoCAD.Application.Update
        
        ActiveCell.Offset(3, 0).Select
        
    Next

End Sub

Sub DrawCANALETA_ALV()

'Variáveis do Cad para selecionar o programa
Dim myapp As Object
Dim mydwg As AcadDocument

'Verificação de Erro pra alguma coisa (pesquisar)
Set myapp = GetObject(, "Autocad.application")
ERRORHANDLER:
    If Err.Description <> "" Then
        Err.Clear
        Set myapp = CreateObject("Autocad.application")
    End If

'Selecionar cad aberto
myapp.Visible = True

'Variáveis pra desenhar os círculo
Dim POLYLINE As AcadPolyline  'CAD

    'Variáveis: Pontos e Ciclos do FOR
    Dim pts(0 To 11) As Double
    Dim CicloLim As Single
    
    'Localizando o limite do FOR
    Sheets("Banco de Dados").Select
    CicloLim = Application.WorksheetFunction.Max(Range("CH:CH").Value)
    
    'Selecionando primeiro X
    Range("CK6").Select
    
    'Desenhar Canaletas na Alvenaria
    For Ciclo = 0 To CicloLim - 1
        
        'Pesquisar pontos no Excel
        
        For i = 0 To 11
                    
            'Linha
            lin = Application.WorksheetFunction.RoundDown(i / 3, 0)
            
            'Coluna
            col = i Mod 3
            
            'Valor da Célula
            Select Case True
                'Col = 2 >>> ponto Z que é igual a 0
                Case col = 2
                    pts(i) = 0
                
                Case Not IsError(ActiveCell.Offset(lin, col)) And col <> 2
                    pts(i) = ActiveCell.Offset(lin, col)
                
                'Tratando paredes com Viga acima
                Case IsError(ActiveCell.Offset(lin, col))
                    pts(i) = 10 ^ 12
                    
            End Select
            
        Next
        
        'Desenhando no cad e atualizando visualização
        'Tratando paredes com Viga acima
        If pts(0) <> 10 ^ 12 And pts(1) <> 10 ^ 12 And pts(2) <> 10 ^ 12 Then
            Set POLYLINE = AutoCAD.Application.ActiveDocument.ModelSpace.AddPolyline(pts)
            POLYLINE.Layer = Application.WorksheetFunction.XLookup("DrawCANALETA_ALV", Worksheets("Layers").Range("C:C"), Worksheets("Layers").Range("A:A"))
            POLYLINE.Closed = False
            AutoCAD.Application.Update
            
            'Proxima parede
            ActiveCell.Offset(5, 0).Select
        Else
            'Proxima parede
            ActiveCell.Offset(5, 0).Select
        End If
        
    Next
' - - - - -- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -- - - - - - - - -- - - -
    'Selecionando primeiro X
    Range("CM6").Select
    
    'Desenhar Canaletas no DETALHAMENTO
    For Ciclo = 0 To CicloLim - 1
        
        'Pesquisar pontos no Excel
        
        For i = 0 To 11
                    
            'Linha
            lin = Application.WorksheetFunction.RoundDown(i / 3, 0)
            
            'Coluna
            col = i Mod 3
            
            'Valor da Célula
            Select Case True
            
                'Col = 2 >>> ponto Z que é igual a 0
                Case col = 2
                    pts(i) = 0
                
                Case Not IsError(ActiveCell.Offset(lin, col)) And col <> 2
                    pts(i) = ActiveCell.Offset(lin, col)
                
                'Tratando paredes com Viga acima
                Case IsError(ActiveCell.Offset(lin, col))
                    pts(i) = 10 ^ 12
                    
            End Select

        Next
        
        'Desenhando no cad e atualizando visualização
        'Tratando paredes com Viga acima
        If pts(0) <> 10 ^ 12 And pts(1) <> 10 ^ 12 And pts(2) <> 10 ^ 12 Then
            Set POLYLINE = AutoCAD.Application.ActiveDocument.ModelSpace.AddPolyline(pts)
            POLYLINE.Layer = Application.WorksheetFunction.XLookup("DrawCANALETA_ALV", Worksheets("Layers").Range("C:C"), Worksheets("Layers").Range("A:A"))
            POLYLINE.Closed = False
            AutoCAD.Application.Update
            
            'Proxima parede
            ActiveCell.Offset(5, 0).Select
        Else
            'Proxima parede
            ActiveCell.Offset(5, 0).Select
        End If
        
    Next
End Sub

Sub DrawCANALETA_VIGAS()

'Variáveis do Cad para selecionar o programa
Dim myapp As Object
Dim mydwg As AcadDocument

'Verificação de Erro pra alguma coisa (pesquisar)
Set myapp = GetObject(, "Autocad.application")
ERRORHANDLER:
    If Err.Description <> "" Then
        Err.Clear
        Set myapp = CreateObject("Autocad.application")
    End If

'Selecionar cad aberto
myapp.Visible = True

'Variáveis pra desenhar os círculo
Dim POLYLINE As AcadPolyline  'CAD

    'Variáveis: Pontos e Ciclos do FOR
    Dim pts(0 To 11) As Double
    Dim CicloLim As Single
    
    'Localizando o limite do FOR
    Sheets("Banco de Dados").Select
    CicloLim = 2 * Range("CQ1").Value
    
    'Selecionando primeiro X
    Range("CU6").Select
    
    'Desenhar Ferro NA VIGA
    For Ciclo = 0 To CicloLim - 1
        
        'Pesquisar pontos no Excel
        
        For i = 0 To 11
                    
            'Linha
            lin = Application.WorksheetFunction.RoundDown(i / 3, 0)
            
            'Coluna
            col = i Mod 3
            
            'Valor da Célula
            Select Case True
                'Col = 2 >>> ponto Z que é igual a 0
                Case col = 2
                    pts(i) = 0
                
                Case Not IsError(ActiveCell.Offset(lin, col)) And col <> 2
                    pts(i) = ActiveCell.Offset(lin, col)
                
                'Tratando paredes com Viga acima
                Case IsError(ActiveCell.Offset(lin, col))
                    pts(i) = 10 ^ 12
                    
            End Select
            
        Next
        
        'Desenhando no cad e atualizando visualização
        'Tratando paredes com Viga acima
        If pts(0) <> 10 ^ 12 And pts(1) <> 10 ^ 12 And pts(2) <> 10 ^ 12 Then
            Set POLYLINE = AutoCAD.Application.ActiveDocument.ModelSpace.AddPolyline(pts)
            POLYLINE.Layer = Application.WorksheetFunction.XLookup("DrawCANALETA_VIGAS", Worksheets("Layers").Range("C:C"), Worksheets("Layers").Range("A:A"))
            POLYLINE.Closed = False
            AutoCAD.Application.Update
            
            'Proxima parede
            ActiveCell.Offset(5, 0).Select
        Else
            'Proxima parede
            ActiveCell.Offset(5, 0).Select
        End If
        
    Next
' - - - - -- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -- - - - - - - - -- - - -

    'Selecionando primeiro X
    Range("CW6").Select
    
    'Desenhar Ferro DETALHAMENTO
    For Ciclo = 0 To CicloLim - 1
        
        'Pesquisar pontos no Excel
        
        For i = 0 To 11
                    
            'Linha
            lin = Application.WorksheetFunction.RoundDown(i / 3, 0)
            
            'Coluna
            col = i Mod 3
            
            'Valor da Célula
            Select Case True
                'Col = 2 >>> ponto Z que é igual a 0
                Case col = 2
                    pts(i) = 0
                
                Case Not IsError(ActiveCell.Offset(lin, col)) And col <> 2
                    pts(i) = ActiveCell.Offset(lin, col)
                
                'Tratando paredes com Viga acima
                Case IsError(ActiveCell.Offset(lin, col))
                    pts(i) = 10 ^ 12
                    
            End Select
            
        Next
        
        'Desenhando no cad e atualizando visualização
        'Tratando paredes com Viga acima
        If pts(0) <> 10 ^ 12 And pts(1) <> 10 ^ 12 And pts(2) <> 10 ^ 12 Then
            Set POLYLINE = AutoCAD.Application.ActiveDocument.ModelSpace.AddPolyline(pts)
            POLYLINE.Layer = Application.WorksheetFunction.XLookup("DrawCANALETA_VIGAS", Worksheets("Layers").Range("C:C"), Worksheets("Layers").Range("A:A"))
            POLYLINE.Closed = False
            AutoCAD.Application.Update
            
            'Proxima parede
            ActiveCell.Offset(5, 0).Select
        Else
            'Proxima parede
            ActiveCell.Offset(5, 0).Select
        End If
        
    Next

End Sub

Sub DrawCANALETA_ESQ()

'Variáveis do Cad para selecionar o programa
Dim myapp As Object
Dim mydwg As AcadDocument

'Verificação de Erro pra alguma coisa (pesquisar)
Set myapp = GetObject(, "Autocad.application")
ERRORHANDLER:
    If Err.Description <> "" Then
        Err.Clear
        Set myapp = CreateObject("Autocad.application")
    End If

'Selecionar cad aberto
myapp.Visible = True

'Variáveis pra desenhar os círculo
Dim POLYLINE As AcadPolyline  'CAD

    'Variáveis: Pontos e Ciclos do FOR
    Dim pts(0 To 5) As Double
    Dim CicloLim As Single
    
    'Localizando o limite do FOR
    Sheets("Banco de Dados").Select
    CicloLim = 2 * Application.WorksheetFunction.Max(Range("DA:DA").Value)
    
    'Selecionando primeiro X
    Range("DD6").Select
    
    'Desenhar Ferro NA ALVENARIA
    For Ciclo = 0 To CicloLim - 1
        
        'Pesquisar pontos no Excel
        For i = 0 To 5
                    
            'Linha
            lin = Application.WorksheetFunction.RoundDown(i / 3, 0)
            
            'Coluna
            col = i Mod 3
            
            'Valor da Célula
            Select Case True
                'Col = 2 >>> ponto Z que é igual a 0
                Case col = 2
                    pts(i) = 0
                
                Case Not IsError(ActiveCell.Offset(lin, col)) And col <> 2
                    pts(i) = ActiveCell.Offset(lin, col)
                
                'Tratando paredes com Viga acima
                Case IsError(ActiveCell.Offset(lin, col))
                    pts(i) = 10 ^ 12
                    
            End Select
            
        Next
        
        'Desenhando no cad e atualizando visualização
        'Tratando paredes com Viga acima
        If Not IsError(ActiveCell.Value) Then
            Set POLYLINE = AutoCAD.Application.ActiveDocument.ModelSpace.AddPolyline(pts)
            POLYLINE.Layer = Application.WorksheetFunction.XLookup("DrawCANALETA_ESQ", Worksheets("Layers").Range("C:C"), Worksheets("Layers").Range("A:A"))
            POLYLINE.Closed = False
            AutoCAD.Application.Update
            
            'Proxima parede
            ActiveCell.Offset(3, 0).Select
        Else
            'Proxima parede
            ActiveCell.Offset(3, 0).Select
        End If
        
    Next
' - - - - -- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -- - - - - - - - -- - - -
    'Selecionando primeiro X
    Range("DF6").Select
    
    'Desenhar Ferro NA ALVENARIA
    For Ciclo = 0 To CicloLim - 1
        
        'Pesquisar pontos no Excel
        For i = 0 To 5
                    
            'Linha
            lin = Application.WorksheetFunction.RoundDown(i / 3, 0)
            
            'Coluna
            col = i Mod 3
            
            'Valor da Célula
            Select Case True
                'Col = 2 >>> ponto Z que é igual a 0
                Case col = 2
                    pts(i) = 0
                
                Case Not IsError(ActiveCell.Offset(lin, col)) And col <> 2
                    pts(i) = ActiveCell.Offset(lin, col)
                
                'Tratando paredes com Viga acima
                Case IsError(ActiveCell.Offset(lin, col))
                    pts(i) = 10 ^ 12
                    
            End Select
            
        Next
        
        'Desenhando no cad e atualizando visualização
        'Tratando paredes com Viga acima
        If Not IsError(ActiveCell.Value) Then
            Set POLYLINE = AutoCAD.Application.ActiveDocument.ModelSpace.AddPolyline(pts)
            POLYLINE.Layer = Application.WorksheetFunction.XLookup("DrawCANALETA_ESQ", Worksheets("Layers").Range("C:C"), Worksheets("Layers").Range("A:A"))
            POLYLINE.Closed = False
            AutoCAD.Application.Update
            
            'Proxima parede
            ActiveCell.Offset(3, 0).Select
        Else
            'Proxima parede
            ActiveCell.Offset(3, 0).Select
        End If
        
    Next
    

End Sub
