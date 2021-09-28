Attribute VB_Name = "DrawTextCAD"
Sub TextNFiadas()

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
Dim Text As AcadText  'CAD
Dim Circ As AcadCircle 'CAD

    'TEXTOS
    'Variáveis: Pontos e Ciclos do FOR
    Dim ptsText(0 To 2) As Double
    Dim ptsCirc(0 To 2) As Double
    Dim Nome As String
    Dim CicloLim As Single
    
    'Localizando o limite do FOR
    Sheets("Banco de Dados").Select
    CicloLim = Application.WorksheetFunction.Max(Range("DQ:DQ").Value)
    
    'Selecionando primeiro X
    Range("DU6").Select
    
    'Desenhar N FIADAS
    For Ciclo = 0 To CicloLim - 1
    
        'Pesquisar pontos no Excel
        ptsText(0) = ActiveCell.Offset(0, 0).Value
        ptsText(1) = ActiveCell.Offset(0, 1).Value
        ptsText(2) = 0
        
        ptsCirc(0) = ActiveCell.Offset(0, 2).Value
        ptsCirc(1) = ActiveCell.Offset(0, 3).Value
        ptsCirc(2) = 0
            
        'Nome da Parede
        NfiadaMax = ActiveCell.Offset(0, -1).Value
        
        For Nfiada = 1 To NfiadaMax
        
            If Nfiada < 10 Then Nome = 0 & Nfiada Else Nome = Nfiada
            
            'Desenhando no cad e atualizando visualização
            'Texto
            Set Text = AutoCAD.Application.ActiveDocument.ModelSpace.AddText(Nome, ptsText, 0.05)
            Text.Layer = Application.WorksheetFunction.XLookup("TextNFiadas", Worksheets("Layers").Range("C:C"), Worksheets("Layers").Range("A:A"))
            Text.Color = acYellow
            
            'Texto
            Set Circ = AutoCAD.Application.ActiveDocument.ModelSpace.AddCircle(ptsCirc, 0.065)
            Circ.Layer = Application.WorksheetFunction.XLookup("TextNFiadas", Worksheets("Layers").Range("C:C"), Worksheets("Layers").Range("A:A"))
            
            AutoCAD.Application.Update
            
            ptsText(1) = ptsText(1) + 0.2
            ptsCirc(1) = ptsCirc(1) + 0.2
            
        Next
        
        ActiveCell.Offset(2, 0).Select
        
    Next

End Sub


Sub TextPAREDE()

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
Dim Text As AcadText  'CAD

    'Variáveis: Pontos e Ciclos do FOR
    Dim pts(0 To 2) As Double
    Dim Nome As String
    Dim CicloLim As Single
    
    'Localizando o limite do FOR
    Sheets("Banco de Dados").Select
    CicloLim = Application.WorksheetFunction.Max(Range("DJ:DJ").Value)
    
    'Selecionando primeiro X
    Range("DL6").Select
    
    'Desenhar GABARITO
    For Ciclo = 0 To CicloLim - 1
    
        'Nome da Parede
        Nome = ActiveCell.Offset(0, 2).Value

        'Pesquisar pontos no Excel
        For i = 0 To 2
                    
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
            Set Text = AutoCAD.Application.ActiveDocument.ModelSpace.AddText(Nome, pts, 0.15)
            Text.Layer = Application.WorksheetFunction.XLookup("TextPAREDE", Worksheets("Layers").Range("C:C"), Worksheets("Layers").Range("A:A"))
            AutoCAD.Application.Update
            
            'Proxima parede
            ActiveCell.Offset(2, 0).Select
        Else
            'Proxima parede
            ActiveCell.Offset(2, 0).Select
        End If
        
    Next

End Sub

Sub TextESQ()

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
Dim Text As AcadText  'CAD

    'Variáveis: Pontos e Ciclos do FOR
    Dim pts(0 To 2) As Double
    Dim Nome As String
    Dim CicloLim As Single
    
    'Localizando o limite do FOR
    Sheets("Banco de Dados").Select
    CicloLim = Application.WorksheetFunction.Max(Range("BO:BO").Value)
    
    'Selecionando primeiro X
    Range("BS6").Select
    
    'Desenhar GABARITO
    For Ciclo = 0 To CicloLim - 1
    
        'Nome da Parede
        Nome = ActiveCell.Offset(0, 2).Value

        'Pesquisar pontos no Excel
        For i = 0 To 2
                    
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
            Set Text = AutoCAD.Application.ActiveDocument.ModelSpace.AddText(Nome, pts, 0.1)
            Text.Layer = Application.WorksheetFunction.XLookup("TextESQ", Worksheets("Layers").Range("C:C"), Worksheets("Layers").Range("A:A"))
            AutoCAD.Application.Update
            
            'Proxima parede
            ActiveCell.Offset(6, 0).Select
        Else
            'Proxima parede
            ActiveCell.Offset(6, 0).Select
        End If
        
    Next

End Sub

Sub TextGROUTE()

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
Dim Text As AcadText  'CAD

    'Variáveis: Pontos e Ciclos do FOR
    Dim pts(0 To 2) As Double
    Dim Nome As String
    Dim CicloLim As Single
    
    'Localizando o limite do FOR
    Sheets("Banco de Dados").Select
    CicloLim = Range("Bx1").Value
    
    'Selecionando primeiro X
    Range("CC7").Select
    
    'Desenhar GABARITO
    For Ciclo = 0 To CicloLim - 1
    
        'Nome da Parede
        Nome = ActiveCell.Offset(0, 2).Value

        'Pesquisar pontos no Excel
        For i = 0 To 2
                    
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
            Set Text = AutoCAD.Application.ActiveDocument.ModelSpace.AddText(Nome, pts, 0.1)
            Text.Layer = Application.WorksheetFunction.XLookup("TextGROUTE", Worksheets("Layers").Range("C:C"), Worksheets("Layers").Range("A:A"))
            Text.Rotation = Application.WorksheetFunction.Pi() / 2
            AutoCAD.Application.Update
            
            'Proxima parede
            ActiveCell.Offset(3, 0).Select
        Else
            'Proxima parede
            ActiveCell.Offset(3, 0).Select
        End If
        
    Next

End Sub
