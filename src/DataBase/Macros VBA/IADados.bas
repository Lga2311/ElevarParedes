Attribute VB_Name = "IADados"
Sub IA_GerarLinFiadas()

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
    CicloLim = Range("EA1").Value
    
    'Selecionando primeiro X
    Range("EE6").Select
    
    'Desenhar GABARITO
    For Ciclo = 0 To CicloLim - 1
    
        nFiadas = ActiveCell.Offset(0, -1).Value
        
        For n = 0 To nFiadas - 1
        
            deltaY = 0.2 * n
            'Pesquisar pontos no Excel
            Inicio(0) = ActiveCell.Offset(0, 0)
            Inicio(1) = ActiveCell.Offset(0, 1) + deltaY
            Inicio(2) = 0
            Fim(0) = ActiveCell.Offset(1, 0)
            Fim(1) = ActiveCell.Offset(1, 1) + deltaY
            Fim(2) = 0
            
            'Desenhando no cad e atualizando visualização
            Set LINE = AutoCAD.Application.ActiveDocument.ModelSpace.AddLine(Inicio, Fim)
            LINE.Layer = Application.WorksheetFunction.XLookup("LinFiadas_IA", Worksheets("Layers").Range("C:C"), Worksheets("Layers").Range("A:A"))
    
            AutoCAD.Application.Update
        Next
        
        ActiveCell.Offset(3, 0).Select
        
    Next
    
    'mydwg.SendCommand "_Trim"

End Sub
Sub IA_LerLinFiadas()


'Atualização da tela: DESATIVADO (pro usuário não achar q o pc vai explodir)
Application.ScreenUpdating = False

'Calcular Automáticamente: DESATIVADO (Melhora de FPS)
Application.Calculation = xlManual

'LIMPANDO TABELA
Sheets("Fiadas").Select
    Range("LinFiadas_inicio").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.ClearContents

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


    'ID Seleção
    Dim sDataTime As String
    sDataTime = Now()
    
    'Criar uma seleção
    Dim ssetObj As AcadSelectionSet
    Set ssetObj = AutoCAD.Application.ActiveDocument.SelectionSets.Add(sDataTime)

   'Adicionar objetos a um conjunto de seleção, solicitando que o usuário selecione na tela
    ssetObj.SelectOnScreen
    
    For i = 0 To ssetObj.Count - 1
    
    Obj = ssetObj.Item(i).ObjectName
    Lay = ssetObj.Item(i).Layer
    
    If Lay = "FIADAS" Then
        'Pontos Iniciais
        pts = ssetObj.Item(i).StartPoint
        ptsX = Round(pts(0), 4)
        ptsY = Round(pts(1), 4)
        
        'Pontos Finais
        pte = ssetObj.Item(i).EndPoint
        pteX = Round(pte(0), 4)
        pteY = Round(pte(1), 4)
        
        Xi = Application.WorksheetFunction.Min(ptsX, pteX)
        Yi = Application.WorksheetFunction.Min(ptsY, pteY)
        Xf = Application.WorksheetFunction.Max(ptsX, pteX)
        Yf = Application.WorksheetFunction.Max(ptsY, pteY)
        
        'Texto
        Texto = ""
        
        'Tipo de Objeto
        If Obj = "AcDbLine" Then
            Tipo = "LINE"
        ElseIf Obj = "AcDbPolyline" Then
            Tipo = "POLYLINE"
        End If
        
        'Layer
        Layer = ssetObj.Item(i).Layer
            
    End If
    
    'PLOTAGEM
    Sheets("Fiadas").Select
    Range("LinFiadas_PrimCell").Select
    ActiveCell.Offset(i, 1).Value = Xi
    ActiveCell.Offset(i, 2).Value = Yi
    ActiveCell.Offset(i, 3).Value = Xf
    ActiveCell.Offset(i, 4).Value = Yf
    ActiveCell.Offset(i, 5).Value = Tipo
    ActiveCell.Offset(i, 6).Value = Layer
    
    Next
    
'Calcular Automáticamente: ATIVADO
Application.Calculation = xlAutomatic

'Atualização da tela: ATIVADO
Application.ScreenUpdating = True

'Voltando para a planilha de Controle
'Sheets("Controle").Select
'Range("C5").Select

End Sub
Sub LerBlocos()

'Atualização da tela: DESATIVADO (pro usuário não achar q o pc vai explodir)
'Application.ScreenUpdating = False

'Calcular Automáticamente: DESATIVADO (Melhora de FPS)
Application.Calculation = xlManual

'LIMPANDO TABELA
Sheets("Blocos").Select
    Range("Blocos_inicio").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.ClearContents

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


    'ID Seleção
    Dim sDataTime As String
    sDataTime = Now()
    
    'Criar uma seleção
    Dim ssetObj As AcadSelectionSet
    Set ssetObj = AutoCAD.Application.ActiveDocument.SelectionSets.Add(sDataTime)

   'Adicionar objetos a um conjunto de seleção, solicitando que o usuário selecione na tela
    ssetObj.SelectOnScreen
    
    'Selecionando Celula inicial
    Sheets("Blocos").Select
    Range("Blocos_PrimCell").Select
    
    For i = 0 To ssetObj.Count - 1
    
    Obj = ssetObj.Item(i).ObjectName
    Lay = ssetObj.Item(i).Layer
        
    If Obj = "AcDbPolyline" Then
        'Pontos x & y da Polyline
        Dim x(3), y(3) As Single
        pts = ssetObj.Item(i).Coordinates
        x(0) = Round(pts(0), 4): y(0) = Round(pts(1), 4)
        x(1) = Round(pts(2), 4): y(1) = Round(pts(3), 4)
        x(2) = Round(pts(4), 4): y(2) = Round(pts(5), 4)
        x(3) = Round(pts(6), 4): y(3) = Round(pts(7), 4)
        
        'CORREÇÃO DOS PONTOS, PARA INSERIR AS VISTAS
        'Definindo pontos base
        Xmin = Min(x): Ymin = Min(y)
        Xmax = Max(x): Ymax = Max(y)
        
        'Pontos Iniciais
        ptsX = Round(Xmin, 4)
        ptsY = Round(Ymin, 4)
        
        'Pontos Finais
        pteX = Round(Xmax, 4)
        pteY = Round(Ymax, 4)
        
        'Definindo pontos iniciais e finais
        Xi = Application.WorksheetFunction.Min(ptsX, pteX)
        Yi = Application.WorksheetFunction.Min(ptsY, pteY)
        Xf = Application.WorksheetFunction.Max(ptsX, pteX)
        Yf = Application.WorksheetFunction.Max(ptsY, pteY)
        
        'Tipo de Objeto
        Tipo = "POLYLINE"
        
        'Layer
        Layer = ssetObj.Item(i).Layer
    End If
    
    'PLOTAGEM
    ActiveCell.Offset(0, 1).Value = Xi
    ActiveCell.Offset(0, 2).Value = Yi
    ActiveCell.Offset(0, 3).Value = Xf
    ActiveCell.Offset(0, 4).Value = Yf
    ActiveCell.Offset(0, 5).Value = Tipo
    ActiveCell.Offset(0, 6).Value = Layer
    
    'Selecionando Cell próxima da Linha
    ActiveCell.Offset(1, 0).Select

    Next
    
'Calcular Automáticamente: ATIVADO
Application.Calculation = xlAutomatic

'Atualização da tela: ATIVADO
'Application.ScreenUpdating = True

'Voltando para a planilha de Controle
'Sheets("Controle").Select
'Range("C5").Select

End Sub

