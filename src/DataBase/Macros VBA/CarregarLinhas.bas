Attribute VB_Name = "CarregarLinhas"
Sub CarregarLinhasBase()

'Atualização da tela: DESATIVADO (pro usuário não achar q o pc vai explodir)
Application.ScreenUpdating = False

'Calcular Automáticamente: DESATIVADO (Melhora de FPS)
Application.Calculation = xlManual

'LIMPANDO TABELA
Sheets("Banco de Dados").Select
    Range("LinBase_inicio").Select
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
    
    Select Case Obj
    
        Case "AcDbLine"
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
            Tipo = "LINE"
            
            'Layer
            Layer = ssetObj.Item(i).Layer
            
         Case "AcDbPolyline"
            'Pontos Iniciais
            pts = ssetObj.Item(i).Coordinates
            ptsX = Round(pts(0), 4)
            ptsY = Round(pts(1), 4)
            
            'Pontos Finais
            n = ArrayLen(pts)
            pteX = Round(pts(n - 2), 4)
            pteY = Round(pts(n - 1), 4)
            
            Xi = Application.WorksheetFunction.Min(ptsX, pteX)
            Yi = Application.WorksheetFunction.Min(ptsY, pteY)
            Xf = Application.WorksheetFunction.Max(ptsX, pteX)
            Yf = Application.WorksheetFunction.Max(ptsY, pteY)
            
            'Texto
            Texto = ""
            
            'Tipo de Objeto
            Tipo = "POLYLINE"
            
            'Layer
            Layer = ssetObj.Item(i).Layer
            
         Case "AcDbCircle"
            'Pontos Iniciais
            pts = ssetObj.Item(i).center
            Xi = Round(pts(0), 4)   'Xi
            Yi = Round(pts(1), 4)   'Yi
            
            'Pontos Finais
            Xf = ""
            Yf = ""
            
            'Texto
            Texto = ""
            
            'Tipo de Objeto
            Tipo = "CIRCLE"
            
            'Layer
            Layer = ssetObj.Item(i).Layer
            
        Case "AcDbText"
            'Pontos Iniciais
            pts = ssetObj.Item(i).InsertionPoint
            Xi = Round(pts(0), 4)   'Xi
            Yi = Round(pts(1), 4)   'Yi
            
            'Pontos Finais
            n = ArrayLen(pts)
            Xf = ""
            Yf = ""
            
            'Texto
            Texto = ssetObj.Item(i).TextString
            
            'Tipo de Objeto
            Tipo = "TEXT"
            
            'Layer
            Layer = ssetObj.Item(i).Layer
            
        End Select
    
        'PLOTAGEM
        Sheets("Banco de Dados").Select
        Range("Q6").Select
        ActiveCell.Offset(i, 0).Value = Xi
        ActiveCell.Offset(i, 1).Value = Yi
        ActiveCell.Offset(i, 2).Value = Xf
        ActiveCell.Offset(i, 3).Value = Yf
        ActiveCell.Offset(i, 4).Value = Texto
        ActiveCell.Offset(i, 5).Value = Tipo
        ActiveCell.Offset(i, 6).Value = Layer
    
    Next
    
'Calcular Automáticamente: ATIVADO
Application.Calculation = xlAutomatic

'Atualização da tela: ATIVADO
Application.ScreenUpdating = True

'Voltando para a planilha de Controle
Sheets("Controle").Select
Range("C5").Select

End Sub

