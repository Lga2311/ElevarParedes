Attribute VB_Name = "CarregarConfig"
Sub Config()

'Atualiza��o da tela: DESATIVADO (pro usu�rio n�o achar q o pc vai explodir)
Application.ScreenUpdating = False

'Calcular Autom�ticamente: DESATIVADO (Melhora de FPS)
Application.Calculation = xlManual

'LIMPANDO TABELA
Sheets("Banco de Dados").Select
    Range("B5:D5").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.ClearContents

'Vari�veis do Cad para selecionar o programa
Dim myapp As Object
Dim mydwg As AcadDocument

'Verifica��o de Erro pra alguma coisa (pesquisar)
    Set myapp = GetObject(, "Autocad.application")
ERRORHANDLER:
    If Err.Description <> "" Then
        Err.Clear

        Set myapp = CreateObject("Autocad.application")
    End If

'Selecionar cad aberto
myapp.Visible = True


    'ID Sele��o
    Dim sDataTime As String
    sDataTime = Now()
    
    'Criar uma sele��o
    Dim ssetObj As AcadSelectionSet
    Set ssetObj = AutoCAD.Application.ActiveDocument.SelectionSets.Add(sDataTime)

   'Adicionar objetos a um conjunto de sele��o, solicitando que o usu�rio selecione na tela
    ssetObj.SelectOnScreen
    
    For i = 0 To ssetObj.Count - 1
    
    Obj = ssetObj.Item(i).ObjectName
    'Item = ssetObj.Item(i)
    
        If Obj = "AcDbText" Then
        
            'Coletando dados dos textos com as configura��es
            Texto = ssetObj.Item(i).TextString
            Layer = ssetObj.Item(i).Layer
            Tipo = Mid(Texto, 2, 1)
            Itens = Mid(Texto, 6, 1000)
            
            'Imprimindo na tabela do cad
            Sheets("Banco de Dados").Select
            Range("b5").Select
            
            If Mid(Texto, 1, 1) = "T" Then
                ActiveCell.Offset(i, 0).Value = Texto
                ActiveCell.Offset(i, 1).Value = Layer
                ActiveCell.Offset(i, 2).Value = Tipo
            End If
            
        End If

    Next
    
'Calcular Autom�ticamente: ATIVADO
Application.Calculation = xlAutomatic

'Atualiza��o da tela: ATIVADO
Application.ScreenUpdating = True

Sheets("Controle").Select
Range("C5").Select

End Sub
