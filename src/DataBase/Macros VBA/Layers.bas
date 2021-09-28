Attribute VB_Name = "Layers"
Sub LayersElevarPar()

'===========================================================================================================================
'New_Layer
'Fonte: 'Menssage 2 - Vladimir
        'https://forums.autodesk.com/t5/autocad-mechanical-forum/create-new-layer-in-autocad-using-excel-vba/td-p/7431095
        'Acessado: 14/07/2021
'Autor: LucasGA
'Ultima Atualização: 14/07/2021
'===========================================================================================================================

Dim acadApp As AcadApplication
Dim acadDoc As AcadDocument

Dim newLayer As AcadLayer
Dim layerName As String

'Verifique se o AutoCAD está aberto.
On Error Resume Next
    Set acadApp = GetObject(, "AutoCAD.Application")
On Error GoTo 0
    

'Se o AutoCAD não estiver aberto, crie uma nova instância e torne-a visível.
If acadApp Is Nothing Then
    Set acadApp = New AcadApplication
    acadApp.Visible = True
End If

'Verifique se há um desenho ativo.
On Error Resume Next
    Set acadDoc = acadApp.ActiveDocument
On Error GoTo 0

'Nenhum desenho ativo encontrado. Crie um novo.
If acadDoc Is Nothing Then
    Set acadDoc = acadApp.Documents.Add
    acadApp.Visible = True
End If

'Verifique se a camada já existe e se não fizer uma nova
On Error Resume Next

    'Abrindo planilha com as Layers Armazenadas
    Sheets("Layers").Select
    Range("A2").Select
    i = 0
    
    Do While ActiveCell.Offset(i, 0) <> ""
    
        layerName = ActiveCell.Offset(i, 0).Value
        Set newLayer = acadDoc.Layers.Add(layerName)
        newLayer.Color = ActiveCell.Offset(i, 1).Value
        i = i + 1
        
    Loop
    
On Error GoTo 0

If newLayer Is Nothing Then
    Set newLayer = acadDoc.Layers.Add(layerName)
End If
'acadDoc.ActiveLayer = newLayer 'DEU ALGUM ERRO! (07/08/21 01:27)


End Sub
