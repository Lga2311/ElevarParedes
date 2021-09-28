Attribute VB_Name = "ControleDesenhar"
Sub Desenhar()

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

    LayersElevarPar

    
'Atualização da tela: DESATIVADO (pro usuário não achar q o pc vai explodir)
Application.ScreenUpdating = False
    
    Dim DES_Tudo, DES_Gab, DES_Esq, DES_Groute, DES_Canaleta As Single
    Dim TEX_Tudo, TEX_Par, TEX_Esq, TEX_Groute As Single
    
    Sheets("Controle").Select
    
    'Carregando Dados: Desenhar
    Range("SelecaoDES").Select
    DES_Tudo = ActiveCell.Offset(0, 0).Value
    DES_Gab = ActiveCell.Offset(1, 0).Value
    DES_Esq = ActiveCell.Offset(2, 0).Value
    DES_Groute = ActiveCell.Offset(3, 0).Value
    DES_Canaleta = ActiveCell.Offset(4, 0).Value
    
    'Carregando Dados: Texto
    Range("SelecaoTEXT").Select
    TEX_Tudo = ActiveCell.Offset(0, 0).Value
    TEX_NFiada = ActiveCell.Offset(1, 0).Value
    TEX_Par = ActiveCell.Offset(2, 0).Value
    TEX_Esq = ActiveCell.Offset(3, 0).Value
    TEX_Groute = ActiveCell.Offset(4, 0).Value
    
    If Range("DadosDesenhar").Value = 1 Then
    
        'Imprimir desenhos
        If DES_Gab = 1 Then
            DrawGABARITO
            DrawLAJE
        End If
        If DES_Esq = 1 Then
            DrawESQ
        End If
        If DES_Groute = 1 Then
            DrawGROUTE
        End If
        If DES_Canaleta = 1 Then
            DrawCANALETA_ALV
            DrawCANALETA_VIGAS
            DrawCANALETA_ESQ
        End If
        
        'Imprimir textos
        If TEX_Par = 1 Then
          TextPAREDE
        End If
        If TEX_NFiada = 1 Then
            TextNFiadas
        End If
        If TEX_Esq = 1 Then
            TextESQ
        End If
        If TEX_Groute = 1 Then
            TextGROUTE
        End If
    
    Else
        MsgBox "Faltam dados para a plotagem no CAD ser realizada" & vbCrLf & _
               "Dados preenchidos: " & Range("DadosDesenhar").Value * 100 & "%"
               
    End If
    
    VerControle
    
'Atualização da tela: ATIVADO
Application.ScreenUpdating = True

End Sub
