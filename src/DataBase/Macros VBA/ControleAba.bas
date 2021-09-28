Attribute VB_Name = "ControleAba"
Sub ProxAba()

    Sheets("Controle").Select
    
    Dim ID As Single
    ID = Range("PagAtual").Value
    
    Dim QuantAbas As Single
    QuantAbas = 5
    
    prox = ID + 1
    
    If prox > QuantAbas Then
    
        ID = 1
        Range("ProxPag").Value = ID
        
    Else
    
        Range("ProxPag").Value = prox
        
    End If

End Sub

Sub AbaAnterior()

    Sheets("Controle").Select
    
    Dim ID As Single
    ID = Range("PagAtual").Value
    
    Dim QuantAbas As Single
    QuantAbas = 5
    
    Ant = ID - 1
    
    If Ant = 0 Then
    
        ID = QuantAbas
        Range("ProxPag").Value = ID
        
    Else
    
        Range("ProxPag").Value = Ant
        
    End If

End Sub
