Attribute VB_Name = "Navegacao"
Sub VerPrevia()
Attribute VerPrevia.VB_ProcData.VB_Invoke_Func = " \n14"

    Sheets("Prévia").Select

End Sub

Sub VerControle()

    Sheets("Controle").Select
    Range("C5").Select

End Sub


Sub VerTutrial()

    Sheets("Tutorial").Select
    Range("B3").Select

End Sub
