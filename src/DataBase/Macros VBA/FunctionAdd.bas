Attribute VB_Name = "FunctionAdd"
Public Function ArrayLen(arr As Variant) As Integer
    ArrayLen = UBound(arr) - LBound(arr) + 1
End Function

Public Function Minimo(v1, v2 As Single)

    Dim v1 As Double
    Dim v2 As Double

    Select Case True
    
        Case v1 > v2
            v2
            
        Case v2 > v1
            v1
    
    End Select
    
End Function

Public Function Maximo(v1, v2 As Single)

    Dim v1 As Double
    Dim v2 As Double

    Select Case True
    
        Case v1 > v2
            v1
            
        Case v2 > v1
            v2
    
    End Select
    
End Function

'===========================================================================================================================
'Épar & Éimpar
'Fonte: 'Menssage 2 - Graymalkin
        'https://forum.scriptbrasil.com.br/topic/70827-fun%C3%A7%C3%A3o-par-ou-impar/
        'Acessado: 27/11/2020
'Autor: LucasGA
'Ultima Atualização: 27/11/2020
'===========================================================================================================================

Public Function Épar(ByVal Number As Variant) As Boolean

    If Number Mod 2 = 0 Then
        Épar = True
    End If

End Function

Public Function Éimpar(ByVal Number As Variant) As Boolean

    If Number Mod 2 <> 0 Then
        Éimpar = True
    End If

End Function

Public Function ProcxLayer(Valor As String)

    Dim Pesq, Resul As String
    Dim Lay As Worksheet
    
    Sheets("Layers").Select
    Range("C1").Select
    
    i = 0
    
    Do While Valor <> Pesq
    
        Pesq = ActiveCell.Offset(i, 0)
        i = i + 1
        
    Loop
    
    layerobj = ActiveCell.Offset(i - 1, -2).Value
    
    ProcxLayer = layerobj
    
End Function
'===========================================================================================================================
'Min & Max
'Fonte: 'Menssage 2 of 3
        'https://forums.autodesk.com/t5/vba/max-function/td-p/341125
        'Acessado: 26/11/2020
'Autor: LucasGA
'Ultima Atualização: 26/11/2020
'===========================================================================================================================

'Send back the minimum value from a list of numbers
Public Function Min(ByVal NumericArray As Variant) As Variant

    Dim i As Integer
    Dim MinVal As Double
    Dim CurVal As Double
    
    CurVal = NumericArray(0)
    MinVal = CurVal
    
    For i = LBound(NumericArray) + 1 To UBound(NumericArray)
        CurVal = NumericArray(i)
        If CurVal < MinVal Then MinVal = CurVal
    Next
    
    Min = MinVal
    
End Function

'Send back the maximum value from a list of numbers
Public Function Max(ByVal NumericArray As Variant) As Variant
    
    Dim i As Integer
    Dim MaxVal As Double
    Dim CurVal As Double
    
    CurVal = NumericArray(0)
    MaxVal = CurVal
    
    For i = LBound(NumericArray) + 1 To UBound(NumericArray)
        CurVal = NumericArray(i)
        If CurVal > MaxVal Then MaxVal = CurVal
    Next
    
    Max = MaxVal
    
End Function

