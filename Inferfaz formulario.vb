

Private Sub CommandButton1_Click()

    Call Formulario.Salir

End Sub

Private Sub CommandButton2_Click()

    Call Formulario.Limpiar

End Sub

Private Sub CommandButton4_Click()

    With Me.Lista
    
    Hoja3.Rows(.ListIndex + 1).Delete
    
    End With

End Sub

Private Sub CommandButton5_Click()

    Call Formulario.Guardar
    
End Sub


Private Sub Lista_Click()

    Dim codigo As Integer
    
    codigo = Lista.List(Lista.ListIndex, 0)
    Me.TextBox1.Value = codigo
    
    Dim idinicial As Integer
    
    idinicial = TextBox1.Value
    Me.TextBox2 = Application.WorksheetFunction.VLookup(idinicial, Sheets("Clientes").Range("A:H"), 2, 0)
    Me.TextBox3 = Application.WorksheetFunction.VLookup(idinicial, Sheets("Clientes").Range("A:H"), 3, 0)
    Me.TextBox4 = Application.WorksheetFunction.VLookup(idinicial, Sheets("Clientes").Range("A:H"), 5, 0)
    Me.TextBox5 = Application.WorksheetFunction.VLookup(idinicial, Sheets("Clientes").Range("A:H"), 6, 0)
    Me.TextBox6 = Application.WorksheetFunction.VLookup(idinicial, Sheets("Clientes").Range("A:H"), 4, 0)
    Me.TextBox7 = Application.WorksheetFunction.VLookup(idinicial, Sheets("Clientes").Range("A:H"), 7, 0)
    Me.TextBox8 = Application.WorksheetFunction.VLookup(idinicial, Sheets("Clientes").Range("A:H"), 8, 0)

End Sub


Private Sub TextBox1_Change()

End Sub

Private Sub UserForm_Activate()

    Me.Lista.RowSource = "Clientes"
    Me.Lista.ColumnCount = 8
    
End Sub