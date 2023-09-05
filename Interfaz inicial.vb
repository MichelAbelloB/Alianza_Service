Option Explicit
Private codigobuscado As Range


Sub MostrarFormulario()
    
    FormularioClientes.Show
    
End Sub

Sub Salir()

    Unload FormularioClientes

End Sub

Sub Limpiar()

    FormularioClientes.TextBox1.Text = ""
    FormularioClientes.TextBox2.Text = ""
    FormularioClientes.TextBox3.Text = ""
    FormularioClientes.TextBox4.Text = ""
    FormularioClientes.TextBox5.Text = ""
    FormularioClientes.TextBox6.Text = ""
    FormularioClientes.TextBox7.Text = ""
    FormularioClientes.TextBox8.Text = ""
    
End Sub

Sub Guardar()

    Dim Rcodigos As Range
    Dim Nuevafila As Range
    
    If FormularioClientes.TextBox1.Text = "" Or FormularioClientes.TextBox2.Text = "" Or FormularioClientes.TextBox3.Text = "" Or FormularioClientes.TextBox4.Text = "" Or FormularioClientes.TextBox5.Text = "" Or FormularioClientes.TextBox6.Text = "" Or FormularioClientes.TextBox7.Text = "" Or FormularioClientes.TextBox8.Text = "" Then
    
    MsgBox "Rellenar todos los campos del formulario."
    
    Else
    
        Set Rcodigos = Hoja3.ListObjects("Tabla3").ListColumns(1).Range
        Set codigobuscado = Rcodigos.Find(what:=FormularioClientes.TextBox1.Text, after:=Hoja3.Range("A2"), lookat:=xlWhole)
        
        If codigobuscado Is Nothing Then
            Set Nuevafila = Hoja3.ListObjects("Tabla3").ListRows.Add(1).Range
            Nuevafila.Cells(1).Value = FormularioClientes.TextBox1.Text
            Nuevafila.Cells(2).Value = FormularioClientes.TextBox2.Text
            Nuevafila.Cells(3).Value = FormularioClientes.TextBox3.Text
            Nuevafila.Cells(5).Value = FormularioClientes.TextBox4.Text
            Nuevafila.Cells(6).Value = FormularioClientes.TextBox5.Text
            Nuevafila.Cells(4).Value = FormularioClientes.TextBox6.Text
            Nuevafila.Cells(7).Value = FormularioClientes.TextBox7.Text
            Nuevafila.Cells(8).Value = FormularioClientes.TextBox8.Text
        Else
            MsgBox "El ID ya existe."
        End If
    End If
    
    Call Formulario.Limpiar
    
End Sub