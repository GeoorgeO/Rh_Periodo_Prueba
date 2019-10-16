Imports System.Data.SqlClient

Public Class Form1
    Dim valor As String = ""
    Private Sub Form1_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        recargar()
    End Sub

    Function calis(ByVal nombre As String, ByVal fecha As Date)
        Dim dia As Integer = 0
        Dim mes As Integer = 0
        Dim anio As Integer = 0
        Dim fechaC As Date

        dia = CDate(fecha).Date.Day
        mes = CDate(fecha).Date.Month
        anio = CDate(fecha).Date.Year
        If mes <= 9 Then
            fechaC = dia & "/" & mes + 3 & "/" & anio
        Else
            fechaC = dia & "/" & 3 - (12 - mes) & "/" & anio + 1
        End If

        MessageBox.Show(fechaC, "My Application", MessageBoxButtons.OKCancel, MessageBoxIcon.Asterisk)

        fechaC = fecha.AddMonths(+3)
        MessageBox.Show(fechaC, "My Application", MessageBoxButtons.OKCancel, MessageBoxIcon.Asterisk)

        Return Nothing
    End Function

    Function agregaTabla(ByVal nombre As String, ByVal fecha As Date)
        'Dim dia As Integer = 0
        'Dim mes As Integer = 0
        'Dim anio As Integer = 0
        Dim fechaC As Date
        Dim ndays As Int32 = 0

        'dia = CDate(fecha).Date.Day
        'mes = CDate(fecha).Date.Month
        'anio = CDate(fecha).Date.Year
        'If mes <= 9 Then
        '    fechaC = dia & "/" & mes + 3 & "/" & anio
        'Else
        '    fechaC = dia & "/" & 3 - (12 - mes) & "/" & anio + 1
        'End If
        ndays = CInt(ComboBox2.GetItemText(ComboBox2.SelectedItem)) - 1

        fechaC = fecha.AddDays(ndays)
        'fechaC = fecha.AddMonths(+3)

        Try
            Using cnx = New SqlConnection("Data Source=192.168.3.254;Initial Catalog=Vistas;User ID=sa;Password=inventumc762$")
                cnx.Open()
                Using cmd As New SqlCommand()
                    cmd.Connection = cnx
                    cmd.CommandText = "SET DATEFORMAT dmy   insert into RH_periodoPrueba (Nombre,fechaInicio,fechaFin,NumDays) values (@nombre,@fecha,@fCumple,@ndays)"
                    cmd.Parameters.AddWithValue("@nombre", nombre)
                    cmd.Parameters.AddWithValue("@fecha", fecha)
                    cmd.Parameters.AddWithValue("@fCumple", fechaC)
                    cmd.Parameters.AddWithValue("@ndays", ndays)


                    cmd.ExecuteNonQuery()

                    Dim Litems As New ListViewItem

                    Litems = ListView1.Items.Add(nombre)
                    Litems.SubItems.Add(fechaC)
                    Litems.SubItems.Add(ndays)
                End Using
                cnx.Close()
            End Using

        Catch ex As Exception
            MessageBox.Show(ex.Message, "HA ocurrido un error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

        Return Nothing
    End Function


    Function recargar()
        ComboBox2.SelectedIndex = 1

        ListView1.Clear() 'Limpiamos el ListView
        ListView1.View = View.Details 'Tipo de vista
        ListView1.FullRowSelect = True 'Al seleccionar un elemento, seleccionar la línea completa
        ListView1.GridLines = True 'Mostrar las líneas de la cuadrícula
        ListView1.LabelEdit = False 'No permitir la edición automática del texto
        ListView1.MultiSelect = False 'Permitir múltiple selección
        ListView1.HideSelection = False 'Para que al perder el foco, se siga viendo el que está seleccionado
        ListView1.ShowGroups = False 'Listado NO Agrupado

        ListView1.Columns.Add("Nombre", 270, HorizontalAlignment.Left)
        ListView1.Columns.Add("Fin de periodo prueba", 120, HorizontalAlignment.Left)
        ListView1.Columns.Add("Num. Dias", 70, HorizontalAlignment.Left)
        ListView1.View = View.Details

        cargarTabla()
        Return Nothing
    End Function


    Function AgregaLista(ByVal nombre As String, ByVal fecha As String, ByVal ndays As Int32)

        Dim Litems As New ListViewItem

        Litems = ListView1.Items.Add(nombre)
        Litems.SubItems.Add(CDate(fecha))
        Litems.SubItems.Add(CInt(ndays))

        Return Nothing
    End Function


    Function cargarTabla()
        Dim nombre As String = ""
        Dim fecha As String = ""
        Dim numdays As Int32 = 0
        Dim lector As SqlDataReader
        Try
            Using cnx = New SqlConnection("Data Source=192.168.3.254;Initial Catalog=Vistas;User ID=sa;Password=inventumc762$")
                cnx.Open()
                Using cmd As New SqlCommand()
                    cmd.Connection = cnx
                    cmd.CommandText = "SET DATEFORMAT dmy   select Nombre, CONVERT(CHAR(10), fechaFin, 111) as fechaFin,isnull(NumDays,0) as NumDays from RH_periodoPrueba order by fechaFin"
                    'cmd.ExecuteNonQuery()

                    lector = cmd.ExecuteReader
                    While lector.Read

                        nombre = CStr(lector(0).ToString)
                        fecha = CStr(lector(1).ToString)
                        numdays = CInt(lector(2).ToString)


                        AgregaLista(nombre, fecha, numdays)
                    End While
                    lector.Close()

                End Using
                cnx.Close()
            End Using

        Catch ex As Exception
            MessageBox.Show(ex.Message, "HA ocurrido un error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

        Return Nothing
    End Function

    Private Sub Label3_Click(sender As System.Object, e As System.EventArgs) Handles Label3.Click
        If TextBox1.Text <> "" Then
            agregaTabla(TextBox1.Text, CDate(DateTimePicker1.Text))
            'calis(TextBox1.Text, CDate(DateTimePicker1.Text))
            TextBox1.Text = ""
            DateTimePicker1.Value = Today
        End If
    End Sub

    Private Sub ListView1_SelectedIndexChanged(sender As System.Object, e As System.EventArgs) Handles ListView1.SelectedIndexChanged
        Try
            valor = ListView1.SelectedItems(0).Text
        Catch ex As Exception
        End Try
    End Sub


    Function eliminarTabla(ByVal nombre As String)
        Try
            Using cnx = New SqlConnection("Data Source=192.168.3.254;Initial Catalog=Vistas;User ID=sa;Password=inventumc762$")
                cnx.Open()
                Using cmd As New SqlCommand()
                    cmd.Connection = cnx
                    cmd.CommandText = "delete from RH_periodoPrueba where Nombre='" & nombre & "'"
                    cmd.ExecuteNonQuery()

                    recargar()
                End Using
                cnx.Close()
            End Using

        Catch ex As Exception
            MessageBox.Show(ex.Message, "HA ocurrido un error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
        Return Nothing
    End Function

    Private Sub Label4_Click(sender As System.Object, e As System.EventArgs) Handles Label4.Click
        eliminarTabla(valor)
    End Sub

    Function recargarbusca()
        ListView1.Clear() 'Limpiamos el ListView
        ListView1.View = View.Details 'Tipo de vista
        ListView1.FullRowSelect = True 'Al seleccionar un elemento, seleccionar la línea completa
        ListView1.GridLines = True 'Mostrar las líneas de la cuadrícula
        ListView1.LabelEdit = False 'No permitir la edición automática del texto
        ListView1.MultiSelect = False 'Permitir múltiple selección
        ListView1.HideSelection = False 'Para que al perder el foco, se siga viendo el que está seleccionado
        ListView1.ShowGroups = False 'Listado NO Agrupado

        ListView1.Columns.Add("Nombre", 270, HorizontalAlignment.Left)
        ListView1.Columns.Add("Fin de periodo prueba", 120, HorizontalAlignment.Left)
        ListView1.View = View.Details

        cargarTablabusca()
        Return Nothing
    End Function

    Function cargarTablabusca()
        Dim nombre As String = ""
        Dim fecha As String = ""
        Dim numdays As Int32 = 0
        Dim lector As SqlDataReader
        Try
            Using cnx = New SqlConnection("Data Source=192.168.3.254;Initial Catalog=Vistas;User ID=sa;Password=inventumc762$")
                cnx.Open()
                Using cmd As New SqlCommand()
                    cmd.Connection = cnx
                    cmd.CommandText = "SET DATEFORMAT dmy   select Nombre, CONVERT(CHAR(10), fechaFin, 111) as fechaFin, isnull(NumDays,0) as NumDays from RH_periodoPrueba where Nombre LIKE '" & TextBox2.Text & "%' Order by fechaFin"
                    'cmd.ExecuteNonQuery()

                    lector = cmd.ExecuteReader
                    While lector.Read

                        nombre = CStr(lector(0).ToString)
                        fecha = CStr(lector(1).ToString)
                        numdays = CInt(lector(2).ToString)


                        AgregaLista(nombre, fecha, numdays)
                    End While
                    lector.Close()

                End Using
                cnx.Close()
            End Using

        Catch ex As Exception
            MessageBox.Show(ex.Message, "HA ocurrido un error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

        Return Nothing
    End Function

    Function recargarmes(ByVal mes As Integer, ByVal anio As Integer)
        ListView1.Clear() 'Limpiamos el ListView
        ListView1.View = View.Details 'Tipo de vista
        ListView1.FullRowSelect = True 'Al seleccionar un elemento, seleccionar la línea completa
        ListView1.GridLines = True 'Mostrar las líneas de la cuadrícula
        ListView1.LabelEdit = False 'No permitir la edición automática del texto
        ListView1.MultiSelect = False 'Permitir múltiple selección
        ListView1.HideSelection = False 'Para que al perder el foco, se siga viendo el que está seleccionado
        ListView1.ShowGroups = False 'Listado NO Agrupado

        ListView1.Columns.Add("Nombre", 270, HorizontalAlignment.Left)
        ListView1.Columns.Add("Fin de periodo prueba", 120, HorizontalAlignment.Left)
        ListView1.View = View.Details

        cargarTablames(mes, anio)
        Return Nothing
    End Function

    Function cargarTablames(ByVal mes As Integer, ByVal anio As Integer)
        Dim nombre As String = ""
        Dim fecha As String = ""
        Dim numdays As Int32 = 0
        Dim lector As SqlDataReader
        Try
            Using cnx = New SqlConnection("Data Source=192.168.3.254;Initial Catalog=Vistas;User ID=sa;Password=inventumc762$")
                cnx.Open()
                Using cmd As New SqlCommand()
                    cmd.Connection = cnx
                    Dim messig, aniosig As Integer
                    If mes < 12 Then
                        messig = mes + 1
                        aniosig = anio
                    End If
                    If mes = 12 Then
                        messig = 1

                        aniosig = anio + 1
                    End If

                    cmd.CommandText = "SET DATEFORMAT dmy   select Nombre, CONVERT(CHAR(10), fechaFin, 111) as fechaFin,isnull(NumDays,0) as NumDays from RH_periodoPrueba where fechaFin>='01-" & mes & "-" & anio & "' and fechaFin<'01-" & messig & "-" & aniosig & "' Order by fechaFin"
                    'cmd.ExecuteNonQuery()

                    lector = cmd.ExecuteReader
                    While lector.Read

                        nombre = CStr(lector(0).ToString)
                        fecha = CStr(lector(1).ToString)
                        numdays = CInt(lector(2).ToString)

                        AgregaLista(nombre, fecha, numdays)
                    End While
                    lector.Close()

                End Using
                cnx.Close()
            End Using

        Catch ex As Exception
            MessageBox.Show(ex.Message, "HA ocurrido un error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

        Return Nothing
    End Function

    Private Sub TextBox2_GotFocus(sender As Object, e As System.EventArgs) Handles TextBox2.GotFocus
        If TextBox2.Text = "Buscar por nombre" Then
            TextBox2.ForeColor = Color.Black
            TextBox2.Text = ""
        End If
    End Sub

    Private Sub TextBox2_KeyUp(sender As Object, e As System.Windows.Forms.KeyEventArgs) Handles TextBox2.KeyUp
        recargarbusca()
        If TextBox2.Text <> "" Then
            Label5.Visible = True
        Else
            Label5.Visible = False
        End If
    End Sub


    Private Sub TextBox2_LostFocus(sender As Object, e As System.EventArgs) Handles TextBox2.LostFocus
        TextBox2.ForeColor = Color.Gray
        TextBox2.Text = "Buscar por nombre"
        Label5.Visible = False
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As System.Object, e As System.EventArgs) Handles ComboBox1.SelectedIndexChanged
        If ComboBox1.SelectedIndex = 0 Then
            recargarmes(1, Now.Date.Year)
        End If
        If ComboBox1.SelectedIndex = 1 Then
            recargarmes(2, Now.Date.Year)
        End If
        If ComboBox1.SelectedIndex = 2 Then
            recargarmes(3, Now.Date.Year)
        End If
        If ComboBox1.SelectedIndex = 3 Then
            recargarmes(4, Now.Date.Year)
        End If
        If ComboBox1.SelectedIndex = 4 Then
            recargarmes(5, Now.Date.Year)
        End If
        If ComboBox1.SelectedIndex = 5 Then
            recargarmes(6, Now.Date.Year)
        End If
        If ComboBox1.SelectedIndex = 6 Then
            recargarmes(7, Now.Date.Year)
        End If
        If ComboBox1.SelectedIndex = 7 Then
            recargarmes(8, Now.Date.Year)
        End If
        If ComboBox1.SelectedIndex = 8 Then
            recargarmes(9, Now.Date.Year)
        End If
        If ComboBox1.SelectedIndex = 9 Then
            recargarmes(10, Now.Date.Year)
        End If
        If ComboBox1.SelectedIndex = 10 Then
            recargarmes(11, Now.Date.Year)
        End If
        If ComboBox1.SelectedIndex = 11 Then
            recargarmes(12, Now.Date.Year)
        End If
        If ComboBox1.SelectedIndex = 12 Then
            recargar()
        End If
    End Sub


    Private Sub Label5_Click(sender As System.Object, e As System.EventArgs) Handles Label5.Click
        TextBox2.Text = ""
        recargarbusca()
        Label5.Visible = False
    End Sub

   

End Class
