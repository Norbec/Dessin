Public Class Attrib
    Dim curr_user
    Public ConfConnS As String = "Data Source=VOUTE01-NORBEC\NORBECDEV;Initial Catalog=Configurateur;Persist Security Info=True;User ID=coachstudio;Password=coachstudio"
    Public ConfConn As New SqlConnection(ConfConnS)

    Private Sub TasksBindingNavigatorSaveItem_Click(sender As Object, e As EventArgs)
        Me.Validate()
        Me.TasksBindingSource.EndEdit()
        Me.Tasks_Jobs_approBindingSource.EndEdit()
        Me.Tasks_Jobs_prodBindingSource.EndEdit()
        Me.TableAdapterManager.UpdateAll(Me.NorbecDataDataSet)
        Me.TableAdapterManager.UpdateAll(Me.NorbecDataDataSet1)

    End Sub

    Private Sub FillbycontratsoumToolStripButton_Click(sender As Object, e As EventArgs) Handles FillbycontratsoumToolStripButton.Click
        Try
            Dim stringed_order, stringed_soum As String
            If Strings.Len(ContratToolStripTextBox.Text) > 4 Then stringed_order = ContratToolStripTextBox.Text & "%"
            If Strings.Len(SoumissionToolStripTextBox.Text) > 4 Then stringed_soum = SoumissionToolStripTextBox.Text & "%"
            Me.TasksTableAdapter.Fillbycontratsoum(Me.NorbecDataDataSet.Tasks, stringed_order, stringed_soum)
            Me.Tasks_Jobs_approTableAdapter.FillBycontratsoum(Me.NorbecDataDataSet.Tasks_Jobs_appro, stringed_order, stringed_soum)
            Me.Tasks_Jobs_prodTableAdapter.FillBycontratsoum(Me.NorbecDataDataSet.Tasks_Jobs_prod, stringed_order, stringed_soum)
            Me.Tasks_Jobs_montTableAdapter.FillBycontratsoum(Me.NorbecDataDataSet.Tasks_Jobs_mont, stringed_order, stringed_soum)
        Catch ex As System.Exception
            System.Windows.Forms.MessageBox.Show(ex.Message)
        End Try

    End Sub

    Private Sub Attrib_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'TODO: This line of code loads data into the 'NorbecDataDataSet.Mont' table. You can move, or remove it, as needed.
        Me.MontTableAdapter.Fill(Me.NorbecDataDataSet.Mont)
        'TODO: This line of code loads data into the 'NorbecDataDataSet.Suivi_Changement_Date' table. You can move, or remove it, as needed.
        'Me.Suivi_Changement_DateTableAdapter.Fill(Me.NorbecDataDataSet.Suivi_Changement_Date)
        'TODO: This line of code loads data into the 'NorbecDataDataSet.Employé' table. You can move, or remove it, as needed.

        'TODO: This line of code loads data into the 'NorbecDataDataSet1.Prod' table. You can move, or remove it, as needed.
        Me.ProdTableAdapter.Fill(Me.NorbecDataDataSet1.Prod)
        'TODO: This line of code loads data into the 'NorbecDataDataSet1.Appro' table. You can move, or remove it, as needed.
        Me.ApproTableAdapter.Fill(Me.NorbecDataDataSet1.Appro)
        'TODO: This line of code loads data into the 'NorbecDataDataSet.Tasks_Jobs_prod' table. You can move, or remove it, as needed.
        'Me.Tasks_Jobs_prodTableAdapter.Fill(Me.NorbecDataDataSet.Tasks_Jobs_prod)
        'TODO: This line of code loads data into the 'NorbecDataDataSet.Tasks_Jobs_appro' table. You can move, or remove it, as needed.
        'Me.Tasks_Jobs_approTableAdapter.Fill(Me.NorbecDataDataSet.Tasks_Jobs_appro)
        'TODO: This line of code loads data into the 'NorbecDataDataSet1.Tasks_Jobs_mont' table. You can move, or remove it, as needed.
        'Me.Tasks_Jobs_montTableAdapter.Fill(Me.NorbecDataDataSet1.Tasks_Jobs_mont)

        Me.NorbecDataDataSet.Tasks_Jobs_appro.date_addedColumn.DefaultValue = Date.Today.ToString
        Me.NorbecDataDataSet.Tasks_Jobs_appro.date_modifiedColumn.DefaultValue = Date.Today.ToString
        Me.NorbecDataDataSet1.Tasks_Jobs_prod.date_addedColumn.DefaultValue = Date.Today.ToString
        Me.NorbecDataDataSet1.Tasks_Jobs_prod.date_modifiedColumn.DefaultValue = Date.Today.ToString
        Me.NorbecDataDataSet.Tasks_Jobs_mont.date_addedColumn.DefaultValue = Date.Today.ToString
        Me.NorbecDataDataSet.Tasks_Jobs_mont.date_modifiedColumn.DefaultValue = Date.Today.ToString

        ' find user
        find_user()
        Me.EmployéTableAdapter.Fill(Me.NorbecDataDataSet.Employé)
    End Sub

    Sub find_user()
        Dim user_ta As New NorbecDataDataSetTableAdapters.EmployéTableAdapter
        Dim win_user = Environment.UserName
        user_ta.FillByUser(NorbecDataDataSet.Employé, win_user)
        ' default to user admin
        curr_user = 39
        If NorbecDataDataSet.Employé.Rows.Count > 0 Then
            curr_user = NorbecDataDataSet.Employé.Rows(0).Item("Employé_ID")
        End If
    End Sub

    Private Sub btn_des_prod_Click(sender As Object, e As EventArgs) Handles btn_des_prod.Click
        attrib_dessinateur(2)
    End Sub

    Private Sub btn_des_appro_Click(sender As Object, e As EventArgs) Handles btn_des_appro.Click
        attrib_dessinateur(1)
    End Sub

    Private Sub btn_des_mont_Click(sender As Object, e As EventArgs) Handles btn_des_mont.Click
        attrib_dessinateur(3)
    End Sub


    Sub attrib_dessinateur(type)
        Try

            Dim stringed_order, stringed_soum As String
            If Strings.Len(ContratToolStripTextBox.Text) > 4 Then stringed_order = ContratToolStripTextBox.Text & "%"
            If Strings.Len(SoumissionToolStripTextBox.Text) > 4 Then stringed_soum = SoumissionToolStripTextBox.Text & "%"
            Me.Check_tasks_jobsTableAdapter.Fill(Me.NorbecDataDataSet.check_tasks_jobs, type, stringed_order, stringed_soum)
            Dim sSQL As String = ""
            For Each found_tasks As DataRow In Me.NorbecDataDataSet.check_tasks_jobs.Rows
                If IsDBNull(found_tasks("revno")) Then
                    found_tasks("revno") = 0
                End If
                If combo_dessinateur.SelectedIndex = -1 Then
                    ' Delete
                    If Not IsDBNull(found_tasks("uid")) Then
                        sSQL = sSQL & "DELETE FROM [NorbecData].[dbo].Tasks_Jobs WHERE tasks_id=" & found_tasks("id") & " " & _
                                                " AND job_id=" & type & "; " & vbLf
                        write_histo(found_tasks("id"), "Delete du dessinateur ", type, "", found_tasks("revno"))
                    End If

                ElseIf IsDBNull(found_tasks("uid")) Then
                    ' insert
                    sSQL = sSQL & "INSERT INTO [NorbecData].[dbo].Tasks_Jobs ([tasks_id],[employe_id],[job_id],[date_added],revno,[date_modified]) VALUES " & _
                            "(" & found_tasks("id") & "," & _
                            combo_dessinateur.SelectedValue & "," & _
                            type & "," & _
                            "'" & Date.Today & "', " & _
                            "'" & found_tasks("revno") & "', " & _
                            "'" & Date.Today & "' " & _
                            "); " & vbLf

                    write_histo(found_tasks("id"), "Ajout dessinateur ", type, combo_dessinateur.Text, found_tasks("revno"))
                Else
                    'update
                    sSQL = sSQL & "UPDATE [NorbecData].[dbo].Tasks_Jobs SET [employe_id]=" & combo_dessinateur.SelectedValue & " ,revno=" & found_tasks("revno") & ",[date_modified]='" & Date.Today & "' " & _
                            "WHERE [tasks_id]=" & found_tasks("id") & " AND " & _
                            "[job_id]=" & type & _
                            "; " & vbLf
                    write_histo(found_tasks("id"), "Changement dessinateur ", type, combo_dessinateur.Text, found_tasks("revno"))
                End If

            Next
            'Debug.Print(sSQL)
            If sSQL <> "" Then
                ConfConn.Open()
                Dim SQL_doit As New SqlCommand(sSQL, ConfConn)
                SQL_doit.ExecuteScalar()
                ConfConn.Close()
            End If

            refresh_all()

            'Me.ProdTableAdapter.Fill(Me.NorbecDataDataSet1.Prod)
            'Me.ApproTableAdapter.Fill(Me.NorbecDataDataSet1.Appro)
            'Me.Tasks_Jobs_prodTableAdapter.Fill(Me.NorbecDataDataSet.Tasks_Jobs_prod)
            'Me.Tasks_Jobs_approTableAdapter.Fill(Me.NorbecDataDataSet.Tasks_Jobs_appro)
            'Me.TasksTableAdapter.Fillbycontratsoum(Me.NorbecDataDataSet.Tasks, stringed_order, stringed_soum)


        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try

    End Sub

    Sub refresh_all()
        Dim stringed_order, stringed_soum As String
        If Strings.Len(ContratToolStripTextBox.Text) > 4 Then stringed_order = ContratToolStripTextBox.Text & "%"
        If Strings.Len(SoumissionToolStripTextBox.Text) > 4 Then stringed_soum = SoumissionToolStripTextBox.Text & "%"

        Me.ProdTableAdapter.Fill(Me.NorbecDataDataSet1.Prod)
        Me.ApproTableAdapter.Fill(Me.NorbecDataDataSet1.Appro)
        Me.MontTableAdapter.Fill(Me.NorbecDataDataSet1.Mont)
        'Me.Tasks_Jobs_prodTableAdapter.Fill(Me.NorbecDataDataSet.Tasks_Jobs_prod)
        Me.Tasks_Jobs_prodTableAdapter.FillBycontratsoum(Me.NorbecDataDataSet.Tasks_Jobs_prod, stringed_order, stringed_soum)
        'Me.Tasks_Jobs_approTableAdapter.Fill(Me.NorbecDataDataSet.Tasks_Jobs_appro)
        Me.Tasks_Jobs_approTableAdapter.FillBycontratsoum(Me.NorbecDataDataSet.Tasks_Jobs_appro, stringed_order, stringed_soum)
        'Me.Tasks_Jobs_montTableAdapter.Fill(Me.NorbecDataDataSet.Tasks_Jobs_mont)
        Me.Tasks_Jobs_montTableAdapter.FillBycontratsoum(Me.NorbecDataDataSet.Tasks_Jobs_mont, stringed_order, stringed_soum)
        Me.TasksTableAdapter.Fillbycontratsoum(Me.NorbecDataDataSet.Tasks, stringed_order, stringed_soum)
    End Sub

    Sub write_histo(tasks_id, type_changement, type_dessin, dessinateur, revno)
        Try
            ' Ajoute le [NorbecData].[dbo].[Suivi Changement Date]
            If type_dessin = 1 Then type_changement = type_changement & "appro"
            If type_dessin = 2 Then type_changement = type_changement & "prod"
            If type_dessin = 3 Then type_changement = type_changement & "mont"
            type_changement = type_changement & " - Rev" & revno
            

            Dim new_cd As NorbecDataDataSet.Suivi_Changement_DateRow = NorbecDataDataSet.Tables("Suivi Changement Date").NewRow
            new_cd.TypeID = "TASKS"
            new_cd.ID = tasks_id
            new_cd.Date_du_changement = Date.Now
            new_cd.Type_de_changement = type_changement
            new_cd.Valeur_du_changement = dessinateur
            new_cd.Nom_Formulaire = "Dessin.net"
            new_cd.UserID = curr_user
            new_cd.ChgtDateComment = "MAJ par outils Dessin"

            Me.NorbecDataDataSet.Tables("Suivi Changement Date").Rows.Add(new_cd)
            Me.Suivi_Changement_DateTableAdapter.Update(new_cd)
        Catch ex As Exception

        End Try
    End Sub

    Sub set_dessin_app_fait()
        Dim sSQL
        Dim check_racine = 0
        For Each Task As NorbecDataDataSet.TasksRow In NorbecDataDataSet.Tasks.Rows
            If check_racine <> Task.Racine_ID Then
                sSQL = sSQL & "UPDATE [NorbecData].[dbo].[DessinAppRevision] SET [DessinAppFait]=1 " & _
                        "WHERE [RacineID]=" & Task.Racine_ID & _
                        "; " & vbLf
                write_histo_cocher(Task.Racine_ID, Task.Apprevno)
            End If
            check_racine = Task.Racine_ID
            'write_histo(found_tasks("id"), "Changement dessinateur ", Type, combo_dessinateur.Text, found_tasks("revno"))

            Debug.Print(sSQL)
        Next
        If sSQL <> "" Then
            ConfConn.Open()
            Dim SQL_doit As New SqlCommand(sSQL, ConfConn)
            SQL_doit.ExecuteScalar()
            ConfConn.Close()

        End If
        refresh_all()
    End Sub

    Sub write_histo_cocher(racine_id, revno)
        Try
            ' Ajoute le [NorbecData].[dbo].[Suivi Changement Date]
            Dim type_changement = "Case Dessin en approbation" & " - Rev" & revno

            Dim new_cd As NorbecDataDataSet.Suivi_Changement_DateRow = NorbecDataDataSet.Tables("Suivi Changement Date").NewRow
            new_cd.TypeID = "RACINE"
            new_cd.ID = racine_id
            new_cd.Date_du_changement = Date.Now
            new_cd.Type_de_changement = type_changement
            new_cd.Valeur_du_changement = "Cocher"
            new_cd.Nom_Formulaire = "Dessin.net"
            new_cd.UserID = curr_user
            new_cd.ChgtDateComment = "MAJ par outils Dessin"

            Me.NorbecDataDataSet.Tables("Suivi Changement Date").Rows.Add(new_cd)
            Me.Suivi_Changement_DateTableAdapter.Update(new_cd)
        Catch ex As Exception

        End Try
    End Sub

    Sub set_dessin_final_fait()
        Try
            Dim sSQL
            Dim check_task = 0
            For Each Task As NorbecDataDataSet.TasksRow In NorbecDataDataSet.Tasks.Rows
                If check_task <> Task.id And Task.Dessin_Final_fait <> 1 Then
                    sSQL = sSQL & "UPDATE [NorbecData].[dbo].[Tasks] SET [Dessin Final fait]=1 " & _
                            "WHERE [ID]=" & Task.id & _
                            "; " & vbLf
                    write_histo_final(Task.id)
                End If
                check_task = Task.id
                'Debug.Print(sSQL)
            Next
            If sSQL <> "" Then
                ConfConn.Open()
                Dim SQL_doit As New SqlCommand(sSQL, ConfConn)
                SQL_doit.ExecuteScalar()
                ConfConn.Close()

            End If
            refresh_all()

        Catch ex As Exception
            MsgBox("Problème a l'écriture de dessin final fait!" & vbLf & ex.ToString)
        End Try
    End Sub

    Sub set_verif_fait()
        Try
            Dim sSQL
            Dim check_task = 0
            For Each Task As NorbecDataDataSet.TasksRow In NorbecDataDataSet.Tasks.Rows
                If check_task <> Task.id And Task.Dessin_Final_fait <> 1 Then
                    sSQL = sSQL & "UPDATE [NorbecData].[dbo].[Tasks] SET [verif_conception]=1 " &
                            "WHERE [ID]=" & Task.id &
                            "; " & vbLf
                    'write_histo_final(Task.id)
                End If
                check_task = Task.id
                'Debug.Print(sSQL)
            Next
            If sSQL <> "" Then
                ConfConn.Open()
                Dim SQL_doit As New SqlCommand(sSQL, ConfConn)
                SQL_doit.ExecuteScalar()
                ConfConn.Close()

            End If
            refresh_all()

        Catch ex As Exception
            MsgBox("Problème a l'écriture de la verif!" & vbLf & ex.ToString)
        End Try
    End Sub

    Sub write_histo_final(task_id)
        Try
            ' Ajoute le [NorbecData].[dbo].[Suivi Changement Date]
            Dim type_changement = "Case Dessin final fait"

            Dim new_cd As NorbecDataDataSet.Suivi_Changement_DateRow = NorbecDataDataSet.Tables("Suivi Changement Date").NewRow
            new_cd.TypeID = "TASKS"
            new_cd.ID = task_id
            new_cd.Date_du_changement = Date.Now
            new_cd.Type_de_changement = type_changement
            new_cd.Valeur_du_changement = "Cocher"
            new_cd.Nom_Formulaire = "Dessin.net"
            new_cd.UserID = curr_user
            new_cd.ChgtDateComment = "MAJ par outils Dessin"

            Me.NorbecDataDataSet.Tables("Suivi Changement Date").Rows.Add(new_cd)
            Me.Suivi_Changement_DateTableAdapter.Update(new_cd)
        Catch ex As Exception

        End Try
    End Sub

    Private Sub btn_save_Click(sender As Object, e As EventArgs) Handles btn_save.Click
        'Me.Tasks_Jobs_prodTableAdapter.Fill(Me.NorbecDataDataSet.Tasks_Jobs_prod)
        'Me.Tasks_Jobs_approTableAdapter.Fill(Me.NorbecDataDataSet.Tasks_Jobs_appro)
        'Me.Tasks_Jobs_montTableAdapter.Fill(Me.NorbecDataDataSet1.Tasks_Jobs_mont)
        Dim change_data_tjp As NorbecDataDataSet.Tasks_Jobs_prodDataTable = Me.NorbecDataDataSet.Tasks_Jobs_prod.GetChanges()
        If Not IsNothing(change_data_tjp) Then
            If change_data_tjp.Rows(0)("employe_id", DataRowVersion.Original).ToString <> change_data_tjp.Rows(0)("employe_id", DataRowVersion.Current).ToString Then
                ' write histo
                Dim prod_empl As NorbecDataDataSet.ProdRow() = NorbecDataDataSet1.Prod.Select("Employé_ID=" & change_data_tjp.Rows(0)("employe_id"))
                If prod_empl.Count > 0 Then
                    write_histo(change_data_tjp.Rows(0)("tasks_id"), "Changement dessinateur ", 2, prod_empl(0).NomComplet, change_data_tjp.Rows(0)("revno").ToString)
                End If
            End If
        End If

        Dim change_data_tja As NorbecDataDataSet.Tasks_Jobs_approDataTable = Me.NorbecDataDataSet.Tasks_Jobs_appro.GetChanges
        If Not IsNothing(change_data_tja) Then
            If change_data_tja.Rows(0)("employe_id", DataRowVersion.Original).ToString <> change_data_tja.Rows(0)("employe_id", DataRowVersion.Current).ToString Then
                ' write histo
                Dim appro_empl As NorbecDataDataSet.ProdRow() = NorbecDataDataSet1.Prod.Select("Employé_ID=" & change_data_tja.Rows(0)("employe_id"))
                If appro_empl.Count > 0 Then
                    write_histo(change_data_tja.Rows(0)("tasks_id"), "Changement dessinateur ", 1, appro_empl(0).NomComplet, change_data_tja.Rows(0)("revno").ToString)
                End If
            End If
        End If

        Me.Validate()
        Me.TasksBindingSource.EndEdit()
        Me.Tasks_Jobs_approBindingSource.EndEdit()
        Me.Tasks_Jobs_prodBindingSource.EndEdit()
        Me.Tasks_Jobs_montBindingSource.EndEdit()
        Me.TableAdapterManager.UpdateAll(Me.NorbecDataDataSet)
        Me.TableAdapterManager.UpdateAll(Me.NorbecDataDataSet1)
    End Sub

    Private Sub btn_dess_app_fait_Click(sender As Object, e As EventArgs) Handles btn_dess_app_fait.Click
        set_dessin_app_fait()
    End Sub

    Private Sub btn_verif_fait_Click(sender As Object, e As EventArgs) Handles btn_verif_fait.Click
        set_verif_fait()
    End Sub

    Private Sub btn_dessin_fin_fait_Click(sender As Object, e As EventArgs) Handles btn_dessin_fin_fait.Click
        If MsgBox("Attention cette fonction est en test. Voulez-vous continuer?", MsgBoxStyle.OkCancel, "EN TEST") = MsgBoxResult.Ok Then
            set_dessin_final_fait()
        End If

    End Sub


    Private Sub TasksDataGridView_DataError(sender As Object, e As DataGridViewDataErrorEventArgs) Handles TasksDataGridView.DataError

    End Sub


    Private Sub Tasks_Jobs_montDataGridView_DataError(sender As Object, e As DataGridViewDataErrorEventArgs) Handles Tasks_Jobs_montDataGridView.DataError

    End Sub


    Private Sub Tasks_Jobs_prodDataGridView_DataError(sender As Object, e As DataGridViewDataErrorEventArgs) Handles Tasks_Jobs_prodDataGridView.DataError

    End Sub


    Private Sub Tasks_Jobs_approDataGridView_DataError(sender As Object, e As DataGridViewDataErrorEventArgs) Handles Tasks_Jobs_approDataGridView.DataError

    End Sub


End Class
