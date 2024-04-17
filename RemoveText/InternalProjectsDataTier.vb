'Title:         Internal Projects Data Tier
'Data:          12/30/13
'Author:        Terry Holmes

'Description:   This is the Data Tier for the Internal Projects Form

Option Strict On

Public Class InternalProjectsDataTier

    'Setting Up Internal Projects Modular Variables
    Private aInternalProjectsTableAdapter As InternalProjectsDataSetTableAdapters.internalprojectsTableAdapter
    Private aInternalProjectsDataSet As InternalProjectsDataSet

    Public Function GetInternalProjectsInformation() As InternalProjectsDataSet

        'Setting up this Function of the Data Tier
        Try

            aInternalProjectsDataSet = New InternalProjectsDataSet
            aInternalProjectsTableAdapter = New InternalProjectsDataSetTableAdapters.internalprojectsTableAdapter
            aInternalProjectsTableAdapter.Fill(aInternalProjectsDataSet.internalprojects)
            Return aInternalProjectsDataSet

        Catch ex As Exception

            MessageBox.Show(ex.Message, "Please Check", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return aInternalProjectsDataSet

        End Try
    End Function

    Public Sub UpdateInternalProjectsDB(ByVal aInternalProjectsDataSet As InternalProjectsDataSet)

        Try

            aInternalProjectsTableAdapter.Update(aInternalProjectsDataSet.internalprojects)

        Catch ex As Exception
            MessageBox.Show(ex.Message, "Please Correct", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

    End Sub

End Class
