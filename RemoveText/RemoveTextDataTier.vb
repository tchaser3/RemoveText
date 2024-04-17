'Title:         Remove Text Data Tier
'Date:          1-22-15
'Author:        Terry Holmes

'Description:   This the data tier for the application

Option Strict On

Public Class RemoveTextDataTier

    'Setting up the global variables
    Private aIssuedPartsDataSetTableAdapter As IssuedPartsDataSetTableAdapters.IssuedPartsTableAdapter
    Private aIssuedPartsDataSet As IssuedPartsDataSet

    Private aUsedPartsDataSetTableAdapter As UsedPartsDataSetTableAdapters.BOMPartsTableAdapter
    Private aUsedPartsDataSet As UsedPartsDataSet

    Private aReceivePartsDataSetTableAdapter As ReceivePartsDataSetTableAdapters.ReceivedPartsTableAdapter
    Private aReceivePartsDataSet As ReceivePartsDataSet

    Public Function GetReceivePartsInformation() As ReceivePartsDataSet
        'Setting up the Datatier
        Try
            aReceivePartsDataSet = New ReceivePartsDataSet
            aReceivePartsDataSetTableAdapter = New ReceivePartsDataSetTableAdapters.ReceivedPartsTableAdapter
            aReceivePartsDataSetTableAdapter.Fill(aReceivePartsDataSet.ReceivedParts)
            Return aReceivePartsDataSet

        Catch ex As Exception
            MessageBox.Show(ex.Message, "Please Check", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return aReceivePartsDataSet
        End Try
    End Function
    Public Sub UpdateReceivePartsDB(ByVal aReceivePartsDataSet As ReceivePartsDataSet)

        'This will update the database
        Try
            aReceivePartsDataSetTableAdapter.Update(aReceivePartsDataSet.ReceivedParts)
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Please Correct", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Public Function GetUsedPartsInformation() As UsedPartsDataSet
        'Setting up the Datatier
        Try
            aUsedPartsDataSet = New UsedPartsDataSet
            aUsedPartsDataSetTableAdapter = New UsedPartsDataSetTableAdapters.BOMPartsTableAdapter
            aUsedPartsDataSetTableAdapter.Fill(aUsedPartsDataSet.BOMParts)
            Return aUsedPartsDataSet

        Catch ex As Exception
            MessageBox.Show(ex.Message, "Please Check", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return aUsedPartsDataSet
        End Try
    End Function
    Public Sub UpdateUsedPartsDB(ByVal aUsedPartsDataSet As UsedPartsDataSet)

        'This will update the database
        Try
            aUsedPartsDataSetTableAdapter.Update(aUsedPartsDataSet.BOMParts)
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Please Correct", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Public Function GetIssuedPartsInformation() As IssuedPartsDataSet
        'Setting up the Datatier
        Try
            aIssuedPartsDataSet = New IssuedPartsDataSet
            aIssuedPartsDataSetTableAdapter = New IssuedPartsDataSetTableAdapters.IssuedPartsTableAdapter
            aIssuedPartsDataSetTableAdapter.Fill(aIssuedPartsDataSet.IssuedParts)
            Return aIssuedPartsDataSet

        Catch ex As Exception
            MessageBox.Show(ex.Message, "Please Check", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return aIssuedPartsDataSet
        End Try
    End Function
    Public Sub UpdateIssuedPartsDB(ByVal aIssuedPartsDataSet As IssuedPartsDataSet)

        'This will update the database
        Try
            aIssuedPartsDataSetTableAdapter.Update(aIssuedPartsDataSet.IssuedParts)
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Please Correct", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

End Class
