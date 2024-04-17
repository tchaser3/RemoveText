'Title:         Create Last Transaction ID
'Date:          2-11-15
'Author:        Terry Holmes

'Description:   This will create a last transaction id

Option Strict On

Public Class CreateLastTransactionID

    'Setting global variables
    Private TheLastTransactionIDDataSet As LastTransactionIDDataSet
    Private TheLastTransactionIDDataTier As LastTransactionDataTier
    Private WithEvents TheLastTransactionIDBindingSource As BindingSource

    Private editingBoolean As Boolean = False
    Private addingBoolean As Boolean = False
    Private previousSelectedIndex As Integer

    Private Sub CreateLastTransactionID_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        'Setting local variables
        Dim intNewID As Integer

        'Try Catch to load controls
        Try

            TheLastTransactionIDDataTier = New LastTransactionDataTier
            TheLastTransactionIDDataSet = TheLastTransactionIDDataTier.GetLastTransactionIDInformation
            TheLastTransactionIDBindingSource = New BindingSource

            'Setting up the binding source
            With TheLastTransactionIDBindingSource
                .DataSource = TheLastTransactionIDDataSet
                .DataMember = "lasttransactionid"
                .MoveFirst()
                .MoveLast()
            End With

            'Setting up the combo box
            With cboTransactionID
                .DataSource = TheLastTransactionIDBindingSource
                .DisplayMember = "TransactionID"
                .DataBindings.Add("text", TheLastTransactionIDBindingSource, "TransactionID", False, DataSourceUpdateMode.Never)
            End With

            'Binding the last control
            txtCreatedTransactionID.DataBindings.Add("text", TheLastTransactionIDBindingSource, "CreatedTransactionID")

            'Beginning process to add a record
            intNewID = CInt(txtCreatedTransactionID.Text)
            intNewID += 1
            txtCreatedTransactionID.Text = CStr(intNewID)
            Logon.mintCreatedTransactionID = CInt(intNewID)

            'Saving the record
            TheLastTransactionIDBindingSource.EndEdit()
            TheLastTransactionIDDataTier.UpdateLastTransactionIDDB(TheLastTransactionIDDataSet)
            editingBoolean = False
            addingBoolean = False
            SetComboBoxBinding()

            Me.Close()

        Catch ex As Exception

            'Message to user if there is a problem
            MessageBox.Show(ex.Message, "Please Correct", MessageBoxButtons.OK, MessageBoxIcon.Error)

        End Try

    End Sub

    Private Sub SetComboBoxBinding()

        With cboTransactionID
            If (addingBoolean Or editingBoolean) Then
                .DataBindings!text.DataSourceUpdateMode = DataSourceUpdateMode.OnValidation
                .DropDownStyle = ComboBoxStyle.Simple
            Else
                .DataBindings!text.DataSourceUpdateMode = DataSourceUpdateMode.Never
                .DropDownStyle = ComboBoxStyle.DropDownList
            End If
        End With

    End Sub


End Class