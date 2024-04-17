'Title:         Remove PD and DID Text
'Date:          1-22-15
'Author:        Terry Holmes

'Description:   This will remove the text for DID and PD

Option Strict On

Public Class Form1

    'Setting up the global variables
    Private TheIssuedPartsDataSet As IssuedPartsDataSet
    Private TheIssuedPartsDataTier As RemoveTextDataTier
    Private WithEvents TheIssuedPartsBindingSource As BindingSource

    Private TheUsedPartsDataSet As UsedPartsDataSet
    Private TheUsedPartsDataTier As RemoveTextDataTier
    Private WithEvents TheUsedPartsBindingSource As BindingSource

    Private TheReceivePartsDataSet As ReceivePartsDataSet
    Private TheReceivePartsDataTier As RemoveTextDataTier
    Private WithEvents TheReceivePartsBindingSource As BindingSource

    Private TheInternalProjectDataSet As InternalProjectsDataSet
    Private TheInternalProjectDataTier As InternalProjectsDataTier
    Private WithEvents TheInternalProjectBindingSource As BindingSource

    Private Sub btnClose_Click(sender As Object, e As EventArgs) Handles btnClose.Click

        'This will call the Close Program Form
        CloseProgram.ShowDialog()

    End Sub
    Private Sub ClearDataBindings()

        'This will clear the data bindings
        cboTransactionID.DataBindings.Clear()
        txtProjectID.DataBindings.Clear()
        txtInternalID.DataBindings.Clear()

    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles Me.Load

        'This will clear the data bindings
        ClearDataBindings()

    End Sub
    Private Sub SetProjectDataBindings()

        'Try Catch to catch exceptions
        Try

            TheInternalProjectDataTier = New InternalProjectsDataTier
            TheInternalProjectDataSet = TheInternalProjectDataTier.GetInternalProjectsInformation
            TheInternalProjectBindingSource = New BindingSource

            'Setting up the binding source
            With TheInternalProjectBindingSource
                .DataSource = TheInternalProjectDataSet
                .DataMember = "internalprojects"
                .MoveFirst()
                .MoveLast()
            End With

            'Setting up the combo box
            With cboTransactionID
                .DataSource = TheInternalProjectBindingSource
                .DisplayMember = "InternalProjectID"
                .DataBindings.Add("text", TheInternalProjectBindingSource, "InternalProjectID", False, DataSourceUpdateMode.Never)
            End With

            'Binding the other control
            txtProjectID.DataBindings.Add("text", TheInternalProjectBindingSource, "TWCControlNumber")

        Catch ex As Exception
            MessageBox.Show(ex.Message, "Please Correct", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

    End Sub
    Private Sub SetReceivePartsDataBindings()

        'Try Catch to catch exceptions
        Try

            TheReceivePartsDataTier = New RemoveTextDataTier
            TheReceivePartsDataSet = TheReceivePartsDataTier.GetReceivePartsInformation
            TheReceivePartsBindingSource = New BindingSource

            'Setting up the binding source
            With TheReceivePartsBindingSource
                .DataSource = TheReceivePartsDataSet
                .DataMember = "ReceivedParts"
                .MoveFirst()
                .MoveLast()
            End With

            'Setting up the combo box
            With cboTransactionID
                .DataSource = TheReceivePartsBindingSource
                .DisplayMember = "TransactionID"
                .DataBindings.Add("text", TheReceivePartsBindingSource, "TransactionID", False, DataSourceUpdateMode.Never)
            End With

            'Binding the other control
            txtProjectID.DataBindings.Add("text", TheReceivePartsBindingSource, "ProjectID")
            txtInternalID.DataBindings.Add("text", TheReceivePartsBindingSource, "InternalProjectID")

        Catch ex As Exception
            MessageBox.Show(ex.Message, "Please Correct", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

    End Sub
    Private Sub SetUsedPartsDataBindings()

        'Try Catch to catch exceptions
        Try

            TheUsedPartsDataTier = New RemoveTextDataTier
            TheUsedPartsDataSet = TheUsedPartsDataTier.GetUsedPartsInformation
            TheUsedPartsBindingSource = New BindingSource

            'Setting up the binding source
            With TheUsedPartsBindingSource
                .DataSource = TheUsedPartsDataSet
                .DataMember = "BOMParts"
                .MoveFirst()
                .MoveLast()
            End With

            'Setting up the combo box
            With cboTransactionID
                .DataSource = TheUsedPartsBindingSource
                .DisplayMember = "TransactionID"
                .DataBindings.Add("text", TheUsedPartsBindingSource, "TransactionID", False, DataSourceUpdateMode.Never)
            End With

            'Binding the other control
            txtProjectID.DataBindings.Add("text", TheUsedPartsBindingSource, "ProjectID")
            txtInternalID.DataBindings.Add("text", TheUsedPartsBindingSource, "InternalProjectID")

        Catch ex As Exception
            MessageBox.Show(ex.Message, "Please Correct", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

    End Sub
    Private Sub SetIssuedPartsDataBindings()

        'Try Catch to catch exceptions
        Try

            TheIssuedPartsDataTier = New RemoveTextDataTier
            TheIssuedPartsDataSet = TheIssuedPartsDataTier.GetIssuedPartsInformation
            TheIssuedPartsBindingSource = New BindingSource

            'Setting up the binding source
            With TheIssuedPartsBindingSource
                .DataSource = TheIssuedPartsDataSet
                .DataMember = "IssuedParts"
                .MoveFirst()
                .MoveLast()
            End With

            'Setting up the combo box
            With cboTransactionID
                .DataSource = TheIssuedPartsBindingSource
                .DisplayMember = "TransactionID"
                .DataBindings.Add("text", TheIssuedPartsBindingSource, "TransactionID", False, DataSourceUpdateMode.Never)
            End With

            'Binding the other control
            txtProjectID.DataBindings.Add("text", TheIssuedPartsBindingSource, "ProjectID")
            txtInternalID.DataBindings.Add("text", TheIssuedPartsBindingSource, "InternalProjectID")

        Catch ex As Exception
            MessageBox.Show(ex.Message, "Please Correct", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

    End Sub
    Private Sub UpdateRecords(ByVal strTableType As String)

        'Setting local variables
        Dim intCounter As Integer
        Dim intNumberOfRecords As Integer
        Dim strProjectIDFromTable As String = ""
        Dim intStringLength As Integer
        Dim strPDString As String = ""
        Dim strDIDString As String = ""
        Dim chaKeyWordArray() As Char
        Dim intPDFirstCounter As Integer
        Dim intPDSecondcounter As Integer
        Dim intPDFirstLimit As Integer
        Dim intDIDFirstCounter As Integer
        Dim intDIDSecondCounter As Integer
        Dim intDIDFirstLimit As Integer

        'Getting record count
        intNumberOfRecords = cboTransactionID.Items.Count - 1

        'Performing Loop
        For intCounter = 0 To intNumberOfRecords

            'Incrementing the counter
            cboTransactionID.SelectedIndex = intCounter

            'Loading up the array
            chaKeyWordArray = txtProjectID.Text.ToCharArray

            intStringLength = txtProjectID.Text.Length

            'Loading the PD String
            If intStringLength > 1 Then
                strPDString = CStr(chaKeyWordArray(0)) + CStr(chaKeyWordArray(1))
                intPDFirstLimit = intStringLength - 1
            End If

            If intStringLength > 2 Then
                'Loading the DID String
                strDIDString = CStr(chaKeyWordArray(0)) + CStr(chaKeyWordArray(1)) + CStr(chaKeyWordArray(2))
                intDIDFirstLimit = intStringLength - 1
            End If

            'If statement to see if works
            If strPDString = "PD" Then

                'Setting the PD Counter
                intPDFirstCounter = 2

                'Checking to see if there is a 
                If chaKeyWordArray(2) = " " Then

                    'Adding a number if there is a space
                    intPDFirstCounter += 1

                End If

                'Creating ID
                For intPDSecondcounter = intPDFirstCounter To intPDFirstLimit

                    'Creating the ID
                    strProjectIDFromTable = strProjectIDFromTable + chaKeyWordArray(intPDSecondcounter)

                Next

                'loading the text
                txtProjectID.Text = strProjectIDFromTable

                'updating the record
                SaveRecords(strTableType)

                'clearing the id
                strProjectIDFromTable = ""

            End If

            'If statement to see if works
            If strPDString = "DID" Then

                'Setting the PD Counter
                intDIDFirstCounter = 3

                'Checking to see if there is a 
                If chaKeyWordArray(3) = " " Then

                    'Adding a number if there is a space
                    intDIDFirstCounter += 1

                End If

                'Creating ID
                For intDIDSecondCounter = intDIDFirstCounter To intDIDFirstLimit

                    'Creating the ID
                    strProjectIDFromTable = strProjectIDFromTable + chaKeyWordArray(intDIDSecondCounter)

                Next

                'loading the text
                txtProjectID.Text = strProjectIDFromTable

                'updating the record
                SaveRecords(strTableType)

                'clearing the id
                strProjectIDFromTable = ""

            End If

            If txtProjectID.Text = "TECH OPS" Then
                txtInternalID.Text = "1055"
            End If

            strPDString = ""
            strDIDString = ""

        Next

    End Sub

    Private Sub btnRun_Click(sender As Object, e As EventArgs) Handles btnRun.Click

        ClearDataBindings()
        SetIssuedPartsDataBindings()
        UpdateRecords("ISSUED")
        ClearDataBindings()
        SetReceivePartsDataBindings()
        UpdateRecords("RECEIVE")
        ClearDataBindings()
        SetUsedPartsDataBindings()
        UpdateRecords("USED")
        ClearDataBindings()
        SetProjectDataBindings()
        UpdateRecords("PROJECT")
        Logon.mstrLastTransactionSummary = "COMPLETED REMOVING PD AND DID TEXT"
        LastTransaction.Show()

    End Sub
    Private Sub SaveRecords(ByVal strActiveBindingSource As String)

        'This sub routine will update the data sets
        If strActiveBindingSource = "ISSUED" Then
            TheIssuedPartsBindingSource.EndEdit()
            TheIssuedPartsDataTier.UpdateIssuedPartsDB(TheIssuedPartsDataSet)
        ElseIf strActiveBindingSource = "USED" Then
            TheUsedPartsBindingSource.EndEdit()
            TheUsedPartsDataTier.UpdateUsedPartsDB(TheUsedPartsDataSet)
        ElseIf strActiveBindingSource = "RECEIVE" Then
            TheReceivePartsBindingSource.EndEdit()
            TheReceivePartsDataTier.UpdateReceivePartsDB(TheReceivePartsDataSet)
        ElseIf strActiveBindingSource = "PROJECT" Then
            TheInternalProjectBindingSource.EndEdit()
            TheInternalProjectDataTier.UpdateInternalProjectsDB(TheInternalProjectDataSet)
        End If

    End Sub
End Class
