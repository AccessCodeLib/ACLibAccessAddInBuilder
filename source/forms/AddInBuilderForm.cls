Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private m_Configuration As AddInConfiguration

Private Property Get Configuration() As AddInConfiguration
   If m_Configuration Is Nothing Then
      Set m_Configuration = New AddInConfiguration
   End If
   Set Configuration = m_Configuration
End Property

Private Sub cmdLoadFromAddIn_Click()
   LoadDataFromConfiguration True
End Sub

Private Sub cmdRefreshFile_Click()
   Me.txtFileName.Value = Replace(CurrentProject.Name, ".accdb", ".accda", , , vbTextCompare)
End Sub

Private Sub cmdSaveAddIn_Click()
   SaveToCurrentFile
End Sub

Private Sub Form_Load()

   With CurrentApplication
      Me.lblVersionInfo.Caption = .ApplicationFullName & " " & ChrW(&H2022) & " Version " & CurrentApplication.Version
   End With

   If CurrentDb Is Nothing Then
      Me.cmdLoadFromAddIn.Enabled = False
      Me.cmdSaveAddIn.Caption = "Create Add-in"
      Me.txtFileName.BorderStyle = 1
   Else
      LoadDataFromConfiguration True
   End If

End Sub

Private Sub LoadDataFromConfiguration(ByVal LoadFromFile As Boolean)

   If LoadFromFile Then
      Configuration.LoadFromCurrentFile
   End If

   With Configuration

      Me.txtFileName.Value = .FileName
      Me.txtAppTitle.Value = .AppTitle

      Me.txtAddInTitle.Value = .Title
      Me.txtAddInAuthor.Value = .Author
      Me.txtAddInCompany.Value = .Company
      Me.txtAddInComment.Value = .Comments

      Me.txtAddInRegPathName.Value = .AddInRegPathName
      Me.txtAddInStartFunction.Value = .AddInStartFunction

   End With

End Sub

Private Sub SaveToCurrentFile()

   SaveDataToConfiguration
   Configuration.SaveToCurrentFile

End Sub

Private Sub SaveDataToConfiguration()

   With Configuration

      .FileName = Nz(Me.txtFileName.Value, vbNullString)
      .AppTitle = Nz(Me.txtAppTitle.Value, vbNullString)

      .Title = Nz(Me.txtAddInTitle.Value, vbNullString)
      .Author = Nz(Me.txtAddInAuthor.Value, vbNullString)
      .Company = Nz(Me.txtAddInCompany.Value, vbNullString)
      .Comments = Nz(Me.txtAddInComment.Value, vbNullString)

      .AddInRegPathName = Nz(Me.txtAddInRegPathName.Value, vbNullString)
      .AddInStartFunction = Nz(Me.txtAddInStartFunction.Value, vbNullString)

   End With

End Sub
