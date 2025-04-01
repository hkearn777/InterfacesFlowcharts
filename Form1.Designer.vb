<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class Form1
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()>
    Protected Overrides Sub Dispose(disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

  'NOTE: The following procedure is required by the Windows Form Designer
  'It can be modified using the Windows Form Designer.  
  'Do not modify it using the code editor.
  <System.Diagnostics.DebuggerStepThrough()>
  Private Sub InitializeComponent()
    txtInterfaceFilename = New TextBox()
    txtPumlFolderName = New TextBox()
    btnGo = New Button()
    btnClose = New Button()
    txtInternalPumlFilename = New TextBox()
    txtExternalPumlFilename = New TextBox()
    Label1 = New Label()
    Label2 = New Label()
    Label3 = New Label()
    Label4 = New Label()
    Label5 = New Label()
    txtTargetApp = New TextBox()
    SuspendLayout()
    ' 
    ' txtInterfaceFilename
    ' 
    txtInterfaceFilename.Location = New Point(30, 70)
    txtInterfaceFilename.Name = "txtInterfaceFilename"
    txtInterfaceFilename.Size = New Size(1081, 31)
    txtInterfaceFilename.TabIndex = 0
    txtInterfaceFilename.Text = "C:\Users\906074897\Documents\All Projects\BCBS\Interfaces Internal and external to RTMS 2025-04-01.xlsx"
    ' 
    ' txtPumlFolderName
    ' 
    txtPumlFolderName.Location = New Point(30, 146)
    txtPumlFolderName.Name = "txtPumlFolderName"
    txtPumlFolderName.Size = New Size(1081, 31)
    txtPumlFolderName.TabIndex = 1
    txtPumlFolderName.Text = "C:\Users\906074897\Documents\All Projects\BCBS"
    ' 
    ' btnGo
    ' 
    btnGo.Location = New Point(869, 354)
    btnGo.Name = "btnGo"
    btnGo.Size = New Size(112, 34)
    btnGo.TabIndex = 2
    btnGo.Text = "Go"
    btnGo.UseVisualStyleBackColor = True
    ' 
    ' btnClose
    ' 
    btnClose.Location = New Point(999, 354)
    btnClose.Name = "btnClose"
    btnClose.Size = New Size(112, 34)
    btnClose.TabIndex = 3
    btnClose.Text = "Close"
    btnClose.UseVisualStyleBackColor = True
    ' 
    ' txtInternalPumlFilename
    ' 
    txtInternalPumlFilename.Location = New Point(27, 239)
    txtInternalPumlFilename.Name = "txtInternalPumlFilename"
    txtInternalPumlFilename.Size = New Size(336, 31)
    txtInternalPumlFilename.TabIndex = 4
    txtInternalPumlFilename.Text = "\RTMS Internal Interfaces.puml"
    ' 
    ' txtExternalPumlFilename
    ' 
    txtExternalPumlFilename.Location = New Point(379, 239)
    txtExternalPumlFilename.Name = "txtExternalPumlFilename"
    txtExternalPumlFilename.Size = New Size(336, 31)
    txtExternalPumlFilename.TabIndex = 5
    txtExternalPumlFilename.Text = "\RTMS External Interfaces.puml"
    ' 
    ' Label1
    ' 
    Label1.AutoSize = True
    Label1.Location = New Point(30, 42)
    Label1.Name = "Label1"
    Label1.Size = New Size(114, 25)
    Label1.TabIndex = 6
    Label1.Text = "Spreadsheet:"
    ' 
    ' Label2
    ' 
    Label2.AutoSize = True
    Label2.Location = New Point(30, 118)
    Label2.Name = "Label2"
    Label2.Size = New Size(117, 25)
    Label2.TabIndex = 7
    Label2.Text = "PUML Folder:"
    ' 
    ' Label3
    ' 
    Label3.AutoSize = True
    Label3.Location = New Point(27, 211)
    Label3.Name = "Label3"
    Label3.Size = New Size(230, 25)
    Label3.TabIndex = 8
    Label3.Text = "PUML Internal Interface File:"
    ' 
    ' Label4
    ' 
    Label4.AutoSize = True
    Label4.Location = New Point(379, 211)
    Label4.Name = "Label4"
    Label4.Size = New Size(232, 25)
    Label4.TabIndex = 9
    Label4.Text = "PUML External Interface File:"
    ' 
    ' Label5
    ' 
    Label5.AutoSize = True
    Label5.Location = New Point(749, 211)
    Label5.Name = "Label5"
    Label5.Size = New Size(159, 25)
    Label5.TabIndex = 10
    Label5.Text = "Target Application:"
    ' 
    ' txtTargetApp
    ' 
    txtTargetApp.Location = New Point(749, 239)
    txtTargetApp.Name = "txtTargetApp"
    txtTargetApp.Size = New Size(150, 31)
    txtTargetApp.TabIndex = 11
    txtTargetApp.Text = "RTMS"
    ' 
    ' Form1
    ' 
    AutoScaleDimensions = New SizeF(10F, 25F)
    AutoScaleMode = AutoScaleMode.Font
    ClientSize = New Size(1140, 415)
    Controls.Add(txtTargetApp)
    Controls.Add(Label5)
    Controls.Add(Label4)
    Controls.Add(Label3)
    Controls.Add(Label2)
    Controls.Add(Label1)
    Controls.Add(txtExternalPumlFilename)
    Controls.Add(txtInternalPumlFilename)
    Controls.Add(btnClose)
    Controls.Add(btnGo)
    Controls.Add(txtPumlFolderName)
    Controls.Add(txtInterfaceFilename)
    Name = "Form1"
    Text = "InterfacesFlowcharts"
    ResumeLayout(False)
    PerformLayout()
  End Sub

  Friend WithEvents txtInterfaceFilename As TextBox
  Friend WithEvents txtPumlFolderName As TextBox
  Friend WithEvents btnGo As Button
  Friend WithEvents btnClose As Button
  Friend WithEvents txtInternalPumlFilename As TextBox
  Friend WithEvents txtExternalPumlFilename As TextBox
  Friend WithEvents Label1 As Label
  Friend WithEvents Label2 As Label
  Friend WithEvents Label3 As Label
  Friend WithEvents Label4 As Label
  Friend WithEvents Label5 As Label
  Friend WithEvents txtTargetApp As TextBox

End Class
