VERSION 5.00
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H8000000C&
   Caption         =   "MDIForm1"
   ClientHeight    =   5370
   ClientLeft      =   165
   ClientTop       =   810
   ClientWidth     =   6240
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   Begin VB.Menu show 
      Caption         =   "Show"
   End
   Begin VB.Menu showdialog 
      Caption         =   "ShowDialog"
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim WithEvents CSharpInteropServiceEvents As CSharpInteropService.LibraryInvoke
Attribute CSharpInteropServiceEvents.VB_VarHelpID = -1

Private Sub CSharpInteropServiceEvents_MessageEvent(ByVal message As String)
   If message = "OpenForm1" Then
      Form1.show
   End If
End Sub

Private Sub show_Click()
   Dim param(0) As Variant
   param(0) = Me.hWnd
   
   Dim load As New LibraryInvoke
   Set CSharpInteropServiceEvents = load
  
   load.GenericInvoke "C:\Temp\CSharpInterop\ClassLibrary1\ClassLibrary1\bin\Debug\ClassLibrary1.dll", "ClassLibrary1.Class1", "ShowFormParent", param
End Sub

Private Sub showdialog_Click()
   Dim retorno As Variant
   Dim param(3) As Variant
   param(0) = True
   param(1) = 1
   param(2) = 1.5
   param(3) = "OK"
   
   Dim load As New LibraryInvoke
   retorno = load.GenericInvoke("C:\Temp\CSharpInterop\ClassLibrary1\ClassLibrary1\bin\Debug\ClassLibrary1.dll", "ClassLibrary1.Class1", "ShowForm", param)

   MsgBox "Return: " & vbCrLf & vbCrLf & retorno(0) & vbCrLf & retorno(1) & vbCrLf & retorno(2)
End Sub
