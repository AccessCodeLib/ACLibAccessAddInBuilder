Attribute VB_Name = "modVbProject"
'---------------------------------------------------------------------------------------
' Module: modVbProject
'---------------------------------------------------------------------------------------
'/**
' <summary>
' VBProject ermitteln
' </summary>
' <remarks>
' </remarks>
' \ingroup base
'**/
'---------------------------------------------------------------------------------------
'<codelib>
'  <file>%AppFolder%/source/modVbProject.bas</file>
'  <license>_codelib/license.bas</license>
'  <use>file/FileTools.bas</use>
'</codelib>
'---------------------------------------------------------------------------------------
'
Option Compare Text
Option Explicit
Option Private Module

#Const EARLYBINDING = 1

Private m_CurrentVbProject As Object
Private m_CodeVbProject As Object

#If EARLYBINDING Then
Public Property Get CurrentVbProject() As VBIDE.VBProject
#Else
Public Property Get CurrentVbProject() As Object
#End If

   If m_CurrentVbProject Is Nothing Then
      Set m_CurrentVbProject = FindVBProject(CurrentDb)
   End If

   Set CurrentVbProject = m_CurrentVbProject

End Property


#If EARLYBINDING Then
Public Property Get CodeVBProject() As VBIDE.VBProject
#Else
Public Property Get CodeVBProject() As Object
#End If

   If m_CodeVbProject Is Nothing Then
      Set m_CodeVbProject = FindVBProject(CodeDb)
   End If

   Set CodeVBProject = m_CodeVbProject

End Property

#If EARLYBINDING Then
Private Function FindVBProject(ByVal ProjectDb As DAO.Database) As VBIDE.VBProject
#Else
Private Function FindVBProject(ByVal ProjectDb As DAO.Database) As Object
#End If

#If EARLYBINDING Then
   Dim Project As VBProject
   Dim TempProject As VBProject
#Else
   Dim Project As Object
   Dim TempProject As Object
#End If
   Dim DbName As String

   Set Project = VBE.ActiveVBProject

   ' compare path from database and VBProject
   DbName = FileTools.UncPath(ProjectDb.Name)
   If Project.FileName <> DbName Then
      Set Project = Nothing
      For Each TempProject In VBE.VBProjects
         If TempProject.FileName = DbName Then
            Set Project = TempProject
            Exit For
         End If
      Next
   End If

   Set FindVBProject = Project

End Function
