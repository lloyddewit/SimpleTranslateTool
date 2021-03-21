' CLIMSOFT - Climate Database Management System
' Copyright (C) 2019
'
' This program is free software: you can redistribute it and/or modify
' it under the terms of the GNU General Public License as published by
' the Free Software Foundation, either version 3 of the License, or
' (at your option) any later version.
'
' This program is distributed in the hope that it will be useful,
' but WITHOUT ANY WARRANTY; without even the implied warranty of
' MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
' GNU General Public License for more details.
'
' You should have received a copy of the GNU General Public License
' along with this program.  If not, see <http://www.gnu.org/licenses/>.
Imports System.Data.SQLite
Imports System.Windows.Forms

Public Class SimpleTranslateTool

    Public Shared Sub translateForm(clsForm As Form, Optional strLanguage As String = "")
        'connect to the SQLite database that contains the translations
        Dim clsBuilder As New SQLiteConnectionStringBuilder With {
            .FailIfMissing = True,
            .DataSource = My.Settings.translationDatabase
        }
        Using clsConnection As New SQLiteConnection(clsBuilder.ConnectionString)
            clsConnection.Open()
            Using clsCommand As New SQLiteCommand(clsConnection)

                'get all translations for the specified form and language
                clsCommand.CommandText = "SELECT control_name, translation FROM form_controls, translations WHERE form_name = """ & clsForm.Name &
                                         """ AND language_code = """ & strLanguage & """ And form_controls.id_text = translations.id_text"
                Dim clsReader As SQLiteDataReader = clsCommand.ExecuteReader()
                Using clsReader

                    'for each translation row
                    While (clsReader.Read())

                        'translate the control's text to the new language
                        Dim strControlName As String = clsReader.GetString(0)
                        Dim strTranslation As String = clsReader.GetString(1)
                        CallByName(clsForm.Controls(strControlName), "Text", CallType.Set, strTranslation)

                    End While
                End Using
            End Using
            clsConnection.Close()
        End Using

        'Dim clsBuilder As New SQLiteConnectionStringBuilder()
        'Dim command As SQLiteCommand
        'Dim connection As New SQLiteConnection()
        'Dim datareader As SQLiteDataReader

        'clsBuilder.FailIfMissing = True
        'clsBuilder.DataSource = My.Settings.localDatabase
        'connection = New SQLiteConnection(clsBuilder.ConnectionString)
        'connection.Open()
        'command = connection.CreateCommand()

        'command.CommandText = "SELECT control_name, translation FROM rinstat_translations WHERE form_name = """ &
        '                      frm.Name & """ And language_code = """ & language & """"
        'datareader = command.ExecuteReader()
        'While (datareader.Read())
        '    Dim strControlName As String = datareader.GetString(0)
        '    Dim strTranslation As String = datareader.GetString(1)
        '    CallByName(frm.Controls(strControlName), "Text", CallType.Set, strTranslation)
        'End While
    End Sub

    Public Shared Sub translateMenuItem(ctrParent As Control, tsItem As ToolStripItem, Optional strLanguage As String = "")
        'connect to the SQLite database that contains the translations
        Dim clsBuilder As New SQLiteConnectionStringBuilder With {
            .FailIfMissing = True,
            .DataSource = My.Settings.translationDatabase
        }
        Using clsConnection As New SQLiteConnection(clsBuilder.ConnectionString)
            clsConnection.Open()
            Using clsCommand As New SQLiteCommand(clsConnection)

                'get all translations for the specified form and language
                clsCommand.CommandText = "SELECT control_name, translation FROM form_controls, translations WHERE form_name = """ & ctrParent.Name &
                                         """ AND control_name = """ & tsItem.Name & """ AND language_code = """ & strLanguage & """ And form_controls.id_text = translations.id_text"
                Dim clsReader As SQLiteDataReader = clsCommand.ExecuteReader()
                Using clsReader

                    'for each translation row
                    While (clsReader.Read())

                        If clsReader.FieldCount < 2 OrElse clsReader.IsDBNull(1) Then
                            Continue While
                        End If

                        'translate the control's text to the new language
                        Dim strMenuItemName As String = clsReader.GetString(0)
                        Dim strTranslation As String = clsReader.GetString(1)
                        If strMenuItemName = tsItem.Name Then
                            CallByName(tsItem, "Text", CallType.Set, strTranslation)
                        End If

                    End While
                End Using
            End Using
            clsConnection.Close()
        End Using
    End Sub

    Public Shared Function listLanguages() As AvailableLanguages
        Dim builder As New SQLiteConnectionStringBuilder()
        Dim command As SQLiteCommand
        Dim connection As New SQLiteConnection()
        Dim currentIndex As Integer = 0
        Dim datareader As SQLiteDataReader
        Dim dict As New Dictionary(Of String, String)
        Dim languages As New List(Of String)
        Dim languageCodes As New List(Of String)
        Dim selectedIndex As Integer = 0

        Try
            builder.FailIfMissing = True
            builder.DataSource = My.Settings.translationDatabase
            connection = New SQLiteConnection(builder.ConnectionString)
            connection.Open()
            command = connection.CreateCommand()
            command.CommandText = "SELECT code, language FROM available_languages"
            datareader = command.ExecuteReader()
            While (datareader.Read())
                languageCodes.Add(datareader.GetValue(0))
                languages.Add(datareader.GetString(1))
                If datareader.GetValue(0) = "en" Then
                    selectedIndex = currentIndex
                End If
                currentIndex += 1
            End While
        Catch ex As Exception
            ' If processing information from SQLite fails for any reason then return
            ' English as the only option (if a connection is open then close it).
            connection.Close()
            languageCodes.Add("en")
            languages.Add("English")
            selectedIndex = 0
        End Try

        'For Each keyValue As String() In languages.Zip(Of String, String())(languageCodes, Tuple.Create())
        For Each keyValue In languages.Zip(Of String, Object)(languageCodes, Function(s1, s2) Tuple.Create(s1, s2))
            dict.Add(keyValue.Item1, keyValue.Item2)
        Next

        Return New AvailableLanguages With {
                .languageCodes = languageCodes,
                .languages = languages,
                .selectedIndex = selectedIndex,
                .lookupCode = dict
            }
    End Function


    'Public Shared Function translateOpenForms(Optional language As String = "")
    '    Dim languageCode As String
    '    For Each frm As Form In My.Application.OpenForms
    '        languageCode = translateForm(frm, language)
    '    Next
    '    Return languageCode
    'End Function

    'Public Shared Function translateForm(frm As Form, Optional language As String = "")
    '    ' Change text in all open forms to new language choice
    '    Dim builder As New SQLiteConnectionStringBuilder()
    '    Dim command As SQLiteCommand
    '    Dim connection As New SQLiteConnection()
    '    Dim controlName As String
    '    Dim datareader As SQLiteDataReader
    '    Dim flag As String
    '    Dim formName As String
    '    Dim lang As String
    '    Dim languageCode As String
    '    Dim keyParts As String()
    '    Dim parent As String
    '    Dim parents As String()
    '    Dim parentParts As String()
    '    Dim parentType As String
    '    Dim propertyName As String
    '    Dim startswith As String
    '    Dim value As String

    '    If String.IsNullOrEmpty(language) Then
    '        languageCode = My.Settings.currentLanguageCode
    '    Else
    '        languageCode = listLanguages().lookupCode(language)
    '    End If

    '    'Try
    '    builder.FailIfMissing = True
    '    builder.DataSource = My.Settings.localDatabase
    '    connection = New SQLiteConnection(builder.ConnectionString)
    '    connection.Open()
    '    command = connection.CreateCommand()

    '    ' For the current languageCode and form.Name get all translation entries
    '    startswith = String.Format("{0}\_\_{1}\_\_", languageCode, frm.Name)
    '    command.CommandText = String.Format(
    '            "SELECT name, value, parent, flag from translations WHERE name LIKE '{0}%' ESCAPE '\';", startswith)
    '    datareader = command.ExecuteReader()

    '    'MsgBox("SELECT name, value, parent, flag from translations WHERE name LIKE '{0}%' ESCAPE '\';", startswith)

    '    While (datareader.Read())
    '        ' Currenty the strings we are splitting contain two consecutive understored "__"
    '        ' Splitting on `__` is the same as spliting on `_` and will return empty strings between each pair of underscores
    '        keyParts = datareader.GetValue(0).Split("_")
    '        lang = keyParts(0)
    '        formName = keyParts(2)
    '        controlName = keyParts(4)
    '        propertyName = keyParts(6)
    '        value = datareader.GetString(1)
    '        parents = datareader.GetString(2).Split(".")
    '        ' It's always safe to request details of the first parent (root parent)
    '        parentParts = parents(0).Split(":")
    '        parent = parentParts(0)
    '        flag = datareader.GetString(3)

    '        Try
    '            If Not String.IsNullOrEmpty(parent) Then

    '                ' First parent's indicator of type
    '                parentType = parentParts(1)

    '                ' Is there a second parent?
    '                If parents.Length > 1 Then
    '                    ' When there is a second parent, this usually indicates a submenu item
    '                    ' E.g. MenuStrip2:M.mnuInput:S
    '                    Dim secondParentParts = parents(1).Split(":")
    '                    Dim secondParent = secondParentParts(0)
    '                    Dim secondParentType = secondParentParts(1)
    '                    Dim menu As MenuStrip
    '                    Dim toolmenu As ToolStripMenuItem
    '                    menu = frm.Controls(parent)
    '                    toolmenu = menu.Items().Item(secondParent)
    '                    CallByName(toolmenu.DropDownItems().Item(controlName), propertyName, CallType.Set, value)

    '                ElseIf parentType = "C" Then
    '                    ' This will work for standard controls that are inside a panel
    '                    CallByName(frm.Controls(parent).Controls(controlName), propertyName, CallType.Set, value)
    '                ElseIf parentType = "M" Then
    '                    ' This will work for the primary menu items on a menu strip (but not submenu items)
    '                    Dim menu As MenuStrip
    '                    menu = frm.Controls(parent)
    '                    CallByName(menu.Items().Item(controlName), propertyName, CallType.Set, value)
    '                End If

    '            ElseIf controlName = "Me" Then
    '                CallByName(frm, propertyName, CallType.Set, value)
    '            Else
    '                CallByName(frm.Controls(controlName), propertyName, CallType.Set, value)
    '            End If

    '        Catch ex As Exception
    '            '    MsgBox(">> " & parent & " " & frm.Name & " " & controlName & " " & propertyName & " " & value & "::" & ex.Message)
    '        End Try

    '    End While
    '    'Catch ex As Exception
    '    '    Throw
    '    'Finally
    '    connection.Close()
    '    'End Try

    '    Return languageCode
    'End Function
End Class

Public Class AvailableLanguages
    Public Property languageCodes As List(Of String)
    Public Property languages As List(Of String)
    Public Property selectedIndex As Integer
    Public Property lookupCode As Dictionary(Of String, String)

End Class
