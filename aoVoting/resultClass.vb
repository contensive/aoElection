
Namespace Contensive.Addons.aoVoting
    '
    Public Class resultClass
        '
        Inherits BaseClasses.AddonBaseClass
        '
        Public Overrides Function Execute(ByVal CP As Contensive.BaseClasses.CPBaseClass) As Object
            Dim returnHtml As String = ""
            Try
                Dim cs As BaseClasses.CPCSBaseClass = CP.CSNew
                Dim csWI As BaseClasses.CPCSBaseClass = CP.CSNew
                Dim listHtml As String = ""
                Dim electionID As Integer = CP.Utils.EncodeInteger(CP.Doc.Var(rnElectionID))
                Dim sql As String = ""
                Dim currentPosition As String = ""
                Dim lastPosition As String = ""
                Dim nextPosition As String = ""
                Dim canS As String = ""
                Dim rowPointer As Integer = 0
                Dim rowClass As String = ""
                Dim positionID As Integer = 0
                Dim writeIn As String = ""
                '
                returnHtml += CP.Content.GetCopy("Instructions - Form Election Results", "Select an election from the menu below to see the results.")
                '
                listHtml = CP.Html.li(CP.Html.SelectContent(rnElectionID, "", cnElections))
                listHtml += CP.Html.li(CP.Html.Button(rnButton, btnResult), , "buttonRow")
                returnHtml += CP.Html.ul(listHtml, , "formList")
                '
                If electionID <> 0 Then
                    listHtml = ""
                    '
                    sql = "SELECT m.firstName as fName, m.lastName as lName, p.name as pName, c.id as candidateID, c.positionID as positionID "
                    sql += "FROM candidates c, votingPositions p, ccMembers m "
                    sql += "where (c.electionID=" & electionID & ") and (c.positionID=p.ID) and (c.memberID=m.ID) "
                    sql += "order by p.id"
                    cs.OpenSQL(sql)
                    Do While cs.OK()
                        '
                        '   display position
                        '
                        currentPosition = cs.GetText("pName")
                        If currentPosition <> lastPosition Then
                            listHtml += CP.Html.li(currentPosition, , "positionHeader round4")
                            lastPosition = currentPosition
                            canS = ""
                            positionID = cs.GetInteger("positionID")
                        End If
                        '
                        '   alternate row color
                        '
                        If rowPointer Mod 2 = 0 Then
                            rowClass = "dark"
                        Else
                            rowClass = "light"
                        End If
                        '
                        '   get the candidates
                        '
                        canS += CP.Html.li(cs.GetText("fName") & " " & cs.GetText("lName") & " - " & getVoteCount(CP, cs.GetInteger("candidateID"), electionID) & " votes", , rowClass & " round4")
                        cs.GoNext()
                        rowPointer += 1
                        '
                        If cs.OK() Then
                            nextPosition = cs.GetText("pName")
                        Else
                            nextPosition = ""
                        End If
                        '
                        '   check if the position is changing (or cs not ok)
                        '       if so then wrap up the li lists
                        '
                        If nextPosition <> currentPosition Then
                            '
                            '   add any write ins
                            '
                            sql = "SELECT distinct writeIn "
                            sql += "FROM Votes "
                            sql += "where (candidateID=0) and (electionID=" & electionID & ") and (votingPositionID=" & positionID & ") "
                            csWI.OpenSQL(sql)
                            Do While csWI.OK()
                                '
                                '   alternate rows color
                                '
                                If rowPointer Mod 2 = 0 Then
                                    rowClass = "dark"
                                Else
                                    rowClass = "light"
                                End If
                                '
                                writeIn = csWI.GetText("writeIn")
                                '
                                canS += CP.Html.li(writeIn & " - " & getVoteCountByWriteIn(CP, electionID, writeIn, positionID) & " votes", , rowClass & " round4")
                                '
                                csWI.GoNext()
                            Loop
                            csWI.Close()
                            '
                            canS = CP.Html.ul(canS, , "candidateList")
                            listHtml += CP.Html.li(canS)
                            returnHtml += CP.Html.ul(listHtml, , "positionList")
                            canS = ""
                            listHtml = ""
                        End If
                    Loop
                    cs.Close()
                End If
                '
                returnHtml = CP.Html.Form(CP.Html.div(returnHtml, , "resultContainer"))
                '
            Catch ex As Exception
                CP.Site.ErrorReport(ex.Message)
            End Try
            Return returnHtml
        End Function
        '
        Private Function getVoteCount(ByVal CP As Contensive.BaseClasses.CPBaseClass, ByVal candidateID As Integer, ByVal electionID As Integer) As Integer
            Try
                Dim sql As String = ""
                Dim cs As BaseClasses.CPCSBaseClass = CP.CSNew
                '
                sql = "select count(ID) as votes FROM votes where (candidateID=" & candidateID & ") and (electionID=" & electionID & ")"
                cs.OpenSQL(sql)
                If cs.OK() Then
                    Return cs.GetInteger("votes")
                End If
            Catch ex As Exception
                CP.Site.ErrorReport(ex.Message)
            End Try
        End Function
        '
        Private Function getVoteCountByWriteIn(ByVal CP As Contensive.BaseClasses.CPBaseClass, ByVal electionID As Integer, ByVal writeIn As String, ByVal votingPositionID As Integer) As Integer
            Try
                Dim sql As String = ""
                Dim cs As BaseClasses.CPCSBaseClass = CP.CSNew
                '
                sql = "select count(ID) as votes FROM votes where (electionID=" & electionID & ") and (writeIn=" & CP.Db.EncodeSQLText(writeIn) & ") and (votingPositionID=" & votingPositionID & ")"
                cs.OpenSQL(sql)
                If cs.OK() Then
                    Return cs.GetInteger("votes")
                End If
            Catch ex As Exception
                CP.Site.ErrorReport(ex.Message)
            End Try
        End Function
        '
    End Class
    '
End Namespace