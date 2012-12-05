


Namespace Contensive.Addons.aoElection
    '
    Public Class resultClass
        '
        Inherits BaseClasses.AddonBaseClass
        '
        Public Overrides Function Execute(ByVal CP As Contensive.BaseClasses.CPBaseClass) As Object
            Dim returnHtml As String = ""
            Try
                Dim pageBody As String = ""
                Dim cs As BaseClasses.CPCSBaseClass = CP.CSNew
                Dim csWI As BaseClasses.CPCSBaseClass = CP.CSNew
                Dim listHtml As String = ""
                Dim electionID As Integer = CP.Utils.EncodeInteger(CP.Doc.Var(rnElectionID))
                Dim sql As String = ""
                Dim currentOffice As String = ""
                Dim lastOffice As String = ""
                Dim nextOffice As String = ""
                Dim canS As String = ""
                Dim rowPointer As Integer = 0
                Dim rowClass As String = ""
                Dim officeID As Integer = 0
                Dim writeIn As String = ""
                Dim page As New adminFramework.formSimpleClass
                '
                listHtml = CP.Html.li(CP.Html.SelectContent(rnElectionID, "", cnElections))
                listHtml += CP.Html.li(CP.Html.Button(rnButton, btnResult), , "buttonRow")
                pageBody += CP.Html.ul(listHtml, , "formList")
                '
                If electionID <> 0 Then
                    listHtml = ""
                    '
                    sql = "SELECT c.Name as Name, p.name as pName, c.id as candidateID, p.id as officeID "
                    sql += " FROM electionCandidates c, electionOffices p"
                    sql += " where (c.electionID=" & electionID & ") and (c.officeID=p.ID)"
                    sql += " order by p.id"
                    cs.OpenSQL(sql)
                    Do While cs.OK()
                        '
                        '   display Office
                        '
                        currentOffice = cs.GetText("pName")
                        If currentOffice <> lastOffice Then
                            listHtml += CP.Html.li(currentOffice, , "positionHeader round4")
                            lastOffice = currentOffice
                            canS = ""
                            officeID = cs.GetInteger("officeID")
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
                        canS += CP.Html.li(cs.GetText("Name") & " - " & getVoteCount(CP, cs.GetInteger("candidateID"), electionID) & " votes", , rowClass & " round4")
                        cs.GoNext()
                        rowPointer += 1
                        '
                        If cs.OK() Then
                            nextOffice = cs.GetText("pName")
                        Else
                            nextOffice = ""
                        End If
                        '
                        '   check if the office is changing (or cs not ok)
                        '       if so then wrap up the li lists
                        '
                        If nextOffice <> currentOffice Then
                            '
                            '   add any write-ins
                            '
                            sql = "SELECT distinct writeIn "
                            sql += "FROM electionVotes "
                            sql += "where (writein is not null)and(writein<>'')and(candidateID<=0)and(electionID=" & electionID & ") and (electionOfficeID=" & officeID & ") "
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
                                canS += CP.Html.li("Write-In: " & writeIn & " - " & getVoteCountByWriteIn(CP, electionID, writeIn, officeID) & " votes", , rowClass & " round4")
                                '
                                csWI.GoNext()
                            Loop
                            csWI.Close()
                            '
                            canS = CP.Html.ul(canS, , "candidateList")
                            listHtml += CP.Html.li(canS)
                            pageBody += CP.Html.ul(listHtml, , "positionList")
                            canS = ""
                            listHtml = ""
                        End If
                    Loop
                    cs.Close()
                End If
                '
                pageBody = CP.Html.Form(CP.Html.div(pageBody, , "resultContainer"))
                page.title = "Election Results"
                page.description = "<p>Select an election from the menu below to see the results.</p>"
                page.body = pageBody
                '
                returnHtml = CP.Html.div(page.getHtml(CP), , , "afw")
                Call CP.Doc.AddHeadStyle(page.styleSheet)
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
                sql = "select count(ID) as votes FROM electionVotes where (candidateID=" & candidateID & ") and (electionID=" & electionID & ")"
                cs.OpenSQL(sql)
                If cs.OK() Then
                    Return cs.GetInteger("votes")
                End If
            Catch ex As Exception
                CP.Site.ErrorReport(ex.Message)
            End Try
        End Function
        '
        Private Function getVoteCountByWriteIn(ByVal CP As Contensive.BaseClasses.CPBaseClass, ByVal electionID As Integer, ByVal writeIn As String, ByVal electionOfficeId As Integer) As Integer
            Try
                Dim sql As String = ""
                Dim cs As BaseClasses.CPCSBaseClass = CP.CSNew
                '
                sql = "select count(ID) as votes FROM electionVotes where (electionID=" & electionID & ") and (writeIn=" & CP.Db.EncodeSQLText(writeIn) & ") and (electionOfficeId=" & electionOfficeId & ")"
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