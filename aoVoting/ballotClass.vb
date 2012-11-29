
Imports Contensive.BaseClasses

Namespace Contensive.Addons.aoElection
    '
    Public Class ballotClass
        '
        Inherits BaseClasses.AddonBaseClass
        '
        Public Overrides Function Execute(ByVal CP As CPBaseClass) As Object
            Dim returnHtml As String = ""
            Try
                Dim cs As BaseClasses.CPCSBaseClass = CP.CSNew
                Dim csCandidate As BaseClasses.CPCSBaseClass = CP.CSNew
                Dim csElection As BaseClasses.CPCSBaseClass = CP.CSNew
                Dim electionID As Integer = 0
                Dim spElectionID As Integer = CP.Utils.EncodeInteger(CP.Doc.Var(rnElectionID))
                Dim compatibilityElectionID As Integer = CP.Utils.EncodeInteger(CP.Doc.Var(addonArgument_ElectionId))
                Dim candidateID As Integer = CP.Utils.EncodeInteger(CP.Doc.Var(rnRow))
                Dim ballotCount As Integer = CP.Utils.EncodeInteger(CP.Doc.Var(rnBallotCount))
                Dim pastOfficeID As Integer = CP.Utils.EncodeInteger(CP.Doc.Var(rnPastOfficeID))
                Dim writeIn As String = CP.Doc.Var(rnWriteIn)
                Dim errMsg As String = ""
                Dim electionName As String = ""
                Dim isOpen As Boolean = False
                Dim adminHint As String = ""
                Dim presetList As String = ""
                Dim OfficeID As Integer = 0
                Dim criteria As String = ""
                Dim instanceId As String = CP.Doc.GetText("instanceid")
                Dim electionKey As String
                Dim listLink As String = ""
                Dim rightNow As Date = Date.Now
                Dim sqlRightNow As String = CP.Db.EncodeSQLDate(rightNow)
                Dim votingInstructions As String = ""
                Dim listRows As String = ""
                Dim writeInLine As String = ""
                Dim thankyoucopy As String = ""
                '
                ' electionId is stored in a site property, referenced by the instance Id
                '   admin hint editor can change site property
                '
                If instanceId = "" Then
                    instanceId = "page[" & CP.Doc.PageId & "]"
                End If
                electionKey = "election instance " & instanceId
                '
                ' setup upgrade for old compatibility
                '
                If (spElectionID = 0) And (compatibilityElectionID <> 0) Then
                    spElectionID = compatibilityElectionID
                    Call CP.Site.SetProperty(electionKey, spElectionID)
                End If
                electionID = spElectionID
                '
                ' process admin form before loading default
                '
                If electionID <> 0 Then
                    Call CP.Site.SetProperty(electionKey, electionID)
                Else
                    electionID = CP.Site.GetInteger(electionKey)
                End If
                '
                If Not csElection.Open("elections", "id=" & electionID) Then
                    Call csElection.Close()
                    csElection.Insert("elections")
                    electionID = csElection.GetInteger("id")
                    Call csElection.SetField("name", "New Election")
                    Call csElection.SetField("dateStart", Now)
                    Call csElection.SetField("dateEnd", Now.AddMonths(1))
                    Call csElection.SetField("thankyoucopy", "<p>Thank you for participating.</p>")
                    Call CP.Site.SetProperty(electionKey, electionID)
                End If
                If Not csElection.OK() Then
                    returnHtml = getCopyElectionNotFound(CP)
                    adminHint &= "<p>The election selected for this add-on could not be found. Turn on advanced edit and click the add-on options to check or update the election selected.</p>"
                Else
                    electionName = csElection.GetText("name")
                    thankyoucopy = csElection.GetText("thankyoucopy")
                    isOpen = isElectionOpen(CP, csElection)
                    '
                    '   if election is current continue otherwise block the balott to non admins
                    '
                    If (Not isOpen) Then
                        returnHtml = getCopyElectionClosed(CP)
                    Else
                        If ((candidateID = 0) And (writeIn = "")) And (CP.Doc.Var(rnButton) = btnContinue) Then
                            ballotCount -= 1
                            errMsg = "Please select a candidate for this Offices or include a write-in."
                        ElseIf (CP.Doc.Var(rnButton) = btnContinue) Then
                            processVote(CP, electionID, pastOfficeID, candidateID, writeIn)
                        End If
                        '
                        '   the ballot will loop through all the candidates for each Office until there are none left
                        '       when there are none left - the thank you message is displayed
                        '
                        Call getOfficeID(CP, electionID, ballotCount, OfficeID, votingInstructions)
                        criteria = "(electionID=" & electionID & ") and (OfficeID=" & OfficeID & ")"
                        '
                        If OfficeID = 0 Then
                            returnHtml &= thankyoucopy
                        Else
                            cs.Open("Election Candidates", criteria)
                            If cs.OK Then
                                '
                                If errMsg <> "" Then
                                    listRows &= "<p class=""ccError"">" & errMsg & "</p>"
                                End If
                                Do While (cs.OK)
                                    listRows &= "<div class=""ebRow"">" & getCandidateRow(CP, cs) & "</div>"
                                    cs.GoNext()
                                Loop
                                If CP.User.IsEditingAnything() Then
                                    listRows &= "<div class=""ebRow"">" & cs.GetAddLink("electionid=" & electionID) & "&nbsp;Add a candidate to this election</div>"
                                End If
                                'returnHtml &= "</table>"
                            End If
                            cs.Close()
                            '
                            writeInLine = "" _
                                & "<div class=""ebRow"">" _
                                & CP.Html.div(CP.Html.RadioBox(rnRow, "-1", "") & "" _
                                & "&nbsp;Write-In Candidate:" _
                                & CP.Html.InputText(rnWriteIn, "", , 25).Replace("<input ", "<input onClick=""clearVote('" & rnRow & "');"" ")) & "" _
                                & "</div>"
                            '
                            returnHtml = "" _
                                & "<p class=""ebOffice"">" & CP.Content.GetRecordName("election offices", OfficeID) & "</p>" _
                                & ""
                            If votingInstructions <> "" Then
                                returnHtml &= "<div class=""ebInstructions"">" & votingInstructions & "</div>"
                            End If
                            returnHtml &= "" _
                                & "<div class=""ebList"">" _
                                & listRows _
                                & writeInLine _
                                & "</div>" _
                                & ""
                            returnHtml &= CP.Html.div(CP.Html.Hidden(rnPastOfficeID, OfficeID) & CP.Html.Hidden(rnBallotCount, ballotCount + 1) & CP.Html.Button(rnButton, btnContinue), , "ebButtonContainer")
                            returnHtml = CP.Html.Form(returnHtml)
                        End If
                        If CP.User.IsAdmin() And (Not isOpen) Then
                            adminHint &= "<p>This election is closed.</p>"
                        End If
                    End If
                End If
                csElection.Close()
                '
                If CP.User.IsAdmin Then
                    Dim editList As String = ""
                    Dim s = CP.Html.SelectContent(rnElectionID, electionID, "elections", "((dateStart is null)or(dateStart<" & sqlRightNow & "))and((dateEnd is null)or(dateend>" & sqlRightNow & "))", "Election created for this page", , "ebSelectElection")
                    s = CP.Html.Form(s, , , "ebSelectElectionForm")
                    editList &= CP.Html.li("Select an election " & s, , "ebHelp")
                    editList &= CP.Html.li("Edit this election " & CP.Content.GetEditLink("elections", electionID, False, "", True), , "ebHelp")
                    '
                    'adminHint &= CP.Html.div("Select an election " & s)
                    '
                    'adminHint &= CP.Html.div("Edit this election " & CP.Content.GetEditLink("elections", electionID, False, "", True))
                    '
                    listLink = "<a class=""ccRecordEditLink"" tabindex=""-1"" href=""" & CP.Site.GetText("adminUrl") & "?cid=" & CP.Content.GetID("election offices") & """><img src=""/ccLib/images/IconContentEdit.gif"" border=""0"" alt=""Add or Modify Election Offices"" title=""Add or Modify Election Offices"" align=""absmiddle""></a>"
                    editList &= CP.Html.li("Add or Modify Election Offices" & listLink, , "ebHelp")
                    'adminHint &= CP.Html.div("Add or Modify Election Office" & listLink)
                    '
                    listLink = "<a class=""ccRecordEditLink"" tabindex=""-1"" href=""" & CP.Site.GetText("adminUrl") & "?cid=" & CP.Content.GetID("election candidates") & """><img src=""/ccLib/images/IconContentEdit.gif"" border=""0"" alt=""Add or Modify Candidates"" title=""Add or Modify Candidates"" align=""absmiddle""></a>"
                    editList &= CP.Html.li("Add or Modify Candidates" & listLink, , "ebHelp")
                    'adminHint &= CP.Html.div("Add or Modify Candidates" & listLink)
                    '
                    listLink = "<a href=""" & CP.Site.GetText("adminUrl") & "?addonGuid={BD797028-938B-4E7D-84FE-F42257AEB461}"">Election Reports</a>"
                    editList &= CP.Html.li("View Election Results " & listLink, , "ebHelp")
                    'adminHint &= CP.Html.div("View Election Results " & listLink)
                    '
                    adminHint &= CP.Html.ul(editList)
                    adminHint &= adminInstructions
                    returnHtml &= getAdminHint(CP, adminHint)
                    'returnHtml &= CP.Html.adminHint(adminHint)
                End If
                '
                returnHtml = CP.Html.div(returnHtml, , "electionBallot")
                '
            Catch ex As Exception
                CP.Site.ErrorReport(ex.Message)
            End Try
            Return returnHtml
        End Function
        '
        '
        '
        Private Function getCandidateRow(ByVal CP As CPBaseClass, ByVal cs As BaseClasses.CPCSBaseClass) As String
            Dim returnHtml As String = ""
            Try
                Dim scriptString As String = ""
                Dim recID As Integer = cs.GetInteger("ID")
                Dim bioID As String = "bio_" & recID
                Dim linkID As String = "link_" & recID
                Dim nameString As String = cs.GetText("Name")
                Dim thumbnailFilename As String = cs.GetText("thumbnailFilename")
                Dim imageFilename As String = ""
                Dim bioString As String = ""
                Dim lightBox As String = ""
                Dim copy As String = ""
                Dim editLink As String = cs.GetEditLink()
                Dim thumbnailImg As String = ""
                Dim imageImg As String = ""
                '
                If thumbnailFilename <> "" Then
                    thumbnailImg = "<img class=""ebPhoto"" src=""" & CP.Site.FilePath & thumbnailFilename & """>"
                Else
                    thumbnailImg = "<img class=""ebPhoto"" src=""/voting/thumbnailDefault.png"">"
                End If
                imageFilename = cs.GetText("imageFilename")
                If imageFilename = "" Then
                    imageImg = thumbnailImg
                Else
                    imageImg = "<img class=""ebPhoto"" src=""" & CP.Site.FilePath & imageFilename & """>"
                End If
                bioString = imageImg
                '
                returnHtml = editLink & CP.Html.RadioBox(rnRow, cs.GetText("ID"), "") & "&nbsp;" & "<span class=""ebName"">" & nameString & "</span>"
                bioString &= CP.Html.h1(nameString)
                '
                copy = cs.GetText("Title")
                If copy <> "" Then
                    returnHtml &= ", " & copy
                    bioString &= CP.Html.p(copy, , "title")
                End If
                '
                copy = cs.GetText("company")
                If copy <> "" Then
                    returnHtml &= ", " & copy
                    bioString &= CP.Html.p(copy, , "company")
                End If
                '
                copy = cs.GetText("phone")
                If copy <> "" Then
                    bioString &= CP.Html.p("Phone: " & copy, , "phone")
                End If
                '
                copy = cs.GetText("bio")
                If copy <> "" Then
                    bioString &= CP.Html.div(copy, , "notes")
                End If
                '
                lightBox = CP.Html.div(CP.Html.div(bioString, , "ebBioContainer", bioID), , "ebBioWrapper")
                returnHtml &= "&nbsp;" _
                    & lightBox _
                    & "(<a id=""" & linkID & """ href=""#" & bioID & """ title=""" & nameString & " Bio"">View</a>)" _
                    & ""
                '
                'returnHtml = returnHtml & "<tr>"
                'returnHtml = returnHtml & "<td class=""radioContainer"">" & CP.Html.RadioBox(rnRow, cs.GetText("ID"), "") & lightBox & editLink & "</td>"
                'returnHtml = returnHtml & "<td class=""imgContainer"">" & thumbnailImg & "</td>"
                'returnHtml = returnHtml & "<td class=""nameContainer""><a id=""" & linkID & """ href=""#" & bioID & """ title=""" & nameString & " Bio"">" & nameString & "</a></td>"
                'returnHtml = returnHtml & "<td class=""companyContainer"">" & cs.GetText("Company") & "</td>"
                'returnHtml = returnHtml & "</tr>"
                '
                scriptString &= vbCrLf & "$('#" & linkID & "').fancybox({" & vbCrLf
                scriptString &= "'titleShow':false," & vbCrLf
                scriptString &= "'transitionIn':'fade'," & vbCrLf
                scriptString &= "'transitionOut':'fade'," & vbCrLf
                scriptString &= "'overlayOpacity':'.6'," & vbCrLf
                scriptString &= "'overlayColor':'#000000'" & vbCrLf
                scriptString &= "});" & vbCrLf & vbCrLf
                '
                CP.Doc.AddHeadJavascript("$(document).ready(function() {" & scriptString & "});")
                '
            Catch ex As Exception
                CP.Site.ErrorReport(ex.Message)
            End Try
            Return returnHtml
        End Function

        Private Sub getOfficeID(ByVal CP As CPBaseClass, ByVal electionID As Integer, ByVal ballotCount As Integer, ByRef officeId As Integer, ByRef votingInstructions As String)
            Try
                Dim cs As BaseClasses.CPCSBaseClass = CP.CSNew
                Dim loopCount As Integer = 0
                Dim sql As String = ""
                Dim criteria As String = "(1=1)"
                '
                '   this will allow users to come back and complete their ballot at a later point
                '       showing them only the Offices left to vote on
                '
                sql = "select electionOfficeID from electionVotes where memberID=" & CP.User.Id & " and electionID=" & electionID
                cs.OpenSQL(sql)
                Do While cs.OK()
                    criteria += " and (OfficeID<>" & cs.GetInteger("electionOfficeID") & ")"
                    cs.GoNext()
                Loop
                cs.Close()
                '
                '   add the criteria into the select for the next Office
                '
                sql = "SELECT DISTINCT OfficeID FROM electionCandidates WHERE (electionID=" & electionID & ") and " & criteria & " ORDER BY OfficeID"
                cs.OpenSQL(sql)
                If cs.OK Then
                    officeId = cs.GetInteger("OfficeID")
                End If
                cs.Close()
                '
                If officeId <> 0 Then
                    If cs.Open("election offices", "id=" & officeId) Then
                        votingInstructions = cs.GetText("votingInstructions")
                    End If
                    Call cs.Close()

                End If
                '
            Catch ex As Exception
                CP.Site.ErrorReport(ex.Message)
            End Try
        End Sub
        '
        Private Sub processVote(ByVal CP As CPBaseClass, ByVal electionID As Integer, ByVal officeID As Integer, ByVal candidateID As Integer, ByVal writeIn As String)
            Try
                Dim cs As BaseClasses.CPCSBaseClass = CP.CSNew
                '
                cs.Open("Election Votes", "(electionID=" & electionID & ") and (electionOfficeID=" & officeID & ") and (memberID=" & CP.User.Id & ")")
                If Not cs.OK Then
                    cs.Close()
                    cs.Insert("Election Votes")
                End If
                If cs.OK Then
                    cs.SetField("electionID", electionID)
                    cs.SetField("electionOfficeID", officeID)
                    cs.SetField("candidateID", candidateID)
                    cs.SetField("writeIn", writeIn)
                    cs.SetField("memberID", CP.User.Id)
                    cs.SetField("visitID", CP.Visit.Id)
                End If
                cs.Close()
            Catch ex As Exception
                CP.Site.ErrorReport(ex.Message)
            End Try
        End Sub
        '
        Private Function isElectionOpen(ByVal CP As CPBaseClass, ByVal cs As BaseClasses.CPCSBaseClass) As Boolean
            Dim returnValid As Boolean = False
            Try
                Dim startDate As Date = cs.GetDate("dateStart")
                Dim enddate As Date = cs.GetDate("dateEnd")
                Dim currentDate As Date = Date.Now()
                '
                If (startDate < #1/1/2000#) And (enddate < #1/1/2000#) Then
                    returnValid = True
                ElseIf (currentDate > startDate) And (currentDate < enddate) Then
                    returnValid = True
                End If
            Catch ex As Exception
                CP.Site.ErrorReport(ex, "unexpected error in isSelectionValid")
            End Try
            Return returnValid
        End Function
        '
        Private Function getCopyElectionNotFound(ByVal cp As CPBaseClass)
            Return cp.Content.GetCopy("Election Not Found", "<p>The ballot you requested is currently unavailable.</p>")
        End Function
        '
        Private Function getCopyElectionClosed(ByVal cp As CPBaseClass)
            Return cp.Content.GetCopy("Election Closed", "<p>The ballot you requested is closed.</p>")
        End Function

        '
        Private Function getAdminHint(ByVal cp As CPBaseClass, ByVal content As String) As String
            Dim returnHtml As String = ""
            '
            returnHtml = "" _
                & "<table border=0 width=""100%"" cellspacing=0 cellpadding=0><tr><td class=""ccHintWrapper"">" _
                    & "<table border=0 width=""100%"" cellspacing=0 cellpadding=0><tr><td class=""ccHintWrapperContent"">" _
                    & "<b>Administrator</b>" _
                    & "<BR>" _
                    & "<BR>" & cp.Utils.EncodeText(content) _
                    & "</td></tr></table>" _
                & "</td></tr></table>"

            Return returnHtml
        End Function
    End Class
    '
End Namespace