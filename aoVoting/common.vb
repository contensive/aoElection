
Module common
    '
    Public Const cr = vbCrLf & vbTab
    Public Const cr2 = cr & vbTab
    Public Const cr3 = cr & vbTab
    '
    Public Const addonArgument_ElectionId = "election"
    '
    Public Const rnPastOfficeID = "fld1"
    Public Const rnButton = "fld2"
    Public Const rnRow = "fld3"
    Public Const rnBallotCount = "fld4"
    Public Const rnWriteIn = "fld5"
    Public Const rnElectionID = "elid"
    '
    Public Const btnContinue = " Continue "
    Public Const btnResult = " Get Results "
    '
    Public Const cnElections = "Elections"
    '
    Public Const adminInstructions = "" _
        & cr & "<h3>Instructions</h3>" _
        & cr & "<p>Use this addon to create an election where each visitor can vote for one candidate for each of one or more offices. A visitor can only vote once.</p>" _
        & cr & "<ol>" _
        & cr2 & "<li><p>Start by dropping this add-on on a page. It will automatically create a new election. Edit it to set the name, start and edit date, like 'Annual Board Election'.</p></li>" _
        & cr2 & "<li><p>Create election offices, like 'President' and 'Vice President'.</p></li>" _
        & cr2 & "<li><p>Create a candidate entry for each candidate in the election.</p></li>" _
        & cr & "</ol>"
End Module
