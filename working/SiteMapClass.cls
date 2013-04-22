VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SiteMapClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

Option Explicit
'
Const siteMapVersion = 1
'
'Private Main As MainClass
Private Csv As Object
Private Main As Object
'
Const cr = vbCrLf & vbTab
'
Type NodeType
    CopyPresent As Boolean
    PageLink As String
    MenuHeadline As String
    PageID As Long
    PageTemplateID As Long
    Link As String
End Type
'
'=================================================================================
'   Execute Method, v3.4 Interface
'=================================================================================
'
Public Function Execute(CsvObject As Object, MainObject As Object, OptionString As String, FilterInput As String) As String
    On Error GoTo ErrorTrap
    '
    Dim indexPageName As String
    'Dim siteMapName As String
    Dim isEditing As Boolean
    Dim Controls As String
    Dim instanceId As String
    Dim remote As RemoteClass
    '
    Set Csv = CsvObject
    Set Main = MainObject
    '
    ' update
    '
    If siteMapVersion > kmaEncodeInteger(Main.GetSiteProperty("SiteMap Version", "0")) Then
        '
        ' rename path to sitemapfiles because collision with link alias
        '
        Set remote = New RemoteClass
        Call remote.ExecuteCmd("deletefolder", "folder=" & Main.physicalwwwpath & "sitemap")
        Set remote = Nothing
        Call Main.SetSiteProperty("SiteMap Version", siteMapVersion)
    End If
    
    '
    indexPageName = Csv.GetAddonOption("IndexPageName", OptionString)
    instanceId = Csv.GetAddonOption("instanceID", OptionString)
    If instanceId = "" Then
        instanceId = "unknown"
    End If
    'siteMapName = Csv.GetAddonOption("New Site Map", OptionString)
    'If siteMapName = "" Then
    '    siteMapName = Csv.GetAddonOption("Site Map", OptionString)
    '    If siteMapName = "" Then
    '        siteMapName = "Default"
    '    End If
    'End If
    isEditing = Main.isEditing("Site Maps")
    Execute = GetSectionBasedSiteMap(indexPageName, instanceId, isEditing)
    If isEditing Then
            '
            ' General Controls
            '
            Controls = Controls _
                & "<h2>Site Map Addon</h2>" _
                & "<p>To remove sections of this site map, check the box next to the name of any section and it will not be displayed on this site map.</p>" _
                & ""
            Execute = Execute & Main.GetAdminHintWrapper(Controls)
    End If
    '
    Exit Function
ErrorTrap:
    Call HandleError("SiteMapClass", "Execute", Err.Number, Err.Source, Err.Description, True, False)
End Function
'
' 12/20/2008
'       fixed - if a sections page contentid is not valid, the section does not appear
'       changed GetPageTree so it loads an array, then prints the array so know what icon to use for the lines
'
'========================================================================
' legacy - left only for binary compatibility
'========================================================================
'
Public Sub Init(MainObject As Object)
End Sub
'
Public Function GetContent(OptionString As String) As String
End Function

'----------------------------------------------------------------------------
'   Populates the site map from the given Page Name
'----------------------------------------------------------------------------
Private Function GetPageTree(Link As String, ContentName As String, RootPageID As Long, SectionCaption As String, NumberofTiersDisplayed As Long, SectionID As Long, isEditing As Boolean, siteMapId As Long) As String
    On Error GoTo ErrorTrap
    '
    Dim CS As Long
    Dim optionCell As String
    Dim NewpageLink As String
    Dim Criteria As String
    Dim CSPage As Long
    Dim SectionHeadlineDisplay As String
    Dim Headline As String
    Dim PageID As Long
    Dim Stream As String
    Dim Counter As Long
    Dim SQL As String
    Dim MenuHeadline As String
    Dim SecondPointer As Long
    Dim PageTemplateID As Long
    Dim TemplatePointer As Long
    Dim CopyPresent As String
    Dim PageLink As String
    Dim ChildListSortMethodID As Long
    Dim OrderByCriteria As String
    Dim Nodes() As NodeType
    Dim NodeCnt As Long
    Dim NodePtr As Long
    '
    SectionHeadlineDisplay = SectionCaption
    CSPage = Main.OpenCSContentRecord(ContentName, RootPageID, , , "MenuHeadline,Name,ID,ChildListSortMethodID")
    If Main.CSOK(CSPage) Then
       '
        ' Set Root Menu entry to firs child page
        '
        MenuHeadline = Main.GetCSText(CSPage, "MenuHeadline")
        ChildListSortMethodID = Main.GetCSInteger(CSPage, "ChildListSortMethodID")
        PageID = Main.GetCSInteger(CSPage, "ID")
        '
        If MenuHeadline = "" Then
            MenuHeadline = Main.GetCSText(CSPage, "Name")
        End If
        MenuHeadline = Main.EncodeContent(MenuHeadline, Main.memberID, , False, , True, True, , True)
        If (SectionHeadlineDisplay = "") Then
            SectionHeadlineDisplay = MenuHeadline
        End If
        OrderByCriteria = Main.GetSortMethodByID(ChildListSortMethodID)
        '
        If Main.SiteProperty_BuildVersion > "3.3.752" Then
            Criteria = "(AllowInMenus<>0)and(ParentID=" & PageID & ")and(DateArchive Is Null)"
        Else
            Criteria = "(Menuheadline<>'')and(ParentID=" & PageID & ")and(DateArchive Is Null)"
        End If
    End If
    Call Main.CloseCS(CSPage)
    '
    If Criteria <> "" Then
        '
        ' Loop through child pages
        '
        SecondPointer = Main.OpenCSContent(ContentName, Criteria, OrderByCriteria, , , , "Name,ID,MenuHeadline")
        NodePtr = 0
        Do While Main.CSOK(SecondPointer)
            '
            ReDim Preserve Nodes(NodePtr)
            With Nodes(NodePtr)
                .MenuHeadline = Main.GetCSText(SecondPointer, "MenuHeadline")
                If .MenuHeadline = "" Then
                    .MenuHeadline = Main.GetCSText(SecondPointer, "Name")
                End If
                .MenuHeadline = Main.EncodeContent(.MenuHeadline, Main.memberID, , False, , True, True, , True)
                .MenuHeadline = Main.ConvertHTML2Text(.MenuHeadline)
                .MenuHeadline = Replace(.MenuHeadline, vbCrLf, "")
                .PageID = Main.GetCSInteger(SecondPointer, "ID")
            End With
            NodePtr = NodePtr + 1
            Main.NextCSRecord (SecondPointer)
        Loop
        Call Main.CloseCS(SecondPointer)
        NodeCnt = NodePtr
        If NodeCnt > 0 Then
            For NodePtr = 0 To NodeCnt - 1
                With Nodes(NodePtr)
                If .MenuHeadline <> "" Then
                    NewpageLink = Main.GetPageLink(.PageID)
                    Stream = Stream & "<table border=""0"" width=""100%"" cellspacing=""0"" cellpadding=""0""><tr>"
                    If isEditing Then
                        Stream = Stream & "<td class=""ccSiteMapOptionCell"">&nbsp;</td>"
                    End If
                    If NodePtr < (NodeCnt - 1) Then
                        Stream = Stream & "<td class=""ccSiteMapImageCell""><img src=""/sitemapFiles/tee.gif"" width=""21"" height=""20""></td>"
                    Else
                        Stream = Stream & "<td class=""ccSiteMapImageCell""><img src=""/sitemapFiles/end.gif""></td>"
                    End If
                    Stream = Stream & "<td class=""ccSiteMapChildCell""><a class=ccSiteMapLink href=""" & NewpageLink & """ class=""ccSiteMapChildCell"">" & .MenuHeadline & "</a></td>"
                    Stream = Stream & "</tr></table>"
                End If
                End With
            Next
        End If
    End If
    '
    optionCell = ""
    If isEditing Then
        optionCell = cr & "<td class=""ccSiteMapOptionCell""><input type=""checkbox"" name="""" value=""1"" onClick=""var q='sitemapid=" & siteMapId & "&sectionid=" & SectionID & "';if(this.checked){q+='&action=add'}else{q+='&action=del'};cj.ajax.addon('sitemapclickajax',q);"""
        Criteria = "(active<>0)and(sectionid=" & SectionID & ")and(siteMapId=" & siteMapId & ")"
        CS = Csv.OpenCSContent("Site Map Section Rules", Criteria)
        If Csv.IsCSOK(CS) Then
            optionCell = optionCell & " checked=""checked"""
        End If
        Call Csv.CloseCS(CS)
        optionCell = optionCell & "></td>"
    End If
    PageLink = Main.GetPageLink(RootPageID)
    Stream = "" _
        & cr & "<table border=""0"" width=""100%"" cellspacing=""0"" cellpadding=""0"">" _
        & cr & "<tr>" _
        & kmaIndent(optionCell) _
        & cr & vbTab & "<td class=""ccSiteMapFolderCell""><a href=""" & PageLink & """><img border=""0"" src=""/sitemapFiles/folder.gif"" width=""21"" height=""20""></a></td>" _
        & cr & vbTab & "<td class=""ccSiteMapSectionCell""><a class=""ccSiteMapLink"" href=""" & PageLink & """><b>" & SectionHeadlineDisplay & "</b></a></td>" _
        & cr & "</tr>" _
        & cr & "</table>" _
        & Stream _
        & ""
    '
    GetPageTree = vbCrLf & "<!-- GetPageTree v2.0 -->" & vbCrLf & Stream & vbCrLf & "<!-- GetPageTree v2.0 -->" & vbCrLf
    '
    Exit Function
ErrorTrap:
    Call HandleError("SiteMapClass", "GetPageTree", Err.Number, Err.Source, Err.Description, True, False)
End Function
'-----------------------------------------------------
    '--- Prints the Site Map Based on Sections
'-----------------------------------------------------
Private Function GetSectionBasedSiteMap(indexPageName As String, instanceId As String, isEditing As Boolean) As String
    On Error GoTo ErrorTrap
    '
    Dim ptr As Long
    Dim cnt As Long
    Dim sortMethods() As String
    Dim RootPageID As Long
    Dim CSSectionPointer As Long
    Dim RootPageName As String
    Dim ContentName As String
    Dim urlLink As String
    Dim SQL As String
    Dim SectionID As Long
    Dim Stream As String
    Dim Caption As String
    Dim ContentID As Long
    Dim PageContentCID As Long
    Dim IsAdmin As Boolean
    Dim BlockSection As Boolean
    Dim SortOrderID As Long
    Dim DefaultSortMethod As String
    Dim SortMethod As String
    Dim IsOldNameMatchMode As Boolean
    Dim siteMapId As Long
    Dim CS As Long
    '
    If instanceId = "" Then
        instanceId = "unknown"
    End If
    siteMapId = Csv.GetRecordID("site maps", instanceId)
    If siteMapId = 0 Then
        'If LCase(siteMapName) <> "default" Then
        '    siteMapName = "Default"
        '    siteMapId = Csv.GetRecordID("site maps", "Default")
        'End If
        If siteMapId = 0 Then
            CS = Csv.InsertCSRecord("Site Maps", 0)
            If Csv.IsCSOK(CS) Then
                siteMapId = Csv.GetCSInteger(CS, "id")
                Call Csv.SetCS(CS, "name", instanceId)
            End If
            Call Csv.CloseCS(CS)
        End If
    End If
    IsAdmin = Main.IsAdmin()
    'IsOldNameMatchMode = (Main.SiteProperty_BuildVersion < "3.3.613")
    SortMethod = Trim(Main.GetContentProperty("Site Sections", "DefaultSortMethod"))
    If SortMethod <> "" Then
        If InStr(1, SortMethod, ",") = 0 Then
            SortMethod = "s." & SortMethod
        Else
            sortMethods = Split(SortMethod, ",")
            SortMethod = ""
            cnt = UBound(sortMethods) + 1
            For ptr = 0 To cnt - 1
                If sortMethods(ptr) <> "" Then
                    SortMethod = SortMethod & ",s." & sortMethods(ptr)
                End If
            Next
            If SortMethod <> "" Then
                SortMethod = Mid(SortMethod, 2)
            End If
        End If
    End If
    If SortMethod = "" Then
        SortMethod = "s.id"
    End If
    If isEditing Then
        '
        ' editing, do not block any sections
        '
        SQL = "Select 0 as blocked,s.BlockSection, s.Name,s.RootPageID, s.ID As SectionID,s.Caption,s.contentID, s.TemplateID, t.link" _
                & " From ((ccSections s" _
                & " left join ccTemplates t on s.TemplateID = t.ID)" _
                & " )" _
                & " Where " _
                    & " (s.Active<>0)" _
                    & " and((s.HideMenu is null)OR(s.HideMenu=0))" _
                & " order by " & SortMethod _
            & ""
    Else
        SQL = "Select " _
            & "(select top 1 id from siteMapSectionRules where (active<>0)and(sectionid=s.id)and(sitemapid=" & siteMapId & ")) as blocked" _
            & " ,s.BlockSection" _
            & " ,s.Name" _
            & " ,s.RootPageID" _
            & " ,s.ID As SectionID" _
            & " ,s.Caption" _
            & " ,s.contentID" _
            & " ,s.TemplateID" _
            & " ,t.link" _
            & " From (ccSections s" _
            & " left join ccTemplates t on s.TemplateID = t.ID)" _
            & " Where (s.active <> 0)" _
            & " and((s.HideMenu is null)OR(s.HideMenu=0))" _
            & " order by " & SortMethod _
            & ""
            '& " order by s.SortOrder"
'        SQL = "Select r.sitemapid as blocked,s.BlockSection, s.Name,s.RootPageID, s.ID As SectionID,s.Caption,s.contentID, s.TemplateID, t.link" _
'                & " From ((ccSections s" _
'                & " left join ccTemplates t on s.TemplateID = t.ID)" _
'                & " left join siteMapSectionRules r on r.sectionId=s.id)" _
'                & " Where " _
'                    & " (s.Active<>0)" _
'                    & " and((s.HideMenu is null)OR(s.HideMenu=0))" _
'                    & " and((r.id is null)or((r.sitemapid=" & siteMapId & ")and(r.active<>0)))" _
'                & " " & SortMethod
    End If
    'If IsOldNameMatchMode Then
    '    '
    '    ' sections - pages connection made with section name
    '    '
    '    SQL = "Select s.BlockSection, s.Name, s.ID As SectionID,s.Caption,s.contentID, s.TemplateID, t.link, 0 as RootPageID" _
    '            & " From ccSections s" _
    '            & " LEFT OUTER JOIN ccTemplates t ON s.TemplateID = t.ID" _
    '            & " Where ((s.HideMenu is null)OR(s.HideMenu=0))And(s.Active<>0)And(s.ContentControlID=" & Main.GetContentID("site Sections") & ")" _
    '            & SortMethod
    'Else
    '    '
    '    ' sections - pages connection made with section RootPageID
    '    '
    '    SQL = "Select s.BlockSection, s.Name,s.RootPageID, s.ID As SectionID,s.Caption,s.contentID, s.TemplateID, t.link" _
    '            & " From ccSections s LEFT OUTER JOIN ccTemplates t ON s.TemplateID = t.ID" _
    '            & " Where ((s.HideMenu is null)OR(s.HideMenu=0))And(s.Active<>0)And(s.ContentControlID=" & Main.GetContentID("site Sections") & ")" _
    '            & SortMethod
    'End If
    CSSectionPointer = Main.OpenCSSQL("Default", SQL)
    If Main.CSOK(CSSectionPointer) Then
        '
        Stream = ""
        PageContentCID = Main.GetContentID("Page Content")
        Do While Main.CSOK(CSSectionPointer)
            '
            If Not Main.GetCSBoolean(CSSectionPointer, "blocked") Then
                ContentID = Main.GetCSInteger(CSSectionPointer, "ContentID")
                BlockSection = Main.GetCSBoolean(CSSectionPointer, "BlockSection")
                SectionID = Main.GetCSInteger(CSSectionPointer, "SectionID")
                urlLink = Main.GetCSText(CSSectionPointer, "link")
                Caption = Main.GetCSText(CSSectionPointer, "Caption")
                If ContentID <= 0 Then
                    ContentID = PageContentCID
                End If
                If ContentID = PageContentCID Then
                    ' just saves a step
                    ContentName = "Page Content"
                Else
                    ContentName = Main.GetContentNameByID(ContentID)
                End If
                If ContentName = "" Then
                    '
                    ' It is possible for the section to be set with a content that has been deleted
                    '
                    ContentName = "Page Content"
                    ContentID = PageContentCID
                End If
                If IsOldNameMatchMode Then
                    RootPageID = Main.GetRecordID(ContentName, Main.GetCSText(CSSectionPointer, "Name"))
                Else
                    RootPageID = Main.GetCSInteger(CSSectionPointer, "RootPageID")
                End If
                If RootPageID <> 0 Then
                    If Not Main.IsSectionBlocked(SectionID, BlockSection) Then
                        'RootPageName = Main.EncodeContent(RootPageName, Main.memberID, , False, , True, True, , True)
                        Caption = Main.EncodeContent(Caption, Main.memberID, , False, , True, True, , True)
                        Caption = Main.ConvertHTML2Text(Caption)
                        Caption = Replace(Caption, vbCrLf, "")
                        '
                        If (indexPageName <> "") Then
                            Stream = Stream & GetPageTree(indexPageName, ContentName, RootPageID, Caption, 2, SectionID, isEditing, siteMapId)
                        Else
                            Stream = Stream & GetPageTree(urlLink, ContentName, RootPageID, Caption, 2, SectionID, isEditing, siteMapId)
                        End If
                    End If
                End If
            End If
            Main.NextCSRecord (CSSectionPointer)
        Loop
       '
    End If
    '
    GetSectionBasedSiteMap = Stream
    '
    Call Main.CloseCS(CSSectionPointer)
    '
    Exit Function
ErrorTrap:
    Call HandleError("SiteMapClass", "GetPageTree", Err.Number, Err.Source, Err.Description, True, False)
    End Function




