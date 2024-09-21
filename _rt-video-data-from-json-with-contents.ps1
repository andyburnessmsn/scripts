# path to the folder containing rooster teeth downloads
$rootFolder = "\\TX100-UNRAID\share1\jdownloader"

# path to the output HTML file
$strHTMLfile = "$rootFolder\_rt-archive-URLs\test.html"

# create file system object
$objFSO = New-Object -ComObject Scripting.FileSystemObject

# array to hold the data
$data = @()

# get all subfolders in the root folder where the subfolder name starts with roosterteeth
write-host "Getting folders..."
$subFolders = get-childitem -Path $rootFolder -Directory -Name roosterteeth*
$subfoldersCount = $subfolders.Count
write-host "Done - found $subfoldersCount folders."

# loop through each subfolder
$i = 0
foreach ($subfolderName in $subFolders){
    $i++

    # get full path of the subfolder
    $subfolderPath = ""
    $subfolderPath = "$rootFolder\$subfolderName"

    ##$subfolderPath = "\\TX100-UNRAID\share1\jdownloader\roosterteeth-47517"
    
    # get all json files in the subfolder (there should only be 1)
    $jsonFiles = ""
    $jsonFiles = get-childitem -Path $subfolderPath -Filter *.json

    # set the path of the first json file in the subfolder (there should only be 1)
    $jsonFilePath = ""
    $jsonFilePath = $jsonFiles[0].FullName

    # read the json file data to string then convert the string to a json object (using objFSO for this as native powershell get-content wasn't happy with it for some reason)
    $objFile = ""
    $objFile = $objFSO.OpenTextFile($jsonFilePath,1)
    $strText = ""
    $strText = $objFile.ReadAll()
    $objFile.Close()

    $jsonData = ""
    $jsonData = $strText | ConvertFrom-Json

    # setup variables 
    $seriesName = ""
    $episodeNumber = ""
    $episodeName = ""

    # read variable data from the json data
    $seriesName = $jsonData.series
    $seasonName = $jsonData.season
    $episodeNumber = $jsonData.episode_number
    $episodeName = $jsonData.episode
    $episodeDescription = $jsondata.description
   
    # create video name full variable - to be used for the "DiplayName" in the HTML table
    $strVideoNameFull = ""
    $strVideoNameFull = "$seriesName - $seasonName - $episodeNumber - $episodeName"

    # output to the console which video we are on
    write-host "$i / $subfoldersCount - $strVideoNameFull"

    # get the path to the first mp4 file in the subfolder (there should only be 1)
    $mp4Files = ""
    $mp4Files = get-childitem -Path $subfolderPath -Filter *.mp4
    $mp4file = ""
    $mp4file = $mp4Files[0]

    # create a HTML link variable to the mp4 file
    $link = ""
    $link = '<a href="file:' + $mp4file.FullName.replace("\","/") + '" target=_blank>Link</a>'

    # create a new object to be added to the "data" array
    $obj = New-Object -TypeName psobject;
    $obj | add-member Link $link
    $obj | add-member Series $seriesName
    $obj | add-member Season $seasonName
    $obj | Add-Member EpisodeNumber $episodeNumber 
    $obj | Add-Member EpisodeName $episodeName
    $obj | add-member Description $episodeDescription
    ##$obj | Add-Member DisplayName $strVideoNameFull

    # add the object to the array
    $data += $obj


}

# group the array data by Series (which is a complete show)
$arrSeriesData = $data | group Series


## $arrSeriesData[0].Group[0]
## $data = import-csv "$rootFolder\_rt-archive-URLs\test.csv" -Delimiter ";"
## $data[0]

# Create variable for the HTML data
$strHTML = ""

# create variable for the HTML contents menu - will be updated below when we loop through the Series data.
# the HTML contents menu will contain links to all Series (shows) and their Seasons. the links will go to the list of episodes within that Season.
$strHTMLcontents = "<ul>"

# create a variable containing an HTML table of the data array, grouped by Series (a complete show), sorted by name, with a count of episodes.
$strHTMLtableGroup = $data | group Series | sort Name | select Name,Count | ConvertTo-Html -Fragment
$strHTML += "$strHTMLtableGroup <hr/>"
$strLinkTop = '<a href=#top>Back to top</a><br/>'

# loop through each Series (a complete show) in the data array
foreach ($series in $($data | group Series | sort Name)){
    # get all Seasons in the Series (each Series contains multiple Seasons, each Season contains Episodes)
    $seasons = $series.Group | group Season | sort Name

    # loop through each Season
    foreach ($season in $seasons){
        # set the "display name" of the link - for example "Rage Quit - Season 1"
        $strSeriesSeason = "$($series.Name) - $($season.Name)"
        # replace any characters which could break the HTML formatting, remove spaces, etc
        $strSeriesSeason = $strSeriesSeason.replace("'","")
        $strSeriesSeason = $strSeriesSeason.replace("-","")

        # set an ID value for the jump link (HTML anchor)
        $idValue = $strSeriesSeason.replace(" ","")
        # update the HTML contents string with an HTML list item containing a link which will go to the Season list of episodes.
        $strHTMLcontents = $strHTMLcontents + '<li><a href="#' + $idValue + '">' + $strSeriesSeason + '(' + $season.Count + ')</a></li>'

        ## update the main HTML with an ID value (used for jump link from contents menu), a new heading of the "Series - Season" name (example "Rage Quit - Season 1"), and an HTML table with the list of episodes in the season.
        $strHTML += '<h2 id="' + $idValue + '">' + $strSeriesSeason + '</h2>' + $strLinkTop + $($season.group | sort EpisodeNumber | convertto-html -Fragment) + '<hr>'
    }
}

# finish the HTML contents string.
$strHTMLcontents += "</ul>"

# add a title to the top of the main HTML. add an "last updated" timestamp.
$timestamp = get-date
$strHeaderBootstrap = '<!-- Latest compiled and minified CSS -->
<link rel="stylesheet" href="https://maxcdn.bootstrapcdn.com/bootstrap/3.4.1/css/bootstrap.min.css">

<!-- jQuery library -->
<script src="https://ajax.googleapis.com/ajax/libs/jquery/3.7.1/jquery.min.js"></script>

<!-- Latest compiled JavaScript -->
<script src="https://maxcdn.bootstrapcdn.com/bootstrap/3.4.1/js/bootstrap.min.js"></script>'

$strHeaderStyle = @"
<style>
BODY {
  FONT-SIZE: 11px; MARGIN: 0px; FONT-FAMILY: Arial,Helvetica,sans-serif
}
BIG {
  FONT-SIZE: 14px
}
BLOCKQUOTE {
  FONT-FAMILY: Arial, Helvetica, sans-serif
}
PRE {
  FONT-FAMILY: Arial, Helvetica, sans-serif
}
DT {
  FONT-WEIGHT: bold; COLOR: #cc3333
}
H1 {
  FONT-WEIGHT: bold; FONT-SIZE: 18px; COLOR: #000000
}
H2 {
  FONT-WEIGHT: bold; FONT-SIZE: 16px; COLOR: #000000
}
H3 {
  FONT-WEIGHT: normal; FONT-SIZE: 14px; COLOR: #000000
}
H4 {
  FONT-WEIGHT: normal; FONT-SIZE: 12px; COLOR: #000000
}
H5 {
  FONT-WEIGHT: bold; FONT-SIZE: 11px; COLOR: #000000
}
H6 {
  FONT-WEIGHT: bold; FONT-SIZE: 10px; COLOR: #000000
}
HR {
  HEIGHT: 1pt
}
OL LI {
  COLOR: #636363; LIST-STYLE-TYPE: decimal
}
OL OL LI {
  LIST-STYLE-TYPE: lower-alpha
}
OL OL OL LI {
  LIST-STYLE-TYPE: lower-roman
}
P {
  COLOR: #636363
}
INPUT {
  COLOR: #636363
}
OPTION {
  COLOR: #636363
}
SELECT {
  COLOR: #636363
}
TEXTAREA {
  COLOR: #636363
}
TR {
  COLOR: #636363
}
TD {
  FONT-SIZE: 12px
}
FONT {
  FONT-SIZE: 12px
}
SMALL {
  FONT-SIZE: 10px
}
TH {
  PADDING-RIGHT: 10px;
    PADDING-LEFT: 10px;
    FONT-WEIGHT: bold;
    FONT-SIZE: 12px;
    VERTICAL-ALIGN: middle;
    COLOR: #ffffff;
    FONT-STYLE: normal;
    FONT-FAMILY: Arial, sans-serif;
    HEIGHT: 20px;
    BACKGROUND-COLOR: #4f7ca0;
    TEXT-ALIGN: left;
    cursor: pointer;
    position: sticky;
    top: 0;
}
THEAD {
  COLOR: #cc3333
}
TFOOT {
  COLOR: #cc3333
}
UL {
  COLOR: #636363
}
UL LI LI {
  LIST-STYLE-TYPE: disc
}
UL LI LI LI {
  LIST-STYLE-TYPE: circle
}
#rightpane {
  PADDING-RIGHT: 1px; PADDING-LEFT: 0px; PADDING-BOTTOM: 0px; COLOR: #666666; LINE-HEIGHT: 15px; MARGIN-RIGHT: 0px; PADDING-TOP: 0px
}
#leftpane {
  PADDING-RIGHT: 10px; PADDING-LEFT: 12px; LINE-HEIGHT: 15px
}
#leftpane .navbar {
  FONT-SIZE: 11px; LINE-HEIGHT: 12px; BORDER-BOTTOM: transparent 0px dashed
}
#leftpane .navbar TD {
  FONT-SIZE: 11px; BORDER-BOTTOM: #b7b7b7 1px dashed
}
#leftpane .none TD {
  FONT-SIZE: 11px; LINE-HEIGHT: 12px; BORDER-BOTTOM: transparent 0px dashed
}
#leftpane .navbar TD TD {
  FONT-SIZE: 11px; LINE-HEIGHT: 12px; BORDER-BOTTOM: transparent 0px dashed
}
#leftpane .navbar TD TD TD {
  FONT-SIZE: 11px; LINE-HEIGHT: 12px; BORDER-BOTTOM: transparent 0px dashed
}
#leftpane .navbar .navSelected {
  FONT-WEIGHT: bold; FONT-SIZE: 11px; COLOR: #666666; LINE-HEIGHT: 12px
}
#leftpane .navbar .navSelected {
  FONT-WEIGHT: bold; FONT-SIZE: 11px; COLOR: #666666; LINE-HEIGHT: 12px
}
#leftpane .NavBarBlueBg TD TD {
  COLOR: #666666
}
#leftpane .NavBarBlueBg .navSelected {
  FONT-WEIGHT: bold; FONT-SIZE: 11px; COLOR: #666666; LINE-HEIGHT: 12px
}
#leftpane .n {
  FONT-WEIGHT: bold; FONT-SIZE: 11px; LINE-HEIGHT: 12px
}
#rightpane TD {
  PADDING-LEFT: 5px; FONT-SIZE: 11px; COLOR: #666666; LINE-HEIGHT: 16px; BACKGROUND-COLOR: transparent
}
#rightpane TR {
  FONT-SIZE: 11px; COLOR: #666666; LINE-HEIGHT: 16px; BACKGROUND-COLOR: transparent
}
#rightpane TBODY {
  BACKGROUND-COLOR: transparent
}
#rightpane TABLE {
  BACKGROUND-COLOR: #e2eeee
}
#rightpane SPAN {
  BACKGROUND-COLOR: transparent
}
#rightpane PANEL {
  BACKGROUND-COLOR: transparent
}
#rightpane DIV {
  BORDER-RIGHT: #8b9999 1px solid; BORDER-BOTTOM: #8b9999 1px solid; BACKGROUND-COLOR: #e2eeee
}
.BannerBg {
  BACKGROUND-COLOR: #4cb7b7
}
.bdy {
  FONT-SIZE: 12px; COLOR: #666666; LINE-HEIGHT: 16px; FONT-FAMILY: Arial,Helvetica,sans-serif
}
.bodytext {
  FONT-SIZE: 12px; COLOR: #666666; LINE-HEIGHT: 16px; FONT-FAMILY: Arial,Helvetica,sans-serif
}
.defaultlabletext {
  FONT-SIZE: 12px; COLOR: #666666; LINE-HEIGHT: 16px; FONT-FAMILY: Arial,Helvetica,sans-serif
}
.instruction {
  FONT-SIZE: 12px; COLOR: #666666; LINE-HEIGHT: 16px; FONT-FAMILY: Arial,Helvetica,sans-serif
}
.newsarticletext {
  FONT-SIZE: 12px; COLOR: #666666; LINE-HEIGHT: 16px; FONT-FAMILY: Arial,Helvetica,sans-serif
}
.Normal {
  FONT-SIZE: 12px; COLOR: #666666; LINE-HEIGHT: 16px; FONT-FAMILY: Arial,Helvetica,sans-serif
}
.NormalTextBox {
  FONT-SIZE: 12px; COLOR: #666666; LINE-HEIGHT: 16px; FONT-FAMILY: Arial,Helvetica,sans-serif
}
.NormalText {
  FONT-SIZE: 12px; COLOR: #666666; LINE-HEIGHT: 16px; FONT-FAMILY: Arial,Helvetica,sans-serif
}
.textbox {
  FONT-SIZE: 12px; COLOR: #666666; LINE-HEIGHT: 16px; FONT-FAMILY: Arial,Helvetica,sans-serif
}
#rightpane .bdy {
  FONT-SIZE: 11px; BACKGROUND-COLOR: transparent
}
#rightpane .bodytext {
  FONT-SIZE: 11px; BACKGROUND-COLOR: transparent
}
#rightpane .defaultlabletext {
  FONT-SIZE: 11px; BACKGROUND-COLOR: transparent
}
#rightpane .instruction {
  FONT-SIZE: 11px; BACKGROUND-COLOR: transparent
}
#rightpane .newsarticletext {
  FONT-SIZE: 11px; BACKGROUND-COLOR: transparent
}
#rightpane .NormalText {
  FONT-SIZE: 11px; BACKGROUND-COLOR: transparent
}
#rightpane .Normal {
  FONT-SIZE: 11px; BACKGROUND-COLOR: transparent
}
#rightpane .NormalTextBox {
  FONT-SIZE: 11px; BACKGROUND-COLOR: transparent
}
#rightpane .textbox {
  FONT-SIZE: 11px; BACKGROUND-COLOR: transparent
}
.CommandButton {
  FONT-WEIGHT: normal
}
.EditThisPage {
  PADDING-LEFT: 5px; FONT-WEIGHT: bold; FONT-SIZE: 11px; COLOR: #666666
}
.EditThisPage:link {
  PADDING-LEFT: 10px; FONT-WEIGHT: bold; FONT-SIZE: 11px; COLOR: #666666; TEXT-DECORATION: none
}
.EditThisPage:active {
  PADDING-LEFT: 10px; FONT-WEIGHT: bold; FONT-SIZE: 11px; COLOR: #666666; TEXT-DECORATION: none
}
.EditThisPage:visited {
  PADDING-LEFT: 10px; FONT-WEIGHT: bold; FONT-SIZE: 11px; COLOR: #666666; TEXT-DECORATION: none
}
.EditThisPage:hover {
  TEXT-DECORATION: underline
}
.GreyNavBg {
  PADDING-RIGHT: 5px; PADDING-LEFT: 0px; PADDING-BOTTOM: 0px; VERTICAL-ALIGN: text-bottom; PADDING-TOP: 0px; BORDER-BOTTOM: #7eaaba 1px solid; BACKGROUND-COLOR: #bfe5e5
}
.HeadBg {
  PADDING-RIGHT: 5px; PADDING-LEFT: 0px; PADDING-BOTTOM: 0px; VERTICAL-ALIGN: text-bottom; PADDING-TOP: 0px; BORDER-BOTTOM: #7eaaba 1px solid; BACKGROUND-COLOR: #bfe5e5
}
.Header {
  VERTICAL-ALIGN: text-bottom; BORDER-BOTTOM: #7eaaba 1px solid; BACKGROUND-COLOR: #bfe5e5
}
.GreyNavSelected {
  BORDER-RIGHT: #7eaaba 1px solid; BORDER-TOP: #7eaaba 1px solid; FONT-WEIGHT: bold; FONT-SIZE: 11px; TEXT-TRANSFORM: capitalize; BORDER-LEFT: #7eaaba 1px solid; COLOR: #47869d; BORDER-BOTTOM: medium none; BACKGROUND-COLOR: #e5f5f5; TEXT-DECORATION: none
}
.SubNav {
  BORDER-BOTTOM: #c7dae1 1px solid; BACKGROUND-COLOR: #e5f5f5
}
.TextBoxBanner {
  BORDER-RIGHT: #4e4e4e 1px groove; BORDER-TOP: #4e4e4e 1px groove; FONT-SIZE: 11px; VERTICAL-ALIGN: middle; BORDER-LEFT: #4e4e4e 1px groove; COLOR: #666666; LINE-HEIGHT: 12px; BORDER-BOTTOM: #4e4e4e 1px groove; BACKGROUND-COLOR: #91c6c6
}
.Message {
  FONT-WEIGHT: normal; BACKGROUND-COLOR: #d6e5ea
}
.OtherTabs {
  PADDING-RIGHT: 3px; PADDING-LEFT: 1px; FONT-SIZE: 11px; PADDING-BOTTOM: 1px; TEXT-TRANSFORM: uppercase; COLOR: #ffffff; PADDING-TOP: 1px; TEXT-ALIGN: center; TEXT-DECORATION: none
}
.OtherTabs_u {
  PADDING-RIGHT: 3px; PADDING-LEFT: 1px; FONT-SIZE: 11px; PADDING-BOTTOM: 1px; TEXT-TRANSFORM: uppercase; COLOR: #ffffff; PADDING-TOP: 1px; TEXT-ALIGN: center; TEXT-DECORATION: none
}
.OtherTabs:link {
  TEXT-DECORATION: none
}
.OtherTabs_u:link {
  TEXT-DECORATION: none
}
.OtherTabs:visited {
  TEXT-DECORATION: none
}
.OtherTabs_u:visited {
  TEXT-DECORATION: none
}
A.SelectedTab:link {
  TEXT-DECORATION: none
}
A.SelectedTab:visited {
  TEXT-DECORATION: none
}
A.SelectedTab_u:link {
  TEXT-DECORATION: none
}
A.SelectedTab_u:visited {
  TEXT-DECORATION: none
}
.OtherTabsBg {
  BORDER-LEFT-COLOR: #636363; BORDER-BOTTOM-COLOR: #636363; VERTICAL-ALIGN: middle; BORDER-TOP-COLOR: #636363; BACKGROUND-COLOR: #636363; BORDER-RIGHT-COLOR: #636363
}
.SelectedTab {
  BORDER-RIGHT: #7eaaba 1px solid;
  PADDING-RIGHT: 6px;
  BORDER-TOP: #7eaaba 1px solid;
  PADDING-LEFT: 6px;
  FONT-WEIGHT: bold;
  FONT-SIZE: 1, 2;
  MARGIN-BOTTOM: -1px;
  PADDING-BOTTOM: 3px;
  VERTICAL-ALIGN: middle;
  BORDER-LEFT: #7eaaba 1px solid;
  COLOR: #FFFFFF;
  PADDING-TOP: 1px;
  BORDER-BOTTOM: medium none;
  BACKGROUND-COLOR: #969696;
  TEXT-ALIGN: center;
  TEXT-DECORATION: none;
  font-style: normal;
}
.SelectedTab_u {
  BORDER-RIGHT: #c7dae1 1px solid; PADDING-RIGHT: 5px; BORDER-TOP: #c7dae1 1px solid; MARGIN-TOP: 1px; PADDING-LEFT: 5px; FONT-WEIGHT: bold; FONT-SIZE: 10px; FLOAT: left; MARGIN-BOTTOM: -3px; PADDING-BOTTOM: 2px; VERTICAL-ALIGN: middle; TEXT-TRANSFORM: uppercase; BORDER-LEFT: #c7dae1 1px solid; COLOR: #47869d; PADDING-TOP: 3px; BORDER-BOTTOM: #c7dae1 1px; POSITION: relative; BACKGROUND-COLOR: #ffffff; TEXT-ALIGN: center
}
.SelectedTab_Text {
  FONT-WEIGHT: bold; FONT-SIZE: 11px; VERTICAL-ALIGN: middle; TEXT-TRANSFORM: uppercase; COLOR: #47869d; BORDER-TOP-STYLE: none; BORDER-RIGHT-STYLE: none; BORDER-LEFT-STYLE: none; TEXT-ALIGN: center; TEXT-DECORATION: none; BORDER-BOTTOM-STYLE: none
}
.SiteTitle {
  FONT-WEIGHT: bold; FONT-SIZE: 20px; COLOR: #000000
}
.TitleDate {
  PADDING-RIGHT: 5px; FONT-WEIGHT: bold; FONT-SIZE: x-small; WIDTH: 200px; COLOR: #ff3333; BACKGROUND-COLOR: transparent; TEXT-ALIGN: right
}
.TitleName {
  FONT-WEIGHT: bold; FONT-SIZE: small; WIDTH: 200px; COLOR: #009999; TEXT-ALIGN: left
}
.EditModuleHead {
  FONT-WEIGHT: normal; COLOR: #47869d; BACKGROUND-COLOR: #ffffff; TEXT-ALIGN: left
}
.ModuleHead {
  PADDING-RIGHT: 20px; PADDING-LEFT: 0px; FONT-WEIGHT: normal; FONT-SIZE: 1px; COLOR: #ffffff; BACKGROUND-COLOR: #d6e5ea
}
.ModuleBg {
  PADDING-RIGHT: 20px; PADDING-LEFT: 0px; FONT-WEIGHT: normal; FONT-SIZE: 1px; COLOR: #ffffff; BACKGROUND-COLOR: #d6e5ea
}
.ModuleText {
  PADDING-RIGHT: 5px; PADDING-LEFT: 0px; FONT-WEIGHT: bold; FONT-SIZE: 14px; PADDING-BOTTOM: 1px; VERTICAL-ALIGN: baseline; TEXT-TRANSFORM: uppercase; COLOR: #000000; PADDING-TOP: 1px; FONT-FAMILY: "Zurich", "Arial", sans-serif; BACKGROUND-COLOR: #ffffff
}
.ModuleButtons {
  PADDING-RIGHT: 5px; FONT-WEIGHT: normal; FONT-SIZE: 14px; VERTICAL-ALIGN: baseline; COLOR: #47869d; BACKGROUND-COLOR: #ffffff
}
.ModuleBtnBg {
  PADDING-RIGHT: 5px; FONT-WEIGHT: normal; VERTICAL-ALIGN: top; COLOR: #47869d; BACKGROUND-COLOR: #ffffff
}
.ModuleButtons {
  PADDING-RIGHT: 5px; FONT-WEIGHT: normal; VERTICAL-ALIGN: top; COLOR: #47869d; BACKGROUND-COLOR: #ffffff
}
.TabLayoutHead {
  PADDING-RIGHT: 5px; PADDING-LEFT: 0px; FONT-WEIGHT: normal; FONT-SIZE: 12px; COLOR: #000000; BACKGROUND-COLOR: #d6e5ea
}
#rightpane .ModuleHead {
  PADDING-RIGHT: 0px; PADDING-LEFT: 4px; FONT-WEIGHT: normal; FONT-SIZE: 11px; PADDING-BOTTOM: 2px; COLOR: #ffffff; PADDING-TOP: 2px; BACKGROUND-COLOR: #c1dbd8; TEXT-ALIGN: left
}
#rightpane .ModuleBg {
  FONT-WEIGHT: normal; FONT-SIZE: 11px; COLOR: #ffffff; BORDER-BOTTOM: #8b9999 1px groove; BACKGROUND-COLOR: #c1dbd8; TEXT-ALIGN: left
}
#rightpane .ModuleText {
  PADDING-RIGHT: 5px; FONT-WEIGHT: bold; FONT-SIZE: 11px; VERTICAL-ALIGN: middle; TEXT-TRANSFORM: none; COLOR: #526262; BACKGROUND-COLOR: #c1dbd8
}
#rightpane .ModuleButtons {
  PADDING-RIGHT: 5px; FONT-WEIGHT: normal; FONT-SIZE: 11px; VERTICAL-ALIGN: top; COLOR: #47869d; BACKGROUND-COLOR: #c1dbd8
}
#rightpane .ModuleBtnBg {
  PADDING-RIGHT: 5px; FONT-WEIGHT: normal; VERTICAL-ALIGN: top; COLOR: #47869d; PADDING-TOP: 5px; BACKGROUND-COLOR: #c1dbd8
}
.footerBG {
  BORDER-TOP: #e2eeee 5px solid
}
.footer {
  FONT-SIZE: 10px; PADDING-BOTTOM: 3px; COLOR: #666666; PADDING-TOP: 3px
}
.footer A:link {
  TEXT-DECORATION: none
}
.footer A:visited {
  TEXT-DECORATION: none
}
.monospace {
  FONT-WEIGHT: normal; FONT-SIZE: 11px; COLOR: #636363; FONT-STYLE: normal; FONT-FAMILY: "Courier New", Courier, monospace
}
.BannerTitle {
  FONT-SIZE: 21px; COLOR: #4db8b8; FONT-FAMILY: Arial, "SansSerif"
}
.leftPane {
  BORDER-RIGHT: #c7dae1 1px dashed
}
#leftpane .Arrow {
  COLOR: #41869d
}
A:link {
  FONT-WEIGHT: normal; FONT-SIZE: 12px; COLOR: #0072bc; LINE-HEIGHT: 16px; TEXT-DECORATION: underline
}
A:visited {
  FONT-WEIGHT: normal; FONT-SIZE: 12px; COLOR: #0072bc; LINE-HEIGHT: 16px; TEXT-DECORATION: underline
}
A:hover {
  TEXT-DECORATION: underline
}
A:active {
  COLOR: #336675; TEXT-DECORATION: underline
}
A.CommandButton:link {
  FONT-WEIGHT: normal; FONT-SIZE: 12px; COLOR: #0072bc; LINE-HEIGHT: 16px; TEXT-DECORATION: underline
}
A.CommandButton:active {
  FONT-WEIGHT: normal; FONT-SIZE: 12px; COLOR: #0072bc; LINE-HEIGHT: 16px; TEXT-DECORATION: underline
}
A.CommandButton:visited {
  FONT-WEIGHT: normal; FONT-SIZE: 12px; COLOR: #0072bc; LINE-HEIGHT: 16px; TEXT-DECORATION: underline
}
A.CommandButton:hover {
  FONT-WEIGHT: normal; FONT-SIZE: 12px; COLOR: #b7b7b7; LINE-HEIGHT: 16px; TEXT-DECORATION: underline
}
A.h:link {
  FONT-SIZE: 12px; COLOR: #47869d; LINE-HEIGHT: 16px; TEXT-DECORATION: underline
}
A.h:active {
  FONT-SIZE: 12px; COLOR: #47869d; LINE-HEIGHT: 16px; TEXT-DECORATION: underline
}
A.h:visited {
  FONT-SIZE: 12px; COLOR: #47869d; LINE-HEIGHT: 16px; TEXT-DECORATION: underline
}
A.h:hover {
  FONT-SIZE: 12px; COLOR: #000000; TEXT-DECORATION: underline
}
A.n:active {
  FONT-WEIGHT: normal; FONT-SIZE: 12px; COLOR: #41869d
}
A.n:link {
  FONT-WEIGHT: normal; FONT-SIZE: 12px; COLOR: #41869d
}
A.n:visited {
  FONT-WEIGHT: normal; FONT-SIZE: 12px; COLOR: #41869d
}
A.p:active {
  FONT-WEIGHT: normal; FONT-SIZE: 12px; COLOR: #41869d
}
A.p:link {
  FONT-WEIGHT: normal; FONT-SIZE: 12px; COLOR: #41869d
}
A.p:visited {
  FONT-WEIGHT: normal; FONT-SIZE: 12px; COLOR: #41869d
}
A.n:hover {
  FONT-WEIGHT: normal; FONT-SIZE: 12px; COLOR: #000000; TEXT-DECORATION: underline
}
A.p:hover {
  FONT-WEIGHT: normal; FONT-SIZE: 12px; COLOR: #000000; TEXT-DECORATION: underline
}
A.m:active {
  FONT-SIZE: 11px; COLOR: #959595
}
A.m:link {
  FONT-SIZE: 11px; COLOR: #959595
}
A.m:visited {
  FONT-SIZE: 11px; COLOR: #959595
}
A.m:hover {
  FONT-SIZE: 11px; COLOR: #000000
}
A.OtherTabs:link {
  FONT-SIZE: 10px; COLOR: #47869d; TEXT-DECORATION: none
}
A.OtherTabs:active {
  FONT-SIZE: 10px; COLOR: #47869d; TEXT-DECORATION: none
}
A.OtherTabs_u:link {
  FONT-SIZE: 10px; COLOR: #47869d; TEXT-DECORATION: none
}
A.OtherTabs_u:active {
  FONT-SIZE: 10px; COLOR: #47869d; TEXT-DECORATION: none
}
A.OtherTabs:visited {
  FONT-SIZE: 10px; COLOR: #47869d; TEXT-DECORATION: none
}
A.OtherTabs_u:visited {
  FONT-SIZE: 10px; COLOR: #47869d; TEXT-DECORATION: none
}
A.OtherTabs:hover {
  FONT-SIZE: 10px; COLOR: #000000; TEXT-DECORATION: underline
}
A.OtherTabs_u:hover {
  FONT-SIZE: 10px; COLOR: #000000; TEXT-DECORATION: underline
}
A.s:link {
  FONT-SIZE: 12px; COLOR: #959595
}
A.s:active {
  FONT-SIZE: 12px; COLOR: #959595
}
A.s:visited {
  FONT-SIZE: 12px; COLOR: #959595
}
A.s:hover {
  FONT-SIZE: 12px; COLOR: #000000; TEXT-DECORATION: underline
}
.SiteLink {
  FONT-WEIGHT: bold; FONT-SIZE: 12px
}
A.SiteLink:link {
  COLOR: #636363
}
A.SiteLink:visited {
  COLOR: #636363
}
A.SiteLink:active {
  COLOR: #636363
}
A.SiteLink:hover {
  COLOR: #000000; TEXT-DECORATION: underline
}
A.whiteLink:link {
  TEXT-DECORATION: none
}
A.whiteLink:visited {
  TEXT-DECORATION: none
}
A.whiteLink:active {
  TEXT-DECORATION: none
}
A.whiteLink:hover {
  TEXT-DECORATION: none
}
A.whiteBannerLink:link {
  FONT-SIZE: 10px; COLOR: #ffffff
}
A.whiteBannerLink:visited {
  FONT-SIZE: 10px; COLOR: #ffffff
}
A.whiteBannerLink:active {
  FONT-SIZE: 10px; COLOR: #ffffff
}
A.whiteBannerLink:hover {
  FONT-SIZE: 10px; COLOR: #ffffff; TEXT-DECORATION: underline
}
A.bread:link {
  FONT-WEIGHT: bold; FONT-SIZE: 11px; COLOR: #ffffff
}
A.bread:visited {
  FONT-WEIGHT: bold; FONT-SIZE: 11px; COLOR: #ffffff
}
A.bread:active {
  FONT-WEIGHT: bold; FONT-SIZE: 11px; COLOR: #ffffff
}
A.whiteBannerLink:hover {
  FONT-SIZE: 11px; COLOR: #ffffff; TEXT-DECORATION: underline
}
.newsarticletextlink:link {
  FONT-SIZE: 12px; COLOR: #315b6a
}
.newsarticletextlink:visited {
  FONT-SIZE: 12px; COLOR: #315b6a
}
.newsarticletextlink:active {
  FONT-SIZE: 12px; COLOR: #315b6a
}
.newsarticletextlink:hover {
  FONT-SIZE: 12px; COLOR: #000000; TEXT-DECORATION: underline
}
.newsdatelink {
  FONT-SIZE: 12px; COLOR: #009999
}
.pop:link {
  FONT-WEIGHT: normal; FONT-SIZE: 12px; COLOR: #47869d; LINE-HEIGHT: 16px; TEXT-DECORATION: none
}
.pop:visited {
  FONT-WEIGHT: normal; FONT-SIZE: 12px; COLOR: #47869d; LINE-HEIGHT: 16px; TEXT-DECORATION: none
}
.pop:hover {
  TEXT-DECORATION: underline
}
.pop:active {
  COLOR: #336675
}
#leftpane A:link {
  TEXT-DECORATION: none
}
#leftpane A:visited {
  TEXT-DECORATION: none
}
#leftpane A:hover {
  TEXT-DECORATION: underline
}
#rightpane .newsarticletextlink:link {
  FONT-SIZE: 11px; COLOR: #315b6a; TEXT-DECORATION: underline
}
#rightpane .newsarticletextlink:visited {
  FONT-SIZE: 11px; COLOR: #315b6a; TEXT-DECORATION: underline
}
#rightpane .newsarticletextlink:active {
  FONT-SIZE: 11px; COLOR: #315b6a; TEXT-DECORATION: underline
}
#rightpane .newsarticletextlink:hover {
  FONT-SIZE: 11px; COLOR: #000000; TEXT-DECORATION: underline
}
#rightpane A:link {
  FONT-SIZE: 11px; COLOR: #41869d; TEXT-DECORATION: underline
}
#rightpane A:active {
  FONT-SIZE: 11px; COLOR: #41869d; TEXT-DECORATION: underline
}
#rightpane A:visited {
  FONT-SIZE: 11px; COLOR: #41869d; TEXT-DECORATION: underline
}
#rightpane A:hover {
  FONT-SIZE: 11px; COLOR: #000000; TEXT-DECORATION: underline
}
#rightpane A.CommandButton:link {
  COLOR: #41869d; TEXT-DECORATION: underline
}
#rightpane A.CommandButton:visited {
  COLOR: #41869d; TEXT-DECORATION: underline
}
#rightpane A.CommandButton:active {
  COLOR: #41869d; TEXT-DECORATION: underline
}
#rightpane A.CommandButton:hover {
  COLOR: #000000; TEXT-DECORATION: underline
}
#rightpane A.h:link {
  COLOR: #41869d; TEXT-DECORATION: underline
}
#rightpane A.h:visited {
  COLOR: #41869d; TEXT-DECORATION: underline
}
#rightpane A.h:active {
  COLOR: #41869d; TEXT-DECORATION: underline
}
#rightpane A.h:hover {
  COLOR: #000000; TEXT-DECORATION: underline
}
#rightpane A.n:active {
  FONT-SIZE: 11px; COLOR: #41869d; TEXT-DECORATION: underline
}
#rightpane A.n:link {
  FONT-SIZE: 11px; COLOR: #41869d; TEXT-DECORATION: underline
}
#rightpane A.n:visited {
  FONT-SIZE: 11px; COLOR: #41869d; TEXT-DECORATION: underline
}
#rightpane A.p:active {
  FONT-SIZE: 11px; COLOR: #41869d; TEXT-DECORATION: underline
}
#rightpane A.p:link {
  FONT-SIZE: 11px; COLOR: #41869d; TEXT-DECORATION: underline
}
#rightpane A.p:visited {
  FONT-SIZE: 11px; COLOR: #41869d; TEXT-DECORATION: underline
}
#rightpane .pop:link {
  FONT-SIZE: 11px; COLOR: #41869d; TEXT-DECORATION: underline
}
#rightpane .pop:visited {
  FONT-SIZE: 11px; COLOR: #41869d; TEXT-DECORATION: underline
}
#rightpane A.n:hover {
  FONT-SIZE: 11px; COLOR: #000000; TEXT-DECORATION: underline
}
#rightpane A.p:hover {
  FONT-SIZE: 11px; COLOR: #000000; TEXT-DECORATION: underline
}
#rightpane A.m:link {
  FONT-SIZE: 11px; COLOR: #41869d
}
#rightpane A.m:active {
  FONT-SIZE: 11px; COLOR: #41869d
}
#rightpane A.m:visited {
  FONT-SIZE: 11px; COLOR: #41869d
}
#rightpane A.m:hover {
  FONT-SIZE: 11px; COLOR: #000000
}
#rightpane A.OtherTabs:link {
  FONT-SIZE: 10px; COLOR: #41869d
}
#rightpane A.OtherTabs:visited {
  FONT-SIZE: 10px; COLOR: #41869d
}
#rightpane A.OtherTabs:active {
  FONT-SIZE: 10px; COLOR: #41869d
}
#rightpane A.OtherTabs_u:link {
  FONT-SIZE: 10px; COLOR: #41869d
}
#rightpane A.OtherTabs_u:visited {
  FONT-SIZE: 10px; COLOR: #41869d
}
#rightpane A.OtherTabs_u:active {
  FONT-SIZE: 10px; COLOR: #41869d
}
#rightpane A.OtherTabs:hover {
  FONT-SIZE: 10px; COLOR: #000000; TEXT-DECORATION: underline
}
#rightpane A.OtherTabs_u:hover {
  FONT-SIZE: 10px; COLOR: #000000; TEXT-DECORATION: underline
}
#rightpane A.s:link {
  FONT-SIZE: 11px; COLOR: #41869d
}
#rightpane A.s:visited {
  FONT-SIZE: 11px; COLOR: #41869d
}
#rightpane A.s:active {
  FONT-SIZE: 11px; COLOR: #41869d
}
#rightpane A.s:hover {
  FONT-SIZE: 11px; COLOR: #000000; TEXT-DECORATION: underline
}
.h1 {
  FONT-WEIGHT: bold; FONT-SIZE: 14px; COLOR: #135b6a
}
.h2 {
  FONT-WEIGHT: bold; FONT-SIZE: 14px; COLOR: #676767
}
.h1green {
  FONT-WEIGHT: bold; FONT-SIZE: 14px; COLOR: #676767
}
.h2green {
  FONT-WEIGHT: bold; FONT-SIZE: 14px; COLOR: #676767
  }
.categoryname {
  FONT-WEIGHT: bold; FONT-SIZE: 14px; COLOR: #135b6a
}
.navisructitle {
  FONT-WEIGHT: bold; FONT-SIZE: 14px; COLOR: #135b6a
}
.h3 {
  FONT-WEIGHT: bold; FONT-SIZE: 12px; COLOR: #135b6a
}
.h3green {
  FONT-WEIGHT: bold; FONT-SIZE: 12px; COLOR: #135b6a
}
.homegreen {
  FONT-WEIGHT: bold; FONT-SIZE: 12px; COLOR: #135b6a
}
.articleheading {
  FONT-WEIGHT: bold; FONT-SIZE: 12px; COLOR: #135b6a
}
.headingtext {
  FONT-WEIGHT: bold; FONT-SIZE: 12px; COLOR: #135b6a
}
.newsheadline {
  FONT-WEIGHT: bold; FONT-SIZE: 12px; COLOR: #135b6a
}
.subcategoryname {
  FONT-WEIGHT: bold; FONT-SIZE: 12px; COLOR: #135b6a
}
.ItemTitle {
  FONT-WEIGHT: bold; FONT-SIZE: 12px; COLOR: #666666
}
.h4 {
  FONT-SIZE: 11px; COLOR: #135b6a
}
.searchnumberofpage {
  FONT-SIZE: 11px; COLOR: #135b6a
}
.h4green {
  FONT-SIZE: 11px; COLOR: #009999
}
.articleheader {
  FONT-SIZE: 11px; COLOR: #009999
}
.bodytextbold {
  FONT-WEIGHT: bold; COLOR: #636363
}
.bdybold {
  FONT-WEIGHT: bold; COLOR: #636363
}
.alerttext {
  FONT-WEIGHT: bold; COLOR: #636363
}
.bodyheading {
  FONT-WEIGHT: bold; COLOR: #636363
}
.SubSubHead {
  FONT-WEIGHT: bold; COLOR: #636363
}
.NormalBold {
  FONT-WEIGHT: bold; COLOR: #636363
}
.bodytextsmall {
  FONT-SIZE: 11px; COLOR: #636363
}
.bdysmall {
  FONT-SIZE: 11px; COLOR: #636363
}
.smallinput {
  FONT-SIZE: 11px; COLOR: #636363
}
.menutext {
  FONT-SIZE: 11px; COLOR: #636363
}
.myfscnet {
  PADDING-RIGHT: 23px; FONT-SIZE: 18px; COLOR: #4db8b8; FONT-FAMILY: "Rotis Sans Serif 55", Sans-Serif
}
.Uppercase {
  FONT-SIZE: 11px; TEXT-TRANSFORM: uppercase
}
#rightpane .bodytextbold {
  FONT-WEIGHT: bold; FONT-SIZE: 11px; COLOR: #636363
}
#rightpane .bdybold {
  FONT-WEIGHT: bold; FONT-SIZE: 11px; COLOR: #636363
}
#rightpane .alerttext {
  FONT-WEIGHT: bold; FONT-SIZE: 11px; COLOR: #636363
}
#rightpane .bodyheading {
  FONT-WEIGHT: bold; FONT-SIZE: 11px; COLOR: #636363
}
#rightpane .SubSubHead {
  FONT-WEIGHT: bold; FONT-SIZE: 11px; COLOR: #636363
}
#rightpane .NormalBold {
  FONT-WEIGHT: bold; FONT-SIZE: 11px; COLOR: #636363
}
#rightpane .bodytextsmall {
  FONT-SIZE: 11px; COLOR: #636363
}
#rightpane .bdysmall {
  FONT-SIZE: 11px; COLOR: #636363
}
#rightpane .smallinput {
  FONT-SIZE: 11px; COLOR: #636363
}
#rightpane .menutext {
  FONT-SIZE: 11px; COLOR: #636363
}
.go {
  COLOR: #636363; TEXT-DECORATION: none
}
#rightpane .NormalBold {
  FONT-WEIGHT: bold; FONT-SIZE: 11px; BACKGROUND-COLOR: transparent
}
.b {
  COLOR: #000000; TEXT-DECORATION: none
}
.bs {
  FONT-SIZE: 11px; COLOR: #000000; TEXT-DECORATION: none
}
.ownertext {
  FONT-SIZE: 11px; COLOR: #000000; TEXT-DECORATION: none
}
.menuheadingtext {
  FONT-SIZE: 11px; COLOR: #000000; TEXT-DECORATION: none
}
.currentpagesmall {
  FONT-SIZE: 11px; COLOR: #000000; TEXT-DECORATION: none
}
.black {
  COLOR: #000000
}
.fo {
  COLOR: #000000
}
.menurollover {
  COLOR: #000000
}
.currentpage {
  COLOR: #000000
}
.blackbold {
  FONT-WEIGHT: bold; COLOR: #000000
}
.green {
  COLOR: #009999
}
.emphasizedtext {
  COLOR: #009999
}
.bodytextgreen {
  COLOR: #009999
}
.confirmlabel {
  COLOR: #009999
}
.sideheadingtext {
  COLOR: #009999
}
.greenbold {
  FONT-WEIGHT: bold; COLOR: #009999
}
.ItemTitle {
  FONT-WEIGHT: bold; FONT-SIZE: 12px; COLOR: #315b6a; LINE-HEIGHT: 16px
}
.orange {
  COLOR: #ff9900
}
.orangebold {
  FONT-WEIGHT: bold; COLOR: #ff9900
}
.blue {
  COLOR: #0066cc
}
.bluebold {
  FONT-WEIGHT: bold; COLOR: #0066cc
}
.red {
  COLOR: #ff3333
}
.bodytextred {
  COLOR: #ff3333
}
.defaulterror {
  COLOR: #ff3333
}
.mandatorylabeltext {
  COLOR: #ff3333
}
.redbold {
  FONT-WEIGHT: bold; COLOR: #ff3333
}
.SubHead {
  FONT-WEIGHT: bold; COLOR: #ff3333
}
.Accent {
  FONT-WEIGHT: bold; COLOR: #ff3333
}
.NormalRed {
  FONT-WEIGHT: bold; COLOR: #ff3333
}
#rightpane .SubHead {
  FONT-WEIGHT: bold; COLOR: #ff3333; BACKGROUND-COLOR: transparent
}
.head {
  COLOR: #47869d
}
.white {
  COLOR: #ffffff
}
.nb {
  COLOR: #ffffff
}
.whitebold {
  FONT-WEIGHT: bold; FONT-SIZE: 11px; COLOR: #ffffff; TEXT-DECORATION: none
}
.whitesmall {
  FONT-SIZE: 8px; COLOR: #ffffff
}
.whitesmallbold {
  FONT-WEIGHT: bold; FONT-SIZE: 10px; COLOR: #ffffff
}
.whiteBannerLink {
  FONT-WEIGHT: bold; FONT-SIZE: 10px; COLOR: #ffffff
}
.whitebutton {
  FONT-WEIGHT: bold; FONT-SIZE: 9px; COLOR: #ffffff
}
.whitedate {
  FONT-SIZE: 11px; COLOR: #ffffff
}
.whitelink {
  FONT-SIZE: 10px; COLOR: #ffffff; TEXT-ALIGN: left
}
.bannerTextBlue {
  FONT-WEIGHT: bold; FONT-SIZE: 11px; COLOR: #b3e1e1
}
.breadcrumb {
  FONT-SIZE: 11px; COLOR: #b3e1e1
}
.pointlessblock {
  BACKGROUND-IMAGE: url(s.gif); WIDTH: 1px; BACKGROUND-REPEAT: no-repeat; HEIGHT: 1px
}
#rightpane .ModuleText {
  BACKGROUND-POSITION: 0px 50%; PADDING-LEFT: 14px; BACKGROUND-IMAGE: url(block.gif); BACKGROUND-REPEAT: no-repeat
}
#orange {
  COLOR: #ff9900
}
#blue {
  COLOR: #0066cc
}
#red {
  COLOR: #ff3333
}
#green {
  COLOR: #009999
}
#white {
  COLOR: #ffffff
}
#black {
  COLOR: #000000
}
#header {
  BACKGROUND-IMAGE: url(/portal/Images/bg_mk3.gif); BACKGROUND-REPEAT: no-repeat; BACKGROUND-COLOR: #ffffff
}
.toprow {
  PADDING-RIGHT: 10px; PADDING-LEFT: 10px; FONT-WEIGHT: bold; FONT-SIZE: 12px; VERTICAL-ALIGN: middle; COLOR: #ffffff; FONT-STYLE: normal; FONT-FAMILY: Arial, sans-serif; HEIGHT: 20px; BACKGROUND-COLOR: #4f7ca0; TEXT-ALIGN: left
}
.defaulttableth {
  PADDING-RIGHT: 10px; PADDING-LEFT: 10px; FONT-WEIGHT: bold; FONT-SIZE: 12px; VERTICAL-ALIGN: middle; COLOR: #ffffff; FONT-STYLE: normal; FONT-FAMILY: Arial, sans-serif; HEIGHT: 20px; BACKGROUND-COLOR: #4f7ca0; TEXT-ALIGN: left
}
.simpletableth {
  PADDING-RIGHT: 10px; PADDING-LEFT: 10px; FONT-WEIGHT: bold; FONT-SIZE: 12px; VERTICAL-ALIGN: middle; COLOR: #ffffff; FONT-STYLE: normal; FONT-FAMILY: Arial, sans-serif; HEIGHT: 20px; BACKGROUND-COLOR: #4f7ca0; TEXT-ALIGN: left
}
.tableth {
  PADDING-RIGHT: 10px; PADDING-LEFT: 10px; FONT-WEIGHT: bold; FONT-SIZE: 12px; VERTICAL-ALIGN: middle; COLOR: #ffffff; FONT-STYLE: normal; FONT-FAMILY: Arial, sans-serif; HEIGHT: 20px; BACKGROUND-COLOR: #4f7ca0; TEXT-ALIGN: left
}
.toprowalt {
  PADDING-RIGHT: 10px; PADDING-LEFT: 10px; FONT-WEIGHT: bold; FONT-SIZE: 12px; VERTICAL-ALIGN: middle; COLOR: #47869d; FONT-STYLE: normal; FONT-FAMILY: Arial, sans-serif; HEIGHT: 20px; BACKGROUND-COLOR: #ffffff; TEXT-ALIGN: left
}
.toprowalt TD {
  FONT-WEIGHT: bold; FONT-SIZE: 12px; FONT-STYLE: normal; FONT-FAMILY: Arial, sans-serif; BACKGROUND-COLOR: #ffffff; TEXT-ALIGN: left
}
.toprowalt A:link {
  FONT-WEIGHT: bold; FONT-SIZE: 12px; FONT-STYLE: normal; FONT-FAMILY: Arial, sans-serif; TEXT-ALIGN: left
}
.toprowalt A:visited {
  FONT-WEIGHT: bold; FONT-SIZE: 12px; FONT-STYLE: normal; FONT-FAMILY: Arial, sans-serif; TEXT-ALIGN: left
}
.toprowalt A:active {
  FONT-WEIGHT: bold; FONT-SIZE: 12px; FONT-STYLE: normal; FONT-FAMILY: Arial, sans-serif; TEXT-ALIGN: left
}
.toprowalt A:hover {
  FONT-WEIGHT: bold; FONT-SIZE: 12px; FONT-STYLE: normal; FONT-FAMILY: Arial, sans-serif; TEXT-ALIGN: left
}
#rightpane .toprowalt TD {
  FONT-WEIGHT: bold; FONT-SIZE: 12px; FONT-STYLE: normal; FONT-FAMILY: Arial, sans-serif; BACKGROUND-COLOR: #009999; TEXT-ALIGN: left
}
#rightpane .toprowalt A:link {
  FONT-WEIGHT: bold; FONT-SIZE: 12px; COLOR: #ffffff; FONT-STYLE: normal; FONT-FAMILY: Arial, sans-serif; BACKGROUND-COLOR: #009999; TEXT-ALIGN: left
}
#rightpane .toprowalt A:visited {
  FONT-WEIGHT: bold; FONT-SIZE: 12px; COLOR: #ffffff; FONT-STYLE: normal; FONT-FAMILY: Arial, sans-serif; BACKGROUND-COLOR: #009999; TEXT-ALIGN: left
}
#rightpane .toprowalt A:active {
  FONT-WEIGHT: bold; FONT-SIZE: 12px; COLOR: #ffffff; FONT-STYLE: normal; FONT-FAMILY: Arial, sans-serif; BACKGROUND-COLOR: #009999; TEXT-ALIGN: left
}
#rightpane .toprowalt A:hover {
  FONT-WEIGHT: bold; FONT-SIZE: 12px; COLOR: #ffffff; FONT-STYLE: normal; FONT-FAMILY: Arial, sans-serif; BACKGROUND-COLOR: #009999; TEXT-ALIGN: left
}
.tablerow {
  PADDING-RIGHT: 10px;
  PADDING-LEFT: 10px;
  FONT-WEIGHT: normal;
  FONT-SIZE: 12px;
  VERTICAL-ALIGN: top;
  COLOR: #636363;
  FONT-STYLE: normal;
  FONT-FAMILY: Arial, sans-serif;
  HEIGHT: 20px;
  BACKGROUND-COLOR: #F3F3F3;
  TEXT-ALIGN: left;
}
.tablerowalt {
  PADDING-RIGHT: 10px; PADDING-LEFT: 10px; FONT-WEIGHT: normal; FONT-SIZE: 12px; VERTICAL-ALIGN: top; COLOR: #636363; FONT-STYLE: normal; FONT-FAMILY: Arial, sans-serif; HEIGHT: 20px; BACKGROUND-COLOR: #ffffff; TEXT-ALIGN: left
}
.example {
  FONT-WEIGHT: bold; FONT-SIZE: 11px; COLOR: #4b9090
}
.PortalLargeLink {
  FONT-WEIGHT: normal; FONT-SIZE: 12px; COLOR: #47869d; LINE-HEIGHT: 16px; TEXT-DECORATION: none
}
.PortalSmallLink {
  FONT-SIZE: 11px; COLOR: #41869d; TEXT-DECORATION: none
}
.PortalLargeLink:link {
  FONT-WEIGHT: normal; FONT-SIZE: 12px; COLOR: #47869d; LINE-HEIGHT: 16px; TEXT-DECORATION: none
}
.PortalLargeLink:active {
  FONT-WEIGHT: normal; FONT-SIZE: 12px; COLOR: #47869d; LINE-HEIGHT: 16px; TEXT-DECORATION: none
}
.PortalLargeLink:visited {
  FONT-WEIGHT: normal; FONT-SIZE: 12px; COLOR: #47869d; LINE-HEIGHT: 16px; TEXT-DECORATION: none
}
.PortalLargeLink:hover {
  TEXT-DECORATION: underline
}
.PortalSmallLink:link {
  FONT-SIZE: 11px; COLOR: #41869d; TEXT-DECORATION: none
}
.PortalSmallLink:active {
  FONT-SIZE: 11px; COLOR: #41869d; TEXT-DECORATION: none
}
.PortalSmallLink:visited {
  FONT-SIZE: 11px; COLOR: #41869d; TEXT-DECORATION: none
}
.PortalSmallLink:hover {
  TEXT-DECORATION: underline
}
.orglarge:link {
  FONT-WEIGHT: normal; FONT-SIZE: 12px; COLOR: #666666; LINE-HEIGHT: 16px; TEXT-DECORATION: none
}
.orglarge:visited {
  FONT-WEIGHT: normal; FONT-SIZE: 12px; COLOR: #666666; LINE-HEIGHT: 16px; TEXT-DECORATION: none
}
.orglarge:hover {
  TEXT-DECORATION: underline
}
.orglarge:active {
  COLOR: #000000
}
.orgsmall:link {
  FONT-WEIGHT: normal; FONT-SIZE: 11px; COLOR: #666666; LINE-HEIGHT: 13px; TEXT-DECORATION: none
}
.orgsmall:visited {
  FONT-WEIGHT: normal; FONT-SIZE: 11px; COLOR: #666666; LINE-HEIGHT: 13px; TEXT-DECORATION: none
}
.orgsmall:hover {
  TEXT-DECORATION: underline
}
.orgsmall:active {
  COLOR: #000000
}
.orchartbgnd {
  PADDING-RIGHT: 1px; PADDING-LEFT: 1px; FONT-WEIGHT: normal; FONT-SIZE: 12px; VERTICAL-ALIGN: top; COLOR: #e5f5f5; FONT-STYLE: normal; FONT-FAMILY: Arial, sans-serif; HEIGHT: 20px; BACKGROUND-COLOR: #91c6c6; TEXT-ALIGN: center
}
.orgchartrow {
  FONT-WEIGHT: normal;
  FONT-SIZE: 12px;
  COLOR: #FFFFFF;
  FONT-STYLE: normal;
  FONT-FAMILY: Arial, sans-serif;
  BACKGROUND-COLOR: #999999;
  TEXT-ALIGN: left;
}

.pagecontent {
padding-left: 10px;
width: 1030px;
}

.pagecontentmain {
vertical-align: top;
width: 765px;
}

.pagecontentmain h1 {
font-weight: bold;
text-indent: 0pt; 
font-size: 1.2em;
line-height: 1.0em;
color: #fff;
margin-top: 8px;
margin-right: 3px;
margin-bottom: 0.2em;
padding: 4px 5px 4px 5px;
background: #969696;
}

tr:nth-child(even) {
    background-color: #f2f2f2
  }


#myInput {
  background-image: url('searchicon.png'); /* Add a search icon to input */
  background-position: 10px 12px; /* Position the search icon */
  background-repeat: no-repeat; /* Do not repeat the icon image */
  width: 100%; /* Full-width */
  font-size: 16px; /* Increase font-size */
  padding: 12px 20px 12px 40px; /* Add some padding */
  border: 1px solid #ddd; /* Add a grey border */
  margin-bottom: 12px; /* Add some space below the input */
}

#sclTable {
  border-collapse: collapse; /* Collapse borders */
  width: 100%; /* Full-width */
  border: 1px solid #ddd; /* Add a grey border */
  font-size: 18px; /* Increase font-size */
}

#sclTable th, #sclTable td {
  text-align: left; /* Left-align text */
  padding: 12px; /* Add padding */
}

#sclTable tr {
  /* Add a bottom border to all table rows */
  border-bottom: 1px solid #ddd;
}

#sclTable tr.header, #sclTable tr:hover {
  /* Add a grey background color to the table header and on hover */
  background-color: #f1f1f1;
}  


</style>
"@

$strHTMLfinal = '<html><head>' + $strHeaderStyle + '</head><body><div style="background-color:black"</style><img src="https://assets.roosterteeth.com/static/media/old-school-logo.3a949b0b.png" style="max-width:100%;max-height:100%;height:auto"/></div><h1 id="top">Local Archive: ' + $rootFolder + '</h1><p>Last updated: ' + $timestamp + '</p>' + $strHTMLcontents + $strHTML + '</body></html>'

# replace https://www.w3schools.com/html/html_entities.asp which would prevent the HTML from displaying correctly.
$strHTMLfinal = $strHTMLfinal.replace("&lt;","<")
$strHTMLfinal = $strHTMLfinal.replace("&gt;",">")
$strHTMLfinal = $strHTMLfinal.replace("&quot;",'"')
$strHTMLfinal = $strHTMLfinal.replace('<table','<table class="table"')
$strHTMLfinal = $strHTMLfinal.replace('<th','<th style="position: sticky"')

# write the HTML string to a file
$strHTMLfinal | out-file $strHTMLfile

# open the HTML file in the default browser
start-process $strHTMLfile