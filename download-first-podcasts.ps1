# edit the value of rssFeedUrl to match the URL from your FIRST premium 
$rssFeedUrl = 'https://YOUR-FIRST-PODCAST-URL.rss'

# edit the value of $PodcastName to match the podcast name.
$PodcastName = "PODCAST-NAME"

# edit the value of $downloadFolder if needed. By default it uses your Downloads folder.
$downloadFolder = "$($env:userprofile)\Downloads"

# set TLS1.2
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

# Download the RSS feed
$rssFeed = Invoke-WebRequest -Uri $rssFeedUrl

# convert the feed to XML format
$xml = [xml]$rssFeed

# write XML to file for future reference
$rssFeed.Content | out-file "$downloadFolder\feed-$PodcastName.xml"

# get all items in the feed
$items = $xml.rss.channel.item

$episodeNumber = 0

# loop through each item from last end (oldest episode) to beginning (newest episode)
for ($i = $items.Count; $i -ge 1; $i--) {
    # get the current item
    $item = $items[$i]

    ## $item = $items[$items.count - 1]

    # set the episode number
    $episodeNumber = $episodeNumber + 1

    # set the episode title
    # $itemTitle = $item.title

    # set the episode download URL 
    $itemURL = $item.enclosure.url

    # set the mp3 file name
    $itemFileName = "$PodcastName $episodeNumber.mp3"

    # set the full path to the mp3 file in the download folder
    $itemFilePath = "$downloadFolder\$itemFileName"
    
    ##write-host "Downloading episode $i / $($items.count) - title: $itemFileName)"

    write-host "$episodeNumber / $($items.count) - " -NoNewline -ForegroundColor Green
    # if the file path does not exist, download the file
    if (!(test-path $itemFilePath)){
        Start-BitsTransfer $itemURL -Destination $itemFilePath
        write-host "Download: " -NoNewline
        write-host $itemURL -NoNewline -ForegroundColor Yellow
        write-host " to: " -NoNewline
    } else {
        write-host "skip already downloaded file: " -NoNewline
    }
    write-host $itemFileName -ForegroundColor Cyan
    
    ## start-sleep -Seconds 5
}