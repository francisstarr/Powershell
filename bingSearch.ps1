for ($a = 0; $a -lt 30; $a++) {
	$raw = Get-Content -Path words.txt 
	$max=$raw.Count
	$index=Get-Random -Minimum -0 -Maximum $max
	$word=Get-Content words.txt |select -Index $index
    Start-Process microsoft-edge:https://www.bing.com/search?q=$word
    $delay=get-random -minimum 2100 -maximum 3000
	start-sleep -Milliseconds $delay
	$delay	
}
