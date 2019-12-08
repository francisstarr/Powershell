
$ie = New-Object -comobject InternetExplorer.Application
$ie.visible=$true
$ie.FullScreen=$false
$ie.navigate("https://account.microsoft.com/rewards")
while($ie.ReadyState -ne 4 ) {start-sleep -m 6001}
$n=$ie.document.getElementsByClassName("c-call-to-action c-glyph f-lightweight").length
#$n
for($i = 0; $i -lt $n; $i=$i+2){
    $link= $ie.document.getElementsByClassName("c-call-to-action c-glyph f-lightweight")[$i]
    if ($link.innerText -like '*10 points*' -AND  $link.tabIndex -eq 0){
        $delay=get-random -minimum 2100 -maximum 3000
        start-sleep -Milliseconds $delay
        $link.click()
    }
}
