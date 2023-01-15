function Get-Page {
    Param
    (
        [Parameter(Mandatory = $true, Position = 0)] # Parametr obowiązkowy to ścieżka
        [string] $Path, # Ustawianie typu parametru na string
        [Parameter(Mandatory = $true, Position = 1)] # parametr obowiązkowy przyjmujący rozszerzenie do znalezienia
        [string] $TextToFind # Ustawianie typu parametru na string
    )
    $matchWildcards = $false 
    $matchCase = $true
    $wdFindStop = 0 # some Word constants
    $wdActiveEndPageNumber = 3 # some Word constants
    $wdStory = 6 # some Word constants
    $wdGoToPage = 1 # some Word constants
    $wdGoToAbsolute = 1 # some Word constants
    $ReadOnly = $false  # when ReadONly was set to $true, it gave me an error on 'Selection.GoTo()' 
    $ConfirmConversions = $false
    $Word = New-Object -ComObject Word.Application
    $Word.Visible = $true
    $Document = $Word.Documents.Open($Path, $ConfirmConversions, $ReadOnly)
    $range = $Document.Content
    $range.Find.ClearFormatting(); 
    $range.Find.Forward = $true 
    $range.Find.Text = $TextToFind
    $range.Find.Wrap = $wdFindStop
    $range.Find.MatchWildcards = $matchWildcards
    $range.Find.MatchCase = $matchCase
    $range.Find.Execute()
    if ($range.Find.Found) {
        # get the pagenumber
        $page = $range.Information($wdActiveEndPageNumber)
        Write-Host "Found '$textToFind' on page $page" -ForegroundColor Green
        $Word.Selection.GoTo($wdGoToPage, $wdGoToAbsolute, $page) | Out-Null 
        foreach ($docrange in $range.Words) {
            if ($docrange.Text.Trim() -eq $TextToFind) 
            {
                $docrange.highlightColorIndex = [Microsoft.Office.Interop.Word.WdColorIndex]::wdYellow 
            }
        }
    }
    else 
    {
        Write-Host "'$textToFind' not found" -ForegroundColor Red
        $Word.Selection.GoTo($wdGoToPage, $wdGoToAbsolute, 1) | Out-Null 
    }
    read-host "Press ENTER to end..."
    $Document.close() # Zamkniecie doca
    $Word.Quit() # Zamkniecie word
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($range) | Out-Null # cleanup com objects
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($Document) | Out-Null # cleanup com objects
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($Word) | Out-Null # cleanup com objects
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()	
}
Get-Page -Path D:\testy\wwww.docx -TextToFind "ipsum" # Uruchomienie funckji 