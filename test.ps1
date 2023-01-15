function Get-Page 
{
    Param
    (
        [Parameter(Mandatory = $true, Position = 0)] # Parametr obowiązkowy to ścieżka
        [string] $Path, # Ustawianie typu parametru na string
        [Parameter(Mandatory = $true, Position = 1)] # parametr obowiązkowy przyjmujący slowo do znalezienia
        [string] $TextToFind, # Ustawianie typu parametru na string
        [Parameter(Mandatory = $false, Position = 2)] # parametr nieobowiazkowy, oznacza on szukanie slow jedynie podobnych od poczatku slowa
        [bool] $MatchPrefix, # Ustawianie typu parametru na bool
        [Parameter(Mandatory = $false, Position = 3)] # parametr nieobowiazkowy, oznacza on szukanie slow jedynie podobnych od konca slowa
        [bool] $MatchSuffix, # Ustawianie typu parametru na bool
        [Parameter(Mandatory = $false, Position = 4)] # parametr nieobowiazkowy, oznacza on szukanie slow calkowicie podobnych do slowa
        [bool] $MatchWholeWord, # Ustawianie typu parametru na bool
        [Parameter(Mandatory = $false, Position = 5)] # parametr nieobowiazkowy, oznacza on szukanie slow czesciowy podobnych do  slowa
        [bool] $MatchSoundsLike # Ustawianie typu parametru na bool
    )
    read-host "Press ENTER to start..."
    $matchWildcards = $false #specjalny operator
    $MatchCase = $false #dopasowanie duzych liter
    $Forward = $ture #tyb wyszukiwanie
    $wdFindStop = 0 # some Word constants
    $wdActiveEndPageNumber = 3 # some Word constants
    $wdStory = 6 # some Word constants
    $wdGoToPage = 1 # some Word constants
    $wdGoToAbsolute = 1 # some Word constants
    $ReadOnly = $false  # when ReadONly was set to $true, it gave me an error on 'Selection.GoTo()' 
    $ConfirmConversions = $false
    $MatchAllWordForms = $false # sit - sitting, sat
    $MatchPhrase = $false #ignores all white space and control characters between words
    $IgnoreSpace = $false
    $IgnorePunct = $false
    $MatchControl = $false #bidirectional control 
    $MatchAlefHamza = $false #Arabic
    $MatchDiacritics = $false # right-to-left language
    $MatchKashida = $false # Arabic
    $Word = New-Object -ComObject Word.Application # nowy objekt word
    $Word.Visible = $true # widzialnosc dokumentu
    $Document = $Word.Documents.Open($Path, $ConfirmConversions, $ReadOnly) # otwierany dokument
    $Paragraphs = $Document.Paragraphs # paragrawfy w dokumencie
    $NumPara = 0 # liczba paragrafow
    $Succes2 = $false # sprawdzenie czy udalo sie znalazc jakie kolwiek slowo
    foreach ($Paragraph in $Paragraphs) 
    {
        $NumPara++ # liczenie paragrafow
        if ($Paragraph.Range.Text.Contains($TextToFind)) 
        {
            $range = $Paragraph.Range # wszystko ponizej do szukanai
            $range.Find.ClearFormatting(); 
            $range.Find.Forward = $Forward 
            $range.Find.Text = $TextToFind
            $range.Find.Wrap = $wdFindStop
            $range.Find.MatchWildcards = $matchWildcards
            $range.Find.MatchCase = $MatchCase
            $range.Find.MatchWholeWord = $MatchWholeWord
            $range.Find.MatchSoundsLike = $MatchSoundsLike
            $range.Find.MatchAllWordForms = $MatchAllWordForms
            $range.Find.MatchPhrase = $MatchPhrase
            $range.Find.MatchPrefix = $MatchPrefix
            $range.Find.MatchSuffix = $MatchSuffix
            $range.Find.IgnoreSpace = $IgnoreSpace
            $range.Find.IgnorePunct = $IgnorePunct
            $range.Find.MatchControl = $MatchControl
            $range.Find.MatchAlefHamza = $MatchAlefHamza
            $range.Find.MatchDiacritics = $MatchDiacritics
            $range.Find.MatchKashida = $MatchKashida
            $range.Find.Execute() | Out-Null # null po to zeby w terminalu nie bylo useless rzeczy
            if ($range.Find.Found) 
            {
                $Succes2 = $true # znalezienie slowa 
                $page = $range.Information($wdActiveEndPageNumber) # pobranie nr strony
                Write-Host "Found '$textToFind' on page $page" -ForegroundColor Green # wyrduk info
                $Word.Selection.GoTo($wdGoToPage, $wdGoToAbsolute, $page) | Out-Null  # przejscie do odpowiedniej strony
                $Succes = $false # znalezienie dokladnego slowo 
                foreach ($docrange in $range.Words) 
                {
                    if ($docrange.Text.Trim() -eq $TextToFind) 
                    {
                        $docrange.highlightColorIndex = [Microsoft.Office.Interop.Word.WdColorIndex]::wdYellow 
                        $Succes = $true
                    }
                    if ($Succes -eq $false) # jesli nie znalazlo dokladnego slowa
                    {
                        $range.highlightColorIndex = [Microsoft.Office.Interop.Word.WdColorIndex]::wdYellow 
                    }
                }
                $range.GoTo() | Out-Null # null po to zeby w terminalu nie bylo useless rzeczy
                read-host "Press ENTER to continue..."
            }
            else 
            {
                Write-Host "'$textToFind' not found" -ForegroundColor Red # wyrduk info
                $Word.Selection.GoTo($wdGoToPage, $wdGoToAbsolute, 1) | Out-Null  # pojscie do poczatkowej strony
            }
        }
    }
    if ($Succes2 -eq $false) 
    {
        Write-Host "'$textToFind' not found" -ForegroundColor Red
        $Word.Selection.GoTo($wdGoToPage, $wdGoToAbsolute, 1) | Out-Null
    }
    read-host "Press ENTER to end..."
    $Document.Close([Microsoft.Office.Interop.Word.WdSaveOptions]::wd) # Zamkniecie doca
    $Word.Quit() # Zamkniecie word
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($Paragraphs) | Out-Null # cleanup com objects
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($Document) | Out-Null # cleanup com objects
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($Word) | Out-Null # cleanup com objects
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
}
Get-Page -Path D:\testy\Przepisy-kulinarne.docx -TextToFind "ko" -MatchPrefix $false -MatchSuffix $true -MatchWholeWord $false -MatchSoundsLike $false # Uruchomienie funckji 