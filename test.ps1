function Get-Tree 
{
    Param
    (
        [Parameter(Mandatory = $true, Position = 0)] # Parametr obowiązkowy to ścieżka
        [string] $Path # Ustawianie typu parametru na string
    )
    $Word = New-Object -ComObject Word.application  # Przypisanie obiektu worda do zmiennej
    $Word.Visible = $False # Ustawienie by otwarcie nie bylo widoczne
    $Document = $word.Documents.Open($Path) # Otworzenie dokumentu
    $Paragraphs = $Document.Paragraphs # Pobranie danych wszystkie paragrafy

    $Document.close() # Zamkniecie doca
    $Word.Quit() # Zamkniecie word
}
Get-Tree -Path D:\testy\wwww.docx  # Uruchomienie funckji 