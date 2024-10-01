function FindReplace{
    [CmdletBinding()]                                                                                   #Zapewnia dostęp do funkcji poleceń cmdlet.
    param                                                                                               #Definicja parametrów
    (
        [Parameter(Mandatory = $true)]                                                                  #Folder to parametr nieobowiązkowy
        [string]
        $Folder,
        [Parameter(Mandatory = $true)]                                                                  #Szukany ciąg znaków jest paramerem obowiązkowym
        [string]
        $Find,
        [Parameter()]                                                                                   #Parametr nieobowiązkowy, jego podanie spowoduje zamianę ciągu znaków Find na podany Replace
        [string]
        $Replace,
        [Parameter()]                                                                                   #Parametr nieobowiązkowy, czy uwzględniamy wielkość znaków
        [switch]
        $LetterSize
    )
    begin                                                                                               #Definicja sekcji skryptu
    {
        $Find = [regex]::Escape($Find)                                                                  #Automatyczna zmiana znaczenia znaków w łańcuchu do wyrażeń regularnych
        $FinalFolder = "$Folder\*"                                                                      #Ścieżka w której szukamy plików txt
        $FilePath = Get-ChildItem -Path $FinalFolder -Include *txt -Recurse -Force                      #"Szukanie" plików .txt. Pobranie z folderu elementów, również systemowych lub ukrytych o rozszerzeniu .txt
        $Asset = New-Object -TypeName PSObject                                                          #Do przechowywania wyników do wykresu                                                                                          
    }
    process                                                                                             #Definicja sekcji skryptu
    {
        try                                                                                             #Definicja sekcji skryptu
        {
            if(-not(Test-Path -Path $Folder))                                                           #Sprawdzenie czy folder istnieje
            {
                throw "Folder nie odnaleziony"                                                          #Jeśli nie, to otrzymujemy informację o błędzie
            }

            if($FilePath.count -gt 0)                                                                   #Jeżeli liczba znalezionych plików jest > 0, to...
            {                                                               
                Write-Host "Analiza plikow" -ForegroundColor Yellow                                     #Informacja o rozpoczęciu analizy plików                                    

                    foreach ($File in $FilePath)                                                        #Dla każdego z plików...
                    {
                        $fileName = $File.Name                                                          #Uzyskanie nazwy pliku

                        Write-Host "[*]Praca nad plikiem $fileName`:\"  -ForegroundColor Green          #Informacja o tym na jakim pliku aktualnie pracuje skrypt

                        if ($Replace)                                                                   #Jeżeli wykorzystano parametr replace...
                        {
                            if ($LetterSize)                                                            #Jeżeli wykorzystano parametr LetterSize
                            {
                                Write-Host "Nastapi zamiana '$Find' na '$Replace' z uwzglednieniem wielkosci znakow `n"     #Indormacja
                                (Get-Content $File) -creplace $Find, $Replace | Add-Content -Path "$File.tmp" -Force        #zamiana Find i Replace, oraz dodanie pliku tmp
                            }
                            else                                                                        #Jeżeli nie wykorzystano parametru LetterSize
                            {
                                Write-Host "Nastapi zamiana '$Find' na '$Replace' bez uwzglednienia wielkosci znakow `n"    #Informacja
                                (Get-Content $File) -replace $Find, $Replace | Add-Content -Path "$File.tmp" -Force         #zamiana Find i Replace, oraz dodanie pliku tmp
                            }
                            Remove-Item -Path $File                                                     #usunięcie istniejącego pliku
                            Move-Item -Path "$File.tmp" -Destination $File                              #przeniesienie pliku ze zmienionymi danymi do folderu z plikami
                        } 
                        else                                                                            #Jeżeli nie wykorzystano parametru replace...
                        {
                            if ($LetterSize)                                                            #Jeżeli wykorzystano parametr LetterSize
                            {
                                $Ret = Select-String -Path $File -Pattern $Find -CaseSensitive          #Znaleznienie ciągów szukanych znaków w pliku z uwzględnieiem wielkości znakow
                            }
                            else                                                                        #Jeżeli nie wykorzystano parametru LetterSize
                            {
                                $Ret = Select-String -Path $File -Pattern $Find                         #Znaleznienie ciągów szukanych znaków w pliku
                            }
                            $Counter =  $Ret.count                                                      #Liczba znalezionych ciągów znaków w pliku
                            $lines = $Ret.LineNumber                                                    #Wiersze w których znaleziono ciągi znaków
                            $NumOfLines = $Lines.count                                                  #Liczba wierszy w których znaleziono ciągi znaków
                            if($Ret.count -gt 0)                                                        #Jeżeli liczba znalezionych ciągów znaków jest > 0 to...
                            {
                                Write-Host "Znaleziono $Counter '$Find' w $NumOfLines wierszach (numery: $Lines) w pliku $FileName `n"    #Infomacja o wynikach  
                            }
                            else                                                                        #Jeżeli liczba znalezionych ciągów znaków nie jest > 0 to...
                            {
                                Write-Host "Nie znaleziono `n"                                               #Infomacja o braku
                            }
                            if($Counter -gt 0)                                                              #Jeżeli liczba znalezionych ciągów znaków w pliku >0...
                            {
                            $Asset | Add-Member -MemberType NoteProperty -Name $fileName -Value $Counter                        #Dodanie nazwy pliku do wyników i liczby znalezionych ciągów znaków
                            }
                        } 
                    }
                    if (!$Replace)                                                                      #Jeżeli nie wykorzystano parametru Replace
                    {                                                      
                    if($Asset.psobject.properties.count -gt 0)                                          #Jeżeli istnieją jakieś dane do stworzenia wykresu
                    {
                    #WYKRES
                        Add-Type -AssemblyName System.Windows.Forms                                                              #Załadowanie wymaganych typów do pracy z wykresami oraz Windows Forms                     
                        Add-Type -AssemblyName System.Windows.Forms.DataVisualization

                        $Chart = New-Object System.Windows.Forms.DataVisualization.Charting.Chart                                #Stworzenie wykresu
                        $Chart.Width = 1700                                                                                      #Określenie wysokości, szerokości i przesunięcia wykresu
                        $Chart.Height = 800                                                                                      
                        $Chart.Left = 80

                        $ChartArea = New-Object System.Windows.Forms.DataVisualization.Charting.ChartArea                        #Stworzenie obszaru wykresu
                        $Chart.ChartAreas.Add($ChartArea)                                                                        #Dodanie go do wykresu

                        $Chart.Series.Add("Data")                                                                                #Dodanie serii danych
                        $Chart.Series["Data"].Points.DataBindXY($Asset.psobject.Properties.name, $Asset.psobject.Properties.value) #Wartości serii - czyli Wyniki, wybrano tylko te które mają wartości liczbowe
                        $Chart.Series["Data"].IsvalueShownAsLabel=$true

                        $Form = New-Object System.Windows.Forms.Form                                                             #Stworzenie okna
                        $Form.Width = 2000                                                                                       #Szerokośc, wysokość okna
                        $Form.Height = 1200
                        $Form.Controls.Add($Chart)                                                                               #Dodanie wykresu do okna

                        $Chart.Titles.Add("Analiza pliku txt `n szukane: $Find")                                                 #Opis wykresu oraz osi
                        $ChartArea.AxisX.Title = "Pliki .txt"
                        $ChartArea.AxisX.Interval = 1
                        $ChartArea.AxisY.Title = "Liczba wystapien slowa $Find "
                        $Form.ShowDialog()
                    }
                    else                                                                                                        #Informajca o braku dopasowan
                    {
                        Write-Host "Brak dopasowan" -ForegroundColor Red
                    }
                }

            }
            elseif ($FilePath.count -eq 0)                                                                            #Jeżeli liczba szukanych plików nie jest > 0
            {
                Write-Host "Brak plikow o rozszerzeniu .txt we wskazanym folderze"  -ForegroundColor Red              #Komunikat o braku plików
            }
        }      
        catch                                                                                                         #Blok obsługi błędu
        {
            Write-Error $_.Exception.Message
        }
    }
}

#TESTY
#FindReplace -Folder "C:\Users\Zuzanna Jakubiak\Desktop\MCHTR\SEM7\SYOP\PROJ3\pusty" -Find "boczek"
#FindReplace -Folder "C:\Users\Zuzanna Jakubiak\Desktop\MCHTR\SEM7\SYOP\PROJ3" -Find "Boczek"
#FindReplace -Folder "C:\Users\Zuzanna Jakubiak\Desktop\MCHTR\SEM7\SYOP\PROJ3" -Find "Boczek" -LetterSize 
#FindReplace -Folder "C:\Users\Zuzanna Jakubiak\Desktop\MCHTR\SEM7\SYOP\PROJ3" -Find "curry" -Replace "garam masala"
#FindReplace -Folder "C:\Users\Zuzanna Jakubiak\Desktop\MCHTR\SEM7\SYOP\PROJ3" -Find "Pomidor" -Replace "Pomidor malinowy" -LetterSize  