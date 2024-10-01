
#Zuzanna Jakubiak, IPIPM-171
function Analyzer{
#Definicja parametrów
param (
    [Parameter(Mandatory = $true)]                                                              #Scieżka do folderu to parametr obowiązkowy
    [String]
    $Folder
)

try                                                                                             #definicja sekcji skryptu
{
    if(-not(Test-Path -Path $Folder))                                                           #Sprawdzenie czy folder istnieje
    {
        throw "Folder is not found"                                                             #Jeśli nie, to otrzymujemy informację o błędzie
    }

    $FinalFolder = "$Folder\*"                                                                  #Ścieżka w której "szukamy" dokumentow .docx oraz .doc                                                            

    $WordFiles = Get-ChildItem -Path $FinalFolder -Include *doc,*docx -Recurse -Force           #"Szukanie" plików .docx oraz .doc. Pobranie z folderu elementów, również systemowych lub ukrytych o rozszerzeniu .docx oraz .doc

    $Properties = @(                                    #Tablica zawierająca podstawowe właściwości pliku
        "Author",                                       #Autor
        "Last author",                                  #Ostatni autor
        "Creation date",                                #Data utworzenia
        "Last save time",                               #Data zapisania
        "Number of pages",                              #Liczba stron
        "Number of words",                              #Liczba słow
        "Number of characters",                         #Liczba znaków
        "Number of characters (with spaces)"            #Liczba znaków włącznie ze spacjami
        )

    if($WordFiles.Count -gt 0)                                                                                          #Jeżeli liczba znalezionych plików jest > 0, to...
    {
        Write-Host "Analiza plikow" -ForegroundColor Yellow                                                                #Informacja o rozpoczęciu analizy plików

        foreach($SingleFile in $WordFiles)                                                                              #Dla każdego z plików...
        {
            $Asset = New-Object -TypeName PSObject                                                                      #Obiekt do przechowywania wyników

            Write-Host "`Praca nad plikiem :  $($SingleFile.Name)" -ForegroundColor Green                               #Informacja o tym, nad jakim plikiem właśnie "pracuje" program
            try                                                                                                         #Definicja sekcji skryptu
            {
                $fileName = $SingleFile.Name                                                                            #Nazwa pliku

                $Asset | Add-Member -MemberType NoteProperty -Name "File Name" -Value $fileName                         #Dodanie nazwy pliku do wyników

                $filePath = $SingleFile.DirectoryName                                                                   #Ścieżka pliku

                $Asset | Add-Member -MemberType NoteProperty -Name "File Path" -Value $filePath                         #Dodanie ścieżki do wyników
                
                $Application = New-Object -ComObject word.application                                                   #komenda do stworzenia obiektu MS Word Application
                $Application.Visible = $false                                                                           #Ustawienie widoczności pliku w tle na false(niewidoczny)

                $Document = $Application.documents.open($SingleFile.fullname,$false,$true)                              #otwarcie dokumentu Word w trybie ReadOnly (true jako trzeci parametr)
                $Document.Repaginate()                                                                                  #"Resetowanie" podziału stron w dokumencie
                $Binding = "System.Reflection.BindingFlags" -as [type]                                                  #Typ flag kontrolujących powiązania

                Foreach($Property in $Document.BuiltInDocumentProperties)                                               #Dla każdej właściowści (Properies) z wbudowanych...
                {
                    try                                                                                                 #Definicja sekcji skryptu
                    {
                        $Key= [System.__ComObject].invokemember("name",$Binding::GetProperty,$null,$property,$null)     #Klucz, wywołanie określonego członka bieżącego typu elementu. Nazwa "name", Binding flags, Binder "null" - DefaultBinder, obiekt - property, ostatni parametr to Object[] - Tablica zawierająca argumenty, które mają być przekazywane do elementu członkowskiego do wywołania.
                        $Val = [System.__ComObject].invokemember("value",$Binding::GetProperty,$null,$property,$null)   #Jak wyżej, tylko wartość.
                        
                        if($Key -in $Properties)                                                                        #Jeżeli klucz w Properties istnieje
                        {
                            $Asset | Add-Member -MemberType NoteProperty -Name $Key -Value $Val                         #Dodanie do wyników wszytskich właściwości i ich wartości
                        }
                    }
                    catch                                                                                               #Blok obsługi błędu
                    {
                        if($Key -in $Properties)
                        {
                            $Asset | Add-Member -MemberType NoteProperty -Name $Key -Value "Not Available"              #Jeżeli jakaś wartość nie występuje
                        }
                    }
                }

                $NumOfImages = $document.inlineshapes.count                                                             #Liczba obrazów w pliku

                $Asset | Add-Member -MemberType NoteProperty -Name "No of Images" -Value $NumOfImages                   #Dodanie do wyników liczby obrazów

                $NumOfSentences = $document.Sentences.count                                                             #Liczba zdań

                $Asset | Add-Member -MemberType NoteProperty -Name "Number of sentences" -Value $NumOfSentences         #Dodanie do wyników liczby zdań

                $AVGNumOfWordsPerSentence = ($Asset.PSObject.Properties["Number of words"].value)/ ($Asset.PSObject.Properties["Number of sentences"].value) #Obliczenie średniej liczby wyrazów na zdanie

                $Asset | Add-Member -MemberType NoteProperty -Name "Average number of words per sentence " -Value $AVGNumOfWordsPerSentence                  #Dodanie do wyników średniej liczby wyrazów na zdanie

                $AVGNumOfCharPerWord = ($Asset.PSObject.Properties["Number of characters"].value)/ ($Asset.PSObject.Properties["Number of words"].value)     #Oblicznie średniej liczby znaków na wyraz

                $Asset | Add-Member -MemberType NoteProperty -Name "Average number of characters per word " -Value $AVGNumOfCharPerWord                      #Dodanie do wyników średniej liczby znaków na wyraz

                $Document.Close([ref] [Microsoft.Office.Interop.Word.WdSaveOptions]::wdDoNotSaveChanges)                 #Zamknięcie pliku bez zapisywania zmian
                
                $Asset.PSObject.Properties | Select-Object Name, Value                                                   #Wyniki

                Add-Type -AssemblyName System.Windows.Forms                                                              #Załadowanie wymaganych typów do pracy z wykresami oraz Windows Forms                     
                Add-Type -AssemblyName System.Windows.Forms.DataVisualization

                $Chart = New-Object System.Windows.Forms.DataVisualization.Charting.Chart                                #Stworzenie wykresu
                $Chart.Width = 500                                                                                       #Określenie wysokości, szerokości i przesunięcia wykresu
                $Chart.Height = 400                                                                                      
                $Chart.Left = 40

                $ChartArea = New-Object System.Windows.Forms.DataVisualization.Charting.ChartArea                        #Stworzenie obszaru wykresu
                $Chart.ChartAreas.Add($ChartArea)                                                                        #Dodanie go do wykresu

                $Chart.Series.Add("Data")                                                                                #Dodanie serii danych
                $Chart.Series["Data"].Points.DataBindXY($Asset.psobject.Properties.name[6..15], $Asset.psobject.Properties.value[6..15]) #Wartości serii - czyli Wyniki, wybrano tylko te które mają wartości liczbowe

                $Form = New-Object System.Windows.Forms.Form                                                             #Stworzenie okna
                $Form.Width = 600                                                                                        #Szerokośc, wysokość okna
                $Form.Height = 600
                $Form.Controls.Add($Chart)                                                                               #Dodanie wykresu do okna

                $Chart.Titles.Add("Word File Analyzer `n File: $fileName")                                               #Opis wykresu oraz osi
                $ChartArea.AxisX.Title = "Properties"
                $ChartArea.AxisY.Title = "Values"
                $Form.ShowDialog()                                                                                       #Wyświetlenie okna
            }
            catch                                                                                                        #Blok obsługi błędu
            {
                Write-Host "Error, plik $($SingleFile.FullName)"
            }
        }
    }
    else                                                                                                                 #Jeżeli liczba szukanych plików = 0
    {
        Write-Host "Brak plikow o rozszerzeniu .doc oraz .docx we wskazanym folderze"  -ForegroundColor Red              #Komunikat o ich braku
    }
}
catch                                                                                                                    #Blok obsługi błędu
{
    Write-Host "Error" -ForegroundColor Red
}
}


Analyzer -Folder "C:\Users\Zuzanna Jakubiak\Desktop\MCHTR\SEM7\SYOP\PROJ2\tu_sa_pliki"                                    #Wywołanie funkcji, testy



