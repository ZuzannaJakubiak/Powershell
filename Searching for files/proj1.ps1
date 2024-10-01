using namespace System.Collections.Generic
function Searching{
    #Definicja parametrów
    param (
        [Parameter(Mandatory = $true)]
        [string] $Path,
        [Parameter(Mandatory = $true)]
        [string] $Pattern,
        [Parameter(Mandatory = $true)]
        $ExpectedSize,
        [Parameter(Mandatory = $true)]
        [string] $Year,
        [Parameter(Mandatory = $true)]
        [string] $Month,
        [Parameter(Mandatory = $true)]
        [string] $Day
    )
#DATY - Parametrami funkcji jest między innymi data. Dlatego sprawdzam czy taka data może istnieć ( nie jest z przyszłości)
if ($Year -and $Month -and $Day) {
    $FindDate = Get-Date -Year $Year -Month $Month -Day $Day   
}
#Czy taka data może istnieć?
$dateToday = Get-Date
if ($dateToday -lt $FindDate) {
    Write-Host "File not found" 
    return ""
}
# Polecenie pozwalające na wyszukanie plików o konkretnym wzorze. W tym prrzypadku z wybranym rozszerzeniem, wybraną datą zapisania z marginesem (+/- 50 dni), oraz wybranym rozmiarze, także z marginesem.
$Elements = Get-ChildItem -Path $Path -Include ("*" + $Pattern) -File -Recurse -ErrorAction SilentlyContinue | Where-Object { $_.LastWriteTime -ge $FindDate.adddays(-50) -and $_.LastWriteTime -le $Finddate.adddays(50) -and $_.Length -ge $ExpectedSize * 0.4 -and $_.Length -le $ExpectedSize * 4 } | Select-Object FullName, Length, LastWriteTime

#Etap rozdzielenia ścieżki no podfoldery i konkretny plik
$ListOfDir = [List[PSObject]]::new()   
$ListOfFiles = [List[PSObject]]::new()   

foreach ($Element in $Elements) {
    $name = $Element.FullName                   #cała ścieżka
    $name = $name.Split("\")                    #podzielenie dzięki \
    $ListOfDir.add($name[0..($name.Count - 2)]) #zapisanie nazw folderow do listy
    $ListOfFiles.add($name[-1])                 #zapisanie nazwy pliku do konkretnej listy
}

#Etap budowy drzewa 
$i = 0
$j = 0
$tree = [List[PSObject]]::new()   

for ($i = 0; $i -lt $ListOfDir.Count; $i++) {
    if ($i -eq 0) {
        for ($j = 0; $j -lt $ListOfDir[$i].Count; $j++) {                       #pętla do wyboru nazw folderow
            $indentation = ("   " * $j) + "----" + $ListOfDir[$i][$j]           #gałąź
            $tree.add($indentation)                                             #dodanie gałęzki do drzewa
            $previousPath = ($j - 1)                                            #głebokość scieżki/wcięcia

        }
    }
    else {                                                                      #dla każdej kolejnej scieżki
        for ($j = 0; $j -lt $ListOfDir[$i].Count; $j++) {                       #pętla do wyboru nazw folderow
            if ($j -lt $previousPath) {                                         #Warunek wypisywania 
                if ($ListOfDir[$i][$j] -ne $ListOfDir[$i - 1][$j]) {            #Wypisujemy tylko te, ktore wczesniej nie wystąpiły
                    $indentation = ("   " * $j) + "----" + $ListOfDir[$i][$j]   #gałąź
                    $tree.add($indentation)                                     #dodanie gałęzki do drzewa
                }
            }
            elseif ($j -eq $previousPath) {                                     #warunek wypisania galezi tak glebokiej jak poprzednia
                if ($ListOfDir[$i][$j] -ne $ListOfDir[$i - 1][$j]) {            #Wypisujemy tylko te, ktore wczesniej nie wystąpiły
                    $indentation = ("   " * $j) + "----" + $ListOfDir[$i][$j]   #gałąź
                    $tree.add($indentation)                                     #dodanie gałęzki do drzewa
                }
            }
            else {                                                              #gałąz glebsza niz poprzednia
                $indentation = ("   " * $j) + "----" + $ListOfDir[$i][$j]       #gałąź
                $tree.add($indentation)                                         #dodanie gałęzki do drzewa
            }
        }
    }

    $indentation =  ("   " * $j) + "----" + $ListOfFiles[$i]                    #Galaz z nazwa pliku
    $tree.add($indentation)                                                     #dodanie gałęzki do drzewa
    $previousPath = ($j - 1)                                                    #przypisanie glebokosci/wciecia
}

#Etap uzupełniania listy dat oraz listy rozmiarów
$i = 0
$ListOfLastWritten = [List[PSObject]]::new()                                    
$ListOfLen = [List[PSObject]]::new()   
foreach ($par in $tree) {                                                       #Pętla po każdej gałezi
    if ($par.Contains($Pattern)) {                                              #Warunek wystepowania rozszerzenia
        $ListOfLastWritten.Add($Elements[$i].LastWriteTime)                     #Dodawanie daty modyfikacji do listy
        $ListOfLen.Add(([string][int]($Elements[$i].Length / 1KB)) + "KB")      #Dodawanie rozmiaru do listy
        $i += 1
    }
    else {                                                                      #dotyczy gałezi bez rozszerzenia
        $ListOfLastWritten.Add("")                                              #puste miejsce do listy dat
        $ListOfLen.Add("")                                                      #puste miejsce do listy rozmiarów
    }
}

#Etap wyświetlania
for ($i = 0; $i -lt $tree.Count; $i++) {
    [PSCustomObject]@{
        TreeV = $tree[$i]
        LastOne = $ListOfLastWritten[$i]
        Size = $ListOfLen[$i]
    }
}
}

#Testy
Searching -Path "C:\Users\Public\jazda" -Pattern ".pdf"  -ExpectedSize 2MB -Year 2022 -Month 12 -Day 1 

