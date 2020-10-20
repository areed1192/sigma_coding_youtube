Sub WorkingWithSlicers()
'
' WorkingWithSlicers Macro
' In this tutorial, we will explore how to work with slicers using Excel VBA.
'

'Declare Variables.
Dim xlApp As Application
Dim xlActiveBook As Workbook
Dim xlActiveSheet As Worksheet

Dim xlPokemonTable As ListObject

Dim xlPokemonSlicers As Slicers
Dim xlPokemonSlicer As Slicer
Dim xlPokemonSlicerCaches As SlicerCaches
Dim xlPokemonSlicerCache As SlicerCache

Dim xlPokemonSetSlicerCache As SlicerCache
Dim xlPokemonNameSlicerCache As SlicerCache
Dim xlPokemonSeriesSlicerCache As SlicerCache
Dim xlPokemonRaritySlicerCache As SlicerCache

Dim xlPokemonSetSlicer As Slicer
Dim xlPokemonNameSlicer As Slicer
Dim xlPokemonSeriesSlicer As Slicer
Dim xlPokemonRaritySlicer As Slicer


'Set Variables.
Set xlApp = Application
Set xlActiveBook = xlApp.ThisWorkbook
Set xlActiveSheet = xlActiveBook.Worksheets("Pokemon_Data")

'Select the List Object
Set xlPokemonTable = xlActiveSheet.ListObjects("CardSeries")

'Grab the Slicer Caches.
Set xlPokemonSlicerCaches = xlActiveBook.SlicerCaches

'loop through each
For Each xlPokemonSlicerCache In xlPokemonSlicerCaches
    
    Debug.Print xlPokemonSlicerCache.Name
    
Next


    'Add a SlicerCache for the "Name" Column, SEEMS TO BE A BUG WITH NAMED ARGUMENTS FOR "ADD2".
    Set xlPokemonNameSlicerCache = xlPokemonSlicerCaches.Add2(xlPokemonTable, "name", "SlicerCachePokemonName")
        
    'Add a Slicer from the Slicer Cache that represents the "Name" Column.
    Set xlPokemonNameSlicer = xlPokemonNameSlicerCache.Slicers.Add(SlicerDestination:=xlActiveSheet.Name, _
                                                                   Name:="PokemonSlicerName", _
                                                                   Caption:="Pokemon Name", _
                                                                   Top:=Application.InchesToPoints(1.75), _
                                                                   Left:=Application.InchesToPoints(0.2), _
                                                                   Width:=Application.InchesToPoints(2.01), _
                                                                   Height:=Application.InchesToPoints(1.18))

    'Add a SlicerCache for the "Rarity" Column.
    Set xlPokemonRaritySlicerCache = xlPokemonSlicerCaches.Add2(xlPokemonTable, "rarity", "PokemonSlicerCacheRarity")

    'Add a Slicer from the Slicer Cache that represents the "Rarity" Column.
    Set xlPokemonRaritySlicer = xlPokemonRaritySlicerCache.Slicers.Add(SlicerDestination:=xlActiveSheet.Name, _
                                                                       Name:="PokemonSlicerRarity", _
                                                                       Caption:="Pokemon Card Rarity", _
                                                                       Top:=Application.InchesToPoints(0.2), _
                                                                       Left:=Application.InchesToPoints(4.62), _
                                                                       Width:=Application.InchesToPoints(6.53), _
                                                                       Height:=Application.InchesToPoints(1.3))

    'Add a SlicerCache for the "Set" Column.
    Set xlPokemonSetSlicerCache = xlPokemonSlicerCaches.Add2(xlPokemonTable, "set", "PokemonSlicerCacheSet")

    'Add a Slicer from the Slicer Cache that represents the "Set" Column.
    Set xlPokemonSetSlicer = xlPokemonSetSlicerCache.Slicers.Add(SlicerDestination:=xlActiveSheet.Name, _
                                                                 Name:="PokemonSlicerSet", _
                                                                 Caption:="Pokemon Card Set", _
                                                                 Top:=Application.InchesToPoints(1.75), _
                                                                 Left:=Application.InchesToPoints(0.87), _
                                                                 Width:=Application.InchesToPoints(2.01), _
                                                                 Height:=Application.InchesToPoints(1.18))

    'Add a SlicerCache for the "Series" Column.
    Set xlPokemonSeriesSlicerCache = xlPokemonSlicerCaches.Add2(xlPokemonTable, "series", "PokemonSlicerCacheSet")

    'Add a Slicer from the Slicer Cache that represents the "Series" Column.
    Set xlPokemonSeriesSlicer = xlPokemonSeriesSlicerCache.Slicers.Add(SlicerDestination:=xlActiveSheet.Name, _
                                                                       Name:="PokemonSlicerSeries", _
                                                                       Caption:="Pokemon Card Series", _
                                                                       Top:=Application.InchesToPoints(1.75), _
                                                                       Left:=Application.InchesToPoints(4.62), _
                                                                       Width:=Application.InchesToPoints(6.53), _
                                                                       Height:=Application.InchesToPoints(1.25))
                                                                 
    'Change the number of column for the Rarity Slicer.
    xlPokemonRaritySlicer.NumberOfColumns = 3
    xlPokemonRaritySlicer.SlicerCache.SortItems = xlSlicerSortAscending
    xlPokemonRaritySlicer.SlicerCache.CrossFilterType = xlSlicerCrossFilterHideButtonsWithNoData
    
    'Loop through each Cache.
    For Each xlPokemonSlicerCache In xlPokemonSlicerCaches
    
        'Loop through each Slicer.
        For Each xlPokemonSlicer In xlPokemonSlicerCache.Slicers
            
            'Change the Style.
            xlPokemonSlicer.Style = "SlicerStyleLight2"
            
        Next
    Next
    
    'Set the Filter.
    xlPokemonSetSlicerCache.SlicerItems("Base").Selected = True
    xlPokemonRaritySlicerCache.SlicerItems("Rare").Selected = True
    
End Sub