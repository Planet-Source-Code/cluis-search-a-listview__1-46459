Attribute VB_Name = "Module1"
Public Function SearchListview(SearchText As String, ListToSearch As ListView, ListToDisplayResults As ListView, ProgressBar As ProgressBar)

'Declares all Variable for Function
Dim i, x, y As Integer
Dim length1, length2, length3 As Integer

'Just For No Messed Up Search Results
ListToDisplayResults.Sorted = False

'Clears The Results List
ListToDisplayResults.ListItems.Clear

'Sets The Progress Bar's Maximun Value
ProgressBar.Max = ListToSearch.ListItems.Count - 1

'Sets The Value To Zero
ProgressBar.Value = 0

'Makes " i " = The Index Of Every Item In The ListToSearch
For i = 1 To ListToSearch.ListItems.Count

'Selects The Item In ListToSearch Depending On The Index
ListToSearch.ListItems.Item(i).Selected = True

'Lets Lenght1 = The Lenght Of The SearchText
length1 = Len(SearchText)

'Lets Lenght2 = The Lenght Of Text In The Selected Item In ListToSearch
length2 = Len(ListToSearch.ListItems.Item(i).Text)

'Lets Lenght3 = Lenght2 - Lenght1
length3 = length2 - length1

'Sets The Maximum Number Of Characters To Search In Text
For x = 0 To length3

'Checks To See If The Text In ListToSearch = SearchText
    If LCase(Mid(ListToSearch.ListItems.Item(i).Text, x + 1, length1)) = LCase(Txtsearch.Text) Then

'Adds a Result to ListToDisplayResults
        ListToDisplayResults.ListItems.Add 1, , ListToSearch.ListItems.Item(i).Text

'This Is To Add All The Sub Items In The Selected Item From ListToSearch Into ListToDisplayResults
        For y = 1 To 4  '<--- All My Subitems Added Up To 4 So You Just Have To Change This Value

'Adds The Subitems In ListToDisplayResults
            ListToDisplayResults.ListItems.Item(1).SubItems(y) = ListToSearch.ListItems.Item(i).SubItems(y)

'Goes To Add The Next Subitem
        Next y

'Prevents This From Adding The Same Item
GoTo Nexti:

'Ends The If Statement
    End If

'Adds +1 To The Value Of The ProgressBar
If ProgressBar.Value < ListToSearch.ListItems.Count - 1 Then ProgressBar.Value = ProgressBar.Value + 1

'Goes To Search Deeper Into The Text Being Searched
Next x

Nexti:
'Goes To Select The Next Item In ListToSearch
Next i

'If No Entries Are Found That Match Then Displays "Search Finished Without Any Results" in ListToDisplayResults
If ListToDisplayResults.ListItems.Count = 0 Then ListToDisplayResults.ListItems.Add , , "Search Finished Without Any Results"

'Sorts The Result List
ListToDisplayResults.Sorted = True

End Function

