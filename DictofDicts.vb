'note the .vb is just so github can see what the code is
'this can be set up as a function or a sub
'at work I initialize the dictionaires in my main sub then pass them to be built to the sub below, i'll start from scratch here
'I use this when i'm trying to group together items for a bigger "item",
Sub DictofDicts()
	
	Dim DictofDicts As New Scripting.Dictionary
	Dim Dict As New Scripting.Dictionary
	
	Dim Tbl As Excel.ListObject
	
	Dim Ar As Variant
	Dim Something As Variant 'just a place holder, could be string, date, array, or another dict if you really wanna do that to yourself
	
	Dim KeyString As String 'key of outer dict
	Dim ItemString As String 'key of inner dict
	
	Set Tbl = ThisWorkbook.Worksheets("Shipments").ListObjects(1)
		
	Ar = Tbl.DatabodyRange
	
	For i = 1 to UBound(Ar, 1)
		
		KeyString = Ar(i, 1) 'whatever your key string is
		Item = Ar(i, 2) 'whatever you want the item to be
		Something = Ar(i, 3) 'again this can be whatever you want but its the "lowest" level of the data
		
		If Not DictofDicts.Exists(KeyString) Then
			
			'do stuff if it doesn't already exist
			'in my cases the items are unique to the KeyStrings, if that is not the case for you, then you should add anothe layer of If Not .Exists for your Dict
			Set Dict(Item) = Something
			
			DictofDicts(KeyString) = Dict
			
		Else
			
				'need to give the inner dict a name first, and get the content that already exists
				Set Dict = DictofDicts(KeyString)
			
				Dict(Item) = Something
				DictofDicts(KeyString) = Dict
					
						
		End If
		
			'need to clear Dict, otherwise it will wipe out what we just added
				
			Set Dict As New Dictionary
				
		
	Next i
	
	
End Sub
