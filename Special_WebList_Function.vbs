fnWebList "select"
Function fnWebList (vName)
   Set wList = Description.Create()
   wList("micclass").value = "WebList"
   set obj = Browser("micclass:=Browser").Page("micclass:=Page").ChildObjects(wList)
   For i = 0 to obj.count - 1
	   If obj(i).GetRoProperty("Name") = vName Then
		    vList = obj(i).GetROProperty("all items")
			vInput = Inputbox (vList)
			vSplit = Split(vList,";")
			For j = 0 to ubound(vSplit)
			If vInput = vSplit(j) Then
				obj(i).Select vInput
				Exit for 
			End If
Next
		   Exit for 
	   End If

   Next
End Function