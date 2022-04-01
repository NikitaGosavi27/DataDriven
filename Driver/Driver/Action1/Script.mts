Services.StartTransaction "start"
'DataTable.AddSheet "Module"
'DataTable.ImportSheet "C:\Users\All UFT\KeyWordDrivenFramework\Organizer\organizer.xlsx",1,"Module"
mrowcount=datatable.GetSheet("Action1").GetRowCount
'msgbox mrowcount

For i = 1 To mrowcount Step 1
	DataTable.SetCurrentRow(i)
	
	ModuleExe=DataTable("Moduleexe","Action1")
	
	'msgbox ModuleExe
	
	If ModuleExe="Y" Then
	
		ModuleID=DataTable("ModuleID","Action1")
		
		'msgbox ModuleID
		
		trowcount=datatable.GetSheet("Action2").GetRowCount
		
		'msgbox trowcount
		
		For j = 1 To trowcount Step 1
			DataTable.SetCurrentRow(j)
			
			If ModuleID=DataTable("ModuleID","Action2") and DataTable("Testcaseexe","Action2")="Y" Then
			
			testcaseId=DataTable("TestcaseId","Action2")
			
			'msgbox testcaseId
			
			tsrowcount=DataTable.GetSheet("Action3").GetRowCount
			'msgbox tsrowcount
			
			For k = 1 To tsrowcount Step 1
				
				DataTable.SetCurrentRow(k)
				
				If testcaseId=DataTable("TestcaseId","Action3") Then
				
				keyword=datatable("Keyword","Action3")
			'	msgbox keyword
				
				Select Case (Keyword)	
				
				case "LN"
				Call Login("John","HP")
				
				Case "oo"
				Call openorder()
				
				Case "ua"
				Call Updateorder()
				
				Case "ca"
				Call Closeapp()
				
				Case "Lnd"
				'Call Login_datadriven()
				
				drowcount=datatable.getsheet("Action4").GetRowCount
				For L = 1 To drowcount Step 1
					
					datatable.setcurrentrow(L)
					
					Call login(datatable("username", "Action4"),datatable("password","Action4"))
					
					Call Closeapp()
				Next
				
				Case "ood"
				
				orrowcount=datatable.getsheet("Action4").GetRowCount
				For m = 1 To orrowcount Step 1
					datatable.SetCurrentRow(m)
					Call openorder(datatable("orderno","Action4"))
				Next
				
				End Select
				End If
			Next
			End If
		Next
	End If
	
	
Next

Services.EndTransaction "start" @@ hightlight id_;_1379198_;_script infofile_;_ZIP::ssf18.xml_;_

 @@ hightlight id_;_1928294544_;_script infofile_;_ZIP::ssf33.xml_;_
 @@ hightlight id_;_2036219320_;_script infofile_;_ZIP::ssf34.xml_;_
