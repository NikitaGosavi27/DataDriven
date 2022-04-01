Dim objuft

Set objuft=CreateObject("QuickTest.Application")
objuft.visible=True
objuft.launch
objuft.open("D:\DataDrivenFramework\Driver\Driver")
objuft.Test.Run
objuft.Test.Close
objuft.quit
set objuft=nothing