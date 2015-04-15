Set fso = CreateObject("Scripting.FileSystemObject")
Set fd = fso.OpenTextFile("items.zof",8)
id = inputbox("Item ID","Enter ID for new item")
itemname = inputbox("Item Name","Enter Name for new Item")
price = inputbox("Price?","Enter price for new item.")
fd.writeLine id
fd.writeLine itemname
fd.writeLine price
fd.close
msgbox "Done.",vbInformation,"Done adding item."