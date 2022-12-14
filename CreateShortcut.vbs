Function CreateShortcut(path, name, targetPath, Icon, description, workingDirectory, windowStyle)
    set WshShell    = Wscript.CreateObject("Wscript.Shell") 
    set oShellLink  = WshShell.CreateShortcut(path & "\\" & name & ".lnk") '创建一个快捷方式对象
    oShellLink.TargetPath  = targetPath '设置快捷方式的执行路径 
    oShellLink.WindowStyle = windowStyle '运行方式
    oShellLink.IconLocation= Icon '设置快捷方式的图标
    oShellLink.Description = description  '设置快捷方式的描述 
    oShellLink.WorkingDirectory = workingDirectory '起始位置
    oShellLink.Save
End Function

Function Readme()
    data = "您可以将fakeTrays快捷方式放入启动项"
    dim fso
    set fso = Wscript.CreateObject("Scripting.FileSystemObject")
    fso.CreateTextFile("./" & data & ".txt", true).Close()
    set fso = nothing
End Function

Function main()
    thisFolder = createobject("Scripting.FileSystemObject").GetFolder(".").Path '获取当前目录
    Call CreateShortcut(thisFolder, "fakeTrays", thisFolder & "\run.bat", thisFolder & "\fakeTrays.exe,0", "", thisFolder, 7)
    Call CreateShortcut(thisFolder, "StopfakeTrays", thisFolder & "\stop.bat", thisFolder & "\fakeTrays.exe,0", "", thisFolder, 7)
    
    Readme()
End Function

Call main()