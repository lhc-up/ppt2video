' 命令行执行, cscript //nologo fkpdf.vbs input.pptx output.pdf
' cscript //nologo E:\CPPT\nodePPT\fkpdf.vbs D:\Desktop\简洁大气红色实用年终总结计划.pptx D:\Desktop\finalfff.pdf
Set args = WScript.Arguments
' 两个参数，输入路径和输出路径
if args.count = 2 then
    dim pptPath, pdfPath
    pptPath = args(0)
    pdfPath = args(1)
    ' res = MsgBox(pptPath, 4+32+256, "fkkk")
    ' res = MsgBox(pdfPath, 4+32+256, "fkkk")
    ppt2pdf pptPath, pdfPath
end if

function ppt2pdf(pptPath, pdfPath)
    if isExistFile(pptPath) then
        Set pptApp = WScript.CreateObject("PowerPoint.Application")
        ' 只读模式，不显示窗口
        Set oppt = pptApp.Presentations.open(pptPath, true, false, false)
        ' PpSaveAsFileType enumeration (PowerPoint)
        ' https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppsaveasfiletype
        ppSaveAsPDF = 32
        ppSaveAsPNG = 18
        oppt.SaveAs pdfPath, ppSaveAsPNG, false
        oppt.Close
        pptApp.Quit
        Set pptApp = Nothing
        Set oppt = Nothing
    end if
end function

' 文件是否存在
function isExistFile(path)
    dim fso
    set fso = CreateObject("Scripting.FileSystemObject")
    if fso.FileExists(path) then
        isExistFile = true
    else 
        isExistFile = false
    end if
end function
