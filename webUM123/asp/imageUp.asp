<%@ LANGUAGE="VBSCRIPT" CODEPAGE="936" %> 
<!--#include file="Uploader.Class.asp"-->
<!--#include file="json.asp"-->

<%
    'Author: techird
    'Date: 2013/09/29

    '配置
    'MAX_SIZE 在这里设定了之后如果出现大上传失败，请执行以下步骤
    'IIS 6 
        '找到位于 C:\Windows\System32\Inetsrv 中的 metabase.XML 打开，找到ASPMaxRequestEntityAllowed 把他修改为需要的值（如10240000即10M）
    'IIS 7
        '打开IIS控制台，选择 ASP，在限制属性里有一个“最大请求实体主题限制”，设置需要的值

    Dim up, json, path, callback

    Set up = new Uploader
    up.MaxSize = 10 * 1024 * 1024
    up.AllowType = Array(".gif", ".png", ".jpg", ".jpeg", ".bmp")
    up.ProcessForm()

    up.FileField = "upfile"
    up.SavePath = "upload/"
    up.SaveFile()

    Session.CodePage = 936
    Response.AddHeader "Content-Type", "text/html;charset=gbk"
    SetLocale 2052

    Set json = jsObject()
    json("originalName") = up.OriginalFileName
    json("name") = up.FileName
    json("url") = up.FilePath
    json("size") = up.FileSize
    json("state") = up.State
    json("type") = up.FileType

    callback = up.FormValues.Item("callback")

    If IsEmpty( callback ) Then
        Response.Write json.jsString()
    Else
        Response.Write "<script>" & callback & "( JSON.parse(" & json.jsString() & "));</script>"
    End If

    Function IsInArray(arr, elem)
        IsInArray = false
        For Each i In arr
            If i = elem Then IsInArray = true
        Next
    End Function
%>