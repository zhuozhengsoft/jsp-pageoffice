<%@ page language="java" import="com.zhuozhengsoft.pageoffice.OpenModeType" pageEncoding="utf-8" %>
<%@ page import="com.zhuozhengsoft.pageoffice.PageOfficeCtrl" %>
<%
    PageOfficeCtrl poCtrl1 = new PageOfficeCtrl(request);
    poCtrl1.setServerPage(request.getContextPath() + "/poserver.zz");
    // Create custom toolbar
    poCtrl1.addCustomToolButton("显示A文档", "ShowFile1View()", 0);
    poCtrl1.addCustomToolButton("显示B文档", "ShowFile2View()", 0);
    poCtrl1.addCustomToolButton("显示比较结果", "ShowCompareView()", 0);
    //poCtrl1.wordCompare("doc/aaa1.doc", "doc/aaa2.doc", OpenModeType.docReadOnly, "张三");
    poCtrl1.wordCompare("doc/aaa1.doc", "doc/aaa2.doc", OpenModeType.docAdmin, "张三");
%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
    <title>Word文档比较</title>
    <meta http-equiv="pragma" content="no-cache">
    <meta http-equiv="cache-control" content="no-cache">
    <meta http-equiv="expires" content="0">
    <meta http-equiv="keywords" content="keyword1,keyword2,keyword3">
    <meta http-equiv="description" content="This is my page">
</head>
<body>
<script language="javascript" type="text/javascript">
    function ShowFile1View() {
        document.getElementById("PageOfficeCtrl1").Document.ActiveWindow.View.ShowRevisionsAndComments = false;
        document.getElementById("PageOfficeCtrl1").Document.ActiveWindow.View.RevisionsView = 1;
    }

    function ShowFile2View() {
        document.getElementById("PageOfficeCtrl1").Document.ActiveWindow.View.ShowRevisionsAndComments = false;
        document.getElementById("PageOfficeCtrl1").Document.ActiveWindow.View.RevisionsView = 0;
    }

    function ShowCompareView() {
        document.getElementById("PageOfficeCtrl1").Document.ActiveWindow.View.ShowRevisionsAndComments = true;
        document.getElementById("PageOfficeCtrl1").Document.ActiveWindow.View.RevisionsView = 0;
    }

    function SetFullScreen() {
        document.getElementById("PageOfficeCtrl1").FullScreen = !document.getElementById("PageOfficeCtrl1").FullScreen;
    }
</script>
<div style="width:1000px; height:800px;">
    <%=poCtrl1.getHtmlCode("PageOfficeCtrl1")%>
</div>
</body>
</html>
