<%@ page language="java" import="com.zhuozhengsoft.pageoffice.PDFCtrl" pageEncoding="utf-8" %>
<%
    PDFCtrl poCtrl1 = new PDFCtrl(request);
    poCtrl1.setServerPage(request.getContextPath() + "/poserver.zz"); //此行必须
    // Create custom toolbar
    poCtrl1.addCustomToolButton("搜索", "SearchText()", 0);
    poCtrl1.addCustomToolButton("搜索下一个", "SearchTextNext()", 0);
    poCtrl1.addCustomToolButton("搜索上一个", "SearchTextPrev()", 0);
    poCtrl1.addCustomToolButton("实际大小", "SetPageReal()", 16);
    poCtrl1.addCustomToolButton("适合页面", "SetPageFit()", 17);
    poCtrl1.addCustomToolButton("适合宽度", "SetPageWidth()", 18);
    poCtrl1.webOpen("doc/test.pdf");
%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
    <title>PDF文档中的关键字搜索</title>
    <meta http-equiv="pragma" content="no-cache">
    <meta http-equiv="cache-control" content="no-cache">
    <meta http-equiv="expires" content="0">
    <meta http-equiv="keywords" content="keyword1,keyword2,keyword3">
    <meta http-equiv="description" content="This is my page">
</head>
<body>
<!--**************   卓正 PageOffice 客户端代码开始    ************************-->
<script language="javascript" type="text/javascript">
    function SearchText() {
        document.getElementById("PDFCtrl1").SearchText();
    }

    function SearchTextNext() {
        document.getElementById("PDFCtrl1").SearchTextNext();
    }

    function SearchTextPrev() {
        document.getElementById("PDFCtrl1").SearchTextPrev();
    }

    function SetPageReal() {
        document.getElementById("PDFCtrl1").SetPageFit(1);
    }

    function SetPageFit() {
        document.getElementById("PDFCtrl1").SetPageFit(2);
    }

    function SetPageWidth() {
        document.getElementById("PDFCtrl1").SetPageFit(3);
    }
</script>
<!--**************   卓正 PageOffice 客户端代码结束    ************************-->
<div style="width:auto; height:600px;">
    <%=poCtrl1.getHtmlCode("PDFCtrl1")%>
</div>
</body>
</html>