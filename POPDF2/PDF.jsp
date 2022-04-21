<%@ page language="java" import="com.zhuozhengsoft.pageoffice.PDFCtrl" pageEncoding="utf-8" %>
<%

    PDFCtrl pdfCtrl1 = new PDFCtrl(request);
    pdfCtrl1.setServerPage(request.getContextPath() + "/poserver.zz"); //此行必须
    // Create custom toolbar
    pdfCtrl1.addCustomToolButton("打印", "PrintFile()", 6);
    pdfCtrl1.addCustomToolButton("隐藏/显示书签", "SetBookmarks()", 0);
    pdfCtrl1.addCustomToolButton("-", "", 0);
    pdfCtrl1.addCustomToolButton("实际大小", "SetPageReal()", 16);
    pdfCtrl1.addCustomToolButton("适合页面", "SetPageFit()", 17);
    pdfCtrl1.addCustomToolButton("适合宽度", "SetPageWidth()", 18);
    pdfCtrl1.addCustomToolButton("-", "", 0);
    pdfCtrl1.addCustomToolButton("首页", "FirstPage()", 8);
    pdfCtrl1.addCustomToolButton("上一页", "PreviousPage()", 9);
    pdfCtrl1.addCustomToolButton("下一页", "NextPage()", 10);
    pdfCtrl1.addCustomToolButton("尾页", "LastPage()", 11);
    pdfCtrl1.addCustomToolButton("-", "", 0);
    pdfCtrl1.addCustomToolButton("向左旋转90度", "SetRotateLeft()", 12);
    pdfCtrl1.addCustomToolButton("向右旋转90度", "SetRotateRight()", 13);
    pdfCtrl1.webOpen("doc/test.pdf");
%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
</head>
<body style="overflow:hidden">
<!--**************   卓正 PageOffice 客户端代码开始    ************************-->
<script language="javascript" type="text/javascript">
    function AfterDocumentOpened() {
        //alert(document.getElementById("PDFCtrl1").Caption);
    }

    function SetBookmarks() {
        document.getElementById("PDFCtrl1").BookmarksVisible = !document.getElementById("PDFCtrl1").BookmarksVisible;
    }

    function PrintFile() {
        document.getElementById("PDFCtrl1").ShowDialog(4);
    }

    function SwitchFullScreen() {
        document.getElementById("PDFCtrl1").FullScreen = !document.getElementById("PDFCtrl1").FullScreen;
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

    function ZoomIn() {
        document.getElementById("PDFCtrl1").ZoomIn();
    }

    function ZoomOut() {
        document.getElementById("PDFCtrl1").ZoomOut();
    }

    function FirstPage() {
        document.getElementById("PDFCtrl1").GoToFirstPage();
    }

    function PreviousPage() {
        document.getElementById("PDFCtrl1").GoToPreviousPage();
    }

    function NextPage() {
        document.getElementById("PDFCtrl1").GoToNextPage();
    }

    function LastPage() {
        document.getElementById("PDFCtrl1").GoToLastPage();
    }

    function SetRotateRight() {
        document.getElementById("PDFCtrl1").RotateRight();
    }

    function SetRotateLeft() {
        document.getElementById("PDFCtrl1").RotateLeft();
    }
</script>
<div style="height:850px;width:auto;">
    <%=pdfCtrl1.getHtmlCode("PDFCtrl1")%>
</div>
</body>
</html>