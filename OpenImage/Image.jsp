<%@ page language="java" import="com.zhuozhengsoft.pageoffice.PDFCtrl" pageEncoding="utf-8" %>
<%

    PDFCtrl poCtrl1 = new PDFCtrl(request);
    poCtrl1.setServerPage(request.getContextPath() + "/poserver.zz"); //此行必须
    // Create custom toolbar
    poCtrl1.addCustomToolButton("打印", "Print()", 6);
    poCtrl1.addCustomToolButton("-", "", 0);
    poCtrl1.addCustomToolButton("实际大小", "SetPageReal()", 16);
    poCtrl1.addCustomToolButton("适合页面", "SetPageFit()", 17);
    poCtrl1.addCustomToolButton("适合宽度", "SetPageWidth()", 18);
    poCtrl1.addCustomToolButton("-", "", 0);
    poCtrl1.addCustomToolButton("左转", "RotateLeft()", 12);
    poCtrl1.addCustomToolButton("右转", "RotateRight()", 13);
    poCtrl1.addCustomToolButton("-", "", 0);
    poCtrl1.addCustomToolButton("放大", "ZoomIn()", 14);
    poCtrl1.addCustomToolButton("缩小", "ZoomOut()", 15);
    poCtrl1.addCustomToolButton("-", "", 0);
    poCtrl1.addCustomToolButton("全屏", "SwitchFullScreen()", 4);
    poCtrl1.webOpen("doc/test.jpg");
%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
    <style>
        #main {
            width: 1030px;
            height: 900px;
            border: #83b3d9 2px solid;
            background: #f2f7fb;

        }
        #shut {
            width: 45px;
            height: 30px;
            float: right;
            margin-right: -1px;
        }
        #shut:hover {
        }
    </style>
</head>
<body style="overflow:hidden">
<script type="text/javascript">
    function Close() {
        window.external.close();
    }

</script>
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

    function RotateRight() {
        document.getElementById("PDFCtrl1").RotateRight();
    }

    function RotateLeft() {
        document.getElementById("PDFCtrl1").RotateLeft();
    }
</script>
<div id="main">
    <div id="content" style="height:850px;width:1028px;overflow-y:auto;">
        <%=poCtrl1.getHtmlCode("PDFCtrl1")%>
    </div>
</div>
</body>
</html>