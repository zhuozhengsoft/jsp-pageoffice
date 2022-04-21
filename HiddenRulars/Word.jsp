<%@ page language="java"
         import="com.zhuozhengsoft.pageoffice.OpenModeType,com.zhuozhengsoft.pageoffice.PageOfficeCtrl"
         pageEncoding="utf-8" %>
<%
    //******************************卓正PageOffice组件的使用*******************************
    PageOfficeCtrl poCtrl1 = new PageOfficeCtrl(request);
    poCtrl1.setServerPage(request.getContextPath() + "/poserver.zz"); //此行必须
    //隐藏菜单栏
    poCtrl1.setMenubar(false);
    //添加自定义按钮
    poCtrl1.addCustomToolButton("显示/隐藏标尺", "Hidden", 3);
    poCtrl1.webOpen("doc/template.doc", OpenModeType.docNormalEdit, "张三");
%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
    <title>显示/隐藏标尺</title>
    <meta http-equiv="pragma" content="no-cache">
    <meta http-equiv="cache-control" content="no-cache">
    <meta http-equiv="expires" content="0">
    <meta http-equiv="keywords" content="keyword1,keyword2,keyword3">
    <meta http-equiv="description" content="This is my page">
</head>
<body>
<div style=" width:1200px; height:700px;">
    <!-- *********************************PageOffice组件的使用************************************ -->
    <script type="text/javascript" language="javascript">
        //显示/隐藏标尺
        function Hidden() {
            document.getElementById("PageOfficeCtrl1").Document.ActiveWindow.ActivePane.DisplayRulers =
                !document.getElementById("PageOfficeCtrl1").Document.ActiveWindow.ActivePane.DisplayRulers;
        }

    </script>
    <%=poCtrl1.getHtmlCode("PageOfficeCtrl1")%>
    <!-- *********************************PageOffice组件的使用************************************ -->
</div>
</body>
</html>
