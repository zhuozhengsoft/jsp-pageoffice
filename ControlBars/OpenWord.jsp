<%@ page language="java"
         import="com.zhuozhengsoft.pageoffice.OpenModeType,com.zhuozhengsoft.pageoffice.PageOfficeCtrl"
         pageEncoding="utf-8" %>
<%
    //******************************卓正PageOffice组件的使用*******************************
    //设置PageOffice服务器组件
    PageOfficeCtrl poCtrl1 = new PageOfficeCtrl(request);
    poCtrl1.setServerPage(request.getContextPath() + "/poserver.zz"); //此行必须
    poCtrl1.setTitlebar(false); //隐藏标题栏
    poCtrl1.setMenubar(false); //隐藏菜单栏
    poCtrl1.setOfficeToolbars(false);//隐藏Office工具条
    poCtrl1.setCustomToolbar(false);//隐藏自定义工具栏
    //打开文件
    //poCtrl1.webOpen("doc/template.doc", OpenModeType.docNormalEdit, "张三");
    poCtrl1.webOpen("doc/template.doc", OpenModeType.docReadOnly, "张三");
%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
    <title>控制标题栏、菜单栏、自定工具栏和Office工具栏的隐藏和显示</title>
    <meta http-equiv="pragma" content="no-cache">
    <meta http-equiv="cache-control" content="no-cache">
    <meta http-equiv="expires" content="0">
    <meta http-equiv="keywords" content="keyword1,keyword2,keyword3">
    <meta http-equiv="description" content="This is my page">
    <!--
    <link rel="stylesheet" type="text/css" href="styles.css">
    -->
</head>
<body>
隐藏了标题栏、菜单栏、自定工具栏和Office工具栏的效果，每个栏都是可以单独的控制是否隐藏。
<div style=" width:auto; height:700px;">
    <%=poCtrl1.getHtmlCode("PageOfficeCtrl1")%>
</div>
</body>
</html>
