<%@ page language="java"
         import="com.zhuozhengsoft.pageoffice.OpenModeType,com.zhuozhengsoft.pageoffice.PageOfficeCtrl"
         pageEncoding="utf-8" %>
<%
    //******************************卓正PageOffice组件的使用*******************************
    PageOfficeCtrl poCtrl1 = new PageOfficeCtrl(request);
    poCtrl1.setServerPage(request.getContextPath() + "/poserver.zz"); //此行必须
    //隐藏菜单栏
    poCtrl1.setMenubar(false);
    poCtrl1.addCustomToolButton("保存", "Save()", 1);
    //设置保存页面
    String id = request.getParameter("id");
    poCtrl1.setSaveFilePage("SaveFile.jsp?id=" + id);
    //打开Word文件
    poCtrl1.webOpen("Openfile.jsp?id=" + id, OpenModeType.docNormalEdit, "张三");
%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
    <title>My JSP 'Edit.jsp' starting page</title>
    <meta http-equiv="pragma" content="no-cache">
    <meta http-equiv="cache-control" content="no-cache">
    <meta http-equiv="expires" content="0">
    <meta http-equiv="keywords" content="keyword1,keyword2,keyword3">
    <meta http-equiv="description" content="This is my page">
</head>
<body>
<a href="#" onclick="window.external.close();">返回列表页</a>
<div style="width: auto; height: 700px;">
    <!-- *************************PageOffice组件的使用******************************** -->
    <script type="text/javascript">
        function Save() {
            document.getElementById("PageOfficeCtrl1").WebSave();
        }
    </script>
    <%=poCtrl1.getHtmlCode("PageOfficeCtrl1")%>
    <!-- *************************PageOffice组件的使用******************************** -->
</div>
</body>
</html>
