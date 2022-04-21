<%@ page language="java"
         import="com.zhuozhengsoft.pageoffice.OpenModeType,com.zhuozhengsoft.pageoffice.PageOfficeCtrl,javax.servlet.*,javax.servlet.http.*"
         pageEncoding="utf-8" %>
<%
    String path = request.getContextPath();
    String filePath = "";
    String id = request.getParameter("id").trim();
    if ("1".equals(id)) {
        filePath = "doc/test1.doc";
    }
    if ("2".equals(id)) {
        filePath = "doc/test2.doc";
    }
    if ("3".equals(id)) {
        filePath = "doc/test3.doc";
    }
    if ("4".equals(id)) {
        filePath = "doc/test4.doc";
    }

    //******************************卓正PageOffice组件的使用*******************************
    PageOfficeCtrl poCtrl1 = new PageOfficeCtrl(request);
    poCtrl1.setServerPage(request.getContextPath() + "/poserver.zz"); //此行必须
    //添加自定义按钮
    poCtrl1.addCustomToolButton("保存", "Save", 1);
    //设置保存页面
    poCtrl1.setSaveFilePage("SaveFile.jsp");

    //打开文件
    poCtrl1.webOpen(filePath, OpenModeType.docNormalEdit, "张三");
%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
    <title>My JSP 'index.jsp' starting page</title>
    <meta http-equiv="pragma" content="no-cache">
    <meta http-equiv="cache-control" content="no-cache">
    <meta http-equiv="expires" content="0">
    <meta http-equiv="keywords" content="keyword1,keyword2,keyword3">
    <meta http-equiv="description" content="This is my page">
    <script type="text/javascript">
        function Save() {
            document.getElementById("PageOfficeCtrl1").WebSave();
        }
    </script>
</head>
<body>
<a href="Default.jsp">返回列表页</a>
<%=poCtrl1.getHtmlCode("PageOfficeCtrl1")%>
</body>
</html>