<%@ page language="java"
         import="com.zhuozhengsoft.pageoffice.FileSaver"
         pageEncoding="utf-8" %>
<%
    FileSaver fs = new FileSaver(request, response);
    String id = request.getParameter("id");
    String err = "";
    if (id != null && id.length() > 0) {
        String fileName = "\\maker" + id + fs.getFileExtName();
        fs.saveToFile(request.getSession().getServletContext()
                .getRealPath("FileMaker/doc")
                + fileName);
    } else {
        err = "<script>alert('未获得文件名称');</script>";
    }
    fs.close();
%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
    <title>My JSP 'SaveMaker.jsp' starting page</title>
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
<%=err%>
</body>
</html>
