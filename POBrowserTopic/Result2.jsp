<%@ page language="java" pageEncoding="utf-8" contentType="text/html; charset=GB2312" %>
<%
    response.setContentType("text/html;charset=gb2312");
    String paramValue = request.getParameter("param");
    session.setAttribute("txt", paramValue);
%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
    <title>My JSP 'result.jsp' starting page</title>
    <meta http-equiv="pragma" content="no-cache">
    <meta http-equiv="cache-control" content="no-cache">
    <meta http-equiv="expires" content="0">
    <meta http-equiv="keywords" content="keyword1,keyword2,keyword3">
    <meta http-equiv="description" content="This is my page">
    <script type="text/javascript"
            src="jquery-1.6.min.js"></script>
    <script type="text/javascript">
    </script>
</head>
<body>
</body>
</html>
