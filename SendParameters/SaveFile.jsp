<%@ page language="java"
         import="com.zhuozhengsoft.pageoffice.FileSaver,java.sql.Connection,java.sql.DriverManager"
         pageEncoding="utf-8" %>
<%@ page import="java.sql.Statement" %>
<%
    FileSaver fs = new FileSaver(request, response);
    fs.saveToFile(request.getSession().getServletContext().getRealPath("SendParameters/doc") + "/" + fs.getFileName());
    int id = 0;
    String userName = "";
    int age = 0;
    String sex = "";

    //获取通过Url传递过来的值
    if (request.getParameter("id") != null
            && request.getParameter("id").trim().length() > 0)
        id = Integer.parseInt(request.getParameter("id").trim());

    //获取通过网页标签控件传递过来的参数值，注意fs.getFormField("参数名")方法中的参数名是值标签的Id

    //获取通过文本框<input type="text" />标签传递过来的值
    if (fs.getFormField("userName") != null
            && fs.getFormField("userName").trim().length() > 0) {
        userName = fs.getFormField("userName");
    }

    //获取通过隐藏域传递过来的值
    if (fs.getFormField("age") != null
            && fs.getFormField("age").trim().length() > 0) {
        age = Integer.parseInt(fs.getFormField("age"));
    }

    //获取通过<select>标签传递过来的值
    if (fs.getFormField("selSex") != null
            && fs.getFormField("selSex").trim().length() > 0) {
        sex = fs.getFormField("selSex");
    }
    Class.forName("org.sqlite.JDBC");//载入驱动程序类别
    String strUrl = "jdbc:sqlite:"
            + this.getServletContext().getRealPath("demodata/SendParameters.db");
    Connection conn = DriverManager.getConnection(strUrl);
    Statement stmt = conn.createStatement();
    String strsql = "Update Users set UserName = '" + userName
            + "', age = " + age + ",sex = '" + sex + "' where id = "
            + id;
    stmt.executeUpdate(strsql);
    conn.close();
    fs.showPage(300, 200);
    fs.close();
%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
    <title>My JSP 'SaveFile.jsp' starting page</title>
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
<div>
    传递的参数为：<br/>
    id:<%=id %><br/>
    userName:<%=userName%><br/>
    age:<%=age%><br/>
    sex:<%=sex%><br/>
</div>
</body>
</html>
