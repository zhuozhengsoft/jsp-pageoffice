<%@ page language="java"
         import="com.zhuozhengsoft.pageoffice.FileSaver,java.sql.Connection,java.sql.DriverManager"
         pageEncoding="utf-8" %>
<%@ page import="java.sql.PreparedStatement" %>
<%
    FileSaver fs = new FileSaver(request, response);
    String err = "";
    if (request.getParameter("id") != null
            && request.getParameter("id").trim().length() > 0) {
        String id = request.getParameter("id").trim();
        Class.forName("org.sqlite.JDBC");
        String strUrl = "jdbc:sqlite:"
                + this.getServletContext().getRealPath("demodata/DataBase.db");
        Connection conn = DriverManager.getConnection(strUrl);
        String sql = "UPDATE  Stream SET Word=?  where ID=" + id;
        PreparedStatement pstmt = null;
        pstmt = conn.prepareStatement(sql);
        pstmt.setBytes(1, fs.getFileBytes());
        //pstmt.setBinaryStream(1, fs.getFileStream(),fs.getFileSize());
        pstmt.executeUpdate();
        pstmt.close();
        conn.close();

        fs.setCustomSaveResult("ok");
    } else {
        err = "<script>alert('未获得文件的ID，保存失败');</script>";
    }
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
<%=err%>
</body>
</html>
