<%@ page language="java"
         import="java.io.OutputStream,java.sql.Connection,java.sql.DriverManager,java.sql.ResultSet,java.sql.Statement"
         pageEncoding="utf-8" %>
<%
    String id = "2";
    if (request.getParameter("id") != null
            && request.getParameter("id").trim().length() > 0) {
        id = request.getParameter("id");
    }
    Class.forName("org.sqlite.JDBC");
    String strUrl = "jdbc:sqlite:"
            + this.getServletContext().getRealPath("demodata/DataBase.db");
    Connection conn = DriverManager.getConnection(strUrl);
    Statement stmt = conn.createStatement();
    ResultSet rs = stmt.executeQuery("select * from stream where id = "
            + id);
    int newID = 1;
    if (rs.next()) {
        //******读取磁盘文件，输出文件流 开始*******************************
        byte[] imageBytes = rs.getBytes("Word");
        int fileSize = imageBytes.length;

        response.reset();
        response.setContentType("application/msword"); // application/x-excel, application/ms-powerpoint, application/pdf
        response.setHeader("Content-Disposition",
                "attachment; filename=down.doc"); //fileN应该是编码后的(utf-8)
        response.setContentLength(fileSize);

        OutputStream outputStream = response.getOutputStream();
        outputStream.write(imageBytes);

        outputStream.flush();
        outputStream.close();
        outputStream = null;
        //******读取磁盘文件，输出文件流 结束*******************************
    }
    rs.close();
    conn.close();
%>

