<%@ page language="java"
         import="javax.servlet.*,javax.servlet.http.*,java.io.OutputStream,java.sql.Connection,java.sql.DriverManager"
         pageEncoding="utf-8" %>
<%@ page import="java.sql.ResultSet" %>
<%@ page import="java.sql.Statement" %>
<%
    String err = "";
    if (request.getParameter("id") != null
            && request.getParameter("id").trim().length() > 0) {
        String id = request.getParameter("id");
        Class.forName("org.sqlite.JDBC");
        String strUrl = "jdbc:sqlite:"
                + this.getServletContext().getRealPath("demodata/ExaminationPaper.db");
        Connection conn = DriverManager.getConnection(strUrl);
        Statement stmt = conn.createStatement();
        String strSql = "select * from stream where id =" + id;
        ResultSet rs = stmt.executeQuery(strSql);
        if (rs.next()) {
            //******读取磁盘文件，输出文件流 开始*******************************
            byte[] imageBytes = rs.getBytes("Word");
            int fileSize = imageBytes.length;

            response.reset();
            response.setContentType("application/msword"); // application/x-excel, application/ms-powerpoint, application/pdf
            response.setHeader("Content-Disposition", "attachment; filename=down.doc"); //fileN应该是编码后的(utf-8)
            response.setContentLength(fileSize);
            OutputStream outputStream = response.getOutputStream();
            outputStream.write(imageBytes);
            outputStream.flush();
            outputStream.close();
            outputStream = null;
            //下面两句代码解决response.getWriter()和response.getOutputStream()冲突问题
            out.clear();
            out = pageContext.pushBody();
            //******读取磁盘文件，输出文件流 结束*******************************
        } else {
            err = "未获得文件的信息";
        }
        rs.close();
        stmt.close();
        conn.close();
    } else {
        err = "未获得文件的ID";
        //out.print(err);
    }
    if (err.length() > 0)
        err = "<script>alert(" + err + ");</script>";
%>
