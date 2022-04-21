<%@ page contentType="text/html; charset=utf-8" language="java" import="java.util.*,java.io.*,java.net.*"  %>
<%@ page pageEncoding="utf-8"%>
<%
    String filePath =request.getSession().getServletContext().getRealPath("FileMakerSingle/doc/maker.doc");
    int fileSize =(int)new File(filePath).length();

    response.reset();
    response.setContentType("application/msword"); // application/x-excel, application/ms-powerpoint, application/pdf
    response.setHeader("Content-Disposition", "attachment; filename=maker.doc");
    response.setContentLength(fileSize);

    OutputStream outputStream = response.getOutputStream();
    InputStream inputStream = new FileInputStream(filePath);
    byte[] buffer = new byte[10240];
    int i = -1;
    while ((i = inputStream.read(buffer)) != -1) {
        outputStream.write(buffer, 0, i);
    }

    outputStream.flush();
    outputStream.close();
    inputStream.close();

%>

