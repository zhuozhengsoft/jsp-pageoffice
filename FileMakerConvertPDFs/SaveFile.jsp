<%@ page language="java" import="com.zhuozhengsoft.pageoffice.FileSaver" contentType="text/html;charset=gb2312" %>
<%
    FileSaver fs = new FileSaver(request, response);
    fs.saveToFile(request.getSession().getServletContext().getRealPath("FileMakerConvertPDFs/doc/") + "/" + fs.getFileName());
    fs.close();
%>

