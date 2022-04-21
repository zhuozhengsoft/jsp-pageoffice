<%@ page language="java" import="com.zhuozhengsoft.pageoffice.FileSaver" pageEncoding="UTF-8" %>
<%
    FileSaver fs = new FileSaver(request, response);
    fs.saveToFile(request.getSession().getServletContext().getRealPath("InsertSeal/Excel/AddSign1/") + "/" + fs.getFileName());
    fs.close();
%>


