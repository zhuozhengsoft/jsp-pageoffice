<%@ page language="java" import="com.zhuozhengsoft.pageoffice.FileSaver" pageEncoding="utf-8" %>
<%
    FileSaver fs = new FileSaver(request, response);
    fs.saveToFile(request.getSession().getServletContext().getRealPath("SimpleExcel2/doc/") + "/" + fs.getFileName());
    fs.close();
%>