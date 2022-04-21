<%@ page language="java" import="com.zhuozhengsoft.pageoffice.FileSaver" pageEncoding="utf-8" %>
<%
    FileSaver fs = new FileSaver(request, response);
    fs.saveToFile(request.getSession().getServletContext().getRealPath("DataRegionEdit/doc/") + "/" + fs.getFileName());
    //fs.showPage(300,300);
    fs.close();
%>
