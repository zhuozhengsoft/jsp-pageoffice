<%@ page language="java" import="com.zhuozhengsoft.pageoffice.FileSaver" pageEncoding="utf-8" %>
<%
    FileSaver fs = new FileSaver(request, response);
    fs.saveToFile(request.getSession().getServletContext().getRealPath("SaveReturnValue/doc") + "/" + fs.getFileName());
    fs.setCustomSaveResult("ok");//设置保存结果
    fs.close();
%>

