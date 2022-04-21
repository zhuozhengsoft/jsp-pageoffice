<%@ page language="java" import="com.zhuozhengsoft.pageoffice.FileSaver" pageEncoding="utf-8" %>
<%
    FileSaver fs = new FileSaver(request, response);
    //String aa=fs.getFileExtName();
    if (fs.getFileExtName().equals(".jpg")) {
        fs.saveToFile(request.getSession().getServletContext().getRealPath("SaveFirstPageAsImg/images") + "/" + fs.getFileName());
    } else {
        fs.saveToFile(request.getSession().getServletContext().getRealPath("SaveFirstPageAsImg/doc") + "/" + fs.getFileName());
    }
    fs.setCustomSaveResult("ok");
    fs.close();
%>

