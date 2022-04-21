<%@ page language="java" import="com.zhuozhengsoft.pageoffice.FileSaver" pageEncoding="utf-8" %>
<%
    FileSaver fs = new FileSaver(request, response);
    String fileName=request.getParameter("fileName");
    //将套红后的文件保存成"zhengshi.doc"
    fs.saveToFile(request.getSession().getServletContext().getRealPath("TaoHong/doc/") + "/" + fileName);
    //fs.showPage(300,300);
    fs.close();
%>

