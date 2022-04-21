<%@ page language="java" import="com.zhuozhengsoft.pageoffice.FileSaver" pageEncoding="utf-8" %>
<%
    response.setCharacterEncoding("UTF-8");
    //获取token值
    String mytoken=request.getHeader("Authorization");
    out.println("文件保存页面的token的值为："+mytoken);
    
    FileSaver fs = new FileSaver(request, response);
    fs.saveToFile(request.getSession().getServletContext().getRealPath("Token/doc/") + "/" + fs.getFileName());
    fs.showPage(300,200);
    fs.close();
%>


