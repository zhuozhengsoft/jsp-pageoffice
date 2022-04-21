<%@ page language="java"
         import="com.zhuozhengsoft.pageoffice.FileSaver"
         pageEncoding="utf-8" %>
<%
    FileSaver fs = new FileSaver(request, response);
    String  pdfName=request.getParameter("pdfName");
    fs.saveToFile(request.getSession().getServletContext().getRealPath("FileMakerPDF/doc") + "/" + pdfName);
    fs.close();
%>

