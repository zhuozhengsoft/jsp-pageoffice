<%@ page language="java"
         import="com.zhuozhengsoft.pageoffice.FileSaver"
         pageEncoding="utf-8" %>
<%
    FileSaver fs = new FileSaver(request, response);
    String id = request.getParameter("id");
    String err = "";
    if (id != null && id.length() > 0) {
        String fileName = "\\maker" + id + fs.getFileExtName();
        fs.saveToFile(request.getSession().getServletContext()
                .getRealPath("PrintFiles/doc")
                + fileName);
    } else {
        err = "<script>alert('未获得文件名称');</script>";
    }
    fs.close();
%>
