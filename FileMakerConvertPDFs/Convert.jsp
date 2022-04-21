<%@ page language="java"
         import="com.zhuozhengsoft.pageoffice.DocumentOpenType,com.zhuozhengsoft.pageoffice.FileMakerCtrl"
         pageEncoding="utf-8" %>
<%
    String filePath = request.getSession().getServletContext().getRealPath("FileMakerConvertPDFs/doc/");
    String id = request.getParameter("id").trim();

    if ("1".equals(id)) {
        filePath = request.getSession().getServletContext().getRealPath("FileMakerConvertPDFs/doc/PageOffice产品简介.doc");
    }
    if ("2".equals(id)) {
        filePath = request.getSession().getServletContext().getRealPath("FileMakerConvertPDFs/doc/Pageoffice客户端安装步骤.doc");
    }
    if ("3".equals(id)) {
        filePath = request.getSession().getServletContext().getRealPath("FileMakerConvertPDFs/doc/PageOffice的应用领域.doc");
    }
    if ("4".equals(id)) {
        filePath = request.getSession().getServletContext().getRealPath("FileMakerConvertPDFs/doc/PageOffice产品对客户端环境要求.doc");
    }

    FileMakerCtrl fmCtrl = new FileMakerCtrl(request);
    fmCtrl.setServerPage(request.getContextPath() + "/poserver.zz");
    fmCtrl.setJsFunction_OnProgressComplete("OnProgressComplete()");
    fmCtrl.setSaveFilePage("SaveFile.jsp");
     //打开文件
    //这里用磁盘路径方式打开文件，为了支持linx服务器，文档路径前面必须加上“file://”前缀
    if(System.getProperty("os.name").toLowerCase().contains("linux")){    //判断该当前系统是否是linux操作系统
       fmCtrl.fillDocumentAsPDF("file://"+filePath, DocumentOpenType.Word, "a.pdf");
    }else{
       fmCtrl.fillDocumentAsPDF(filePath, DocumentOpenType.Word, "a.pdf");
    }
%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
    <title>My JSP 'FileMaker.jsp' starting page</title>
    <meta http-equiv="pragma" content="no-cache">
    <meta http-equiv="cache-control" content="no-cache">
    <meta http-equiv="expires" content="0">
    <meta http-equiv="keywords" content="keyword1,keyword2,keyword3">
    <meta http-equiv="description" content="This is my page">
    <!--
<link rel="stylesheet" type="text/css" href="styles.css">
-->
</head>
<body>
<div>
    <!--**************   卓正 PageOffice 客户端代码开始    ************************-->
    <script language="javascript" type="text/javascript">
        function OnProgressComplete() {
            window.parent.convert(); //调用父页面的js函数
        }
    </script>
    <!--**************   卓正 PageOffice 客户端代码结束    ************************-->
    <%=fmCtrl.getHtmlCode("FileMakerCtrl1")%>
</div>
</body>
</html>