<%@ page language="java"
         import="com.zhuozhengsoft.pageoffice.OpenModeType,com.zhuozhengsoft.pageoffice.PageOfficeCtrl,javax.servlet.*,javax.servlet.http.*"
         pageEncoding="utf-8" %>
<%
    String filePath = "";
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
    //******************************卓正PageOffice组件的使用*******************************
    PageOfficeCtrl poCtrl1 = new PageOfficeCtrl(request);
    poCtrl1.setServerPage(request.getContextPath() + "/poserver.zz"); //此行必须
    poCtrl1.setSaveFilePage("SaveFile.jsp");//如要保存文件，此行必须
    poCtrl1.addCustomToolButton("保存", "Save()", 1);//添加自定义工具栏按钮
    //打开文件
    //这里打开的是中文名称的文件，所以必须用磁盘路径方式打开文件，为了支持linx服务器，文档路径前面必须加上“file://”前缀
    if(System.getProperty("os.name").toLowerCase().contains("linux")){    //判断该当前系统是否是linux操作系统
       poCtrl1.webOpen("file://"+filePath, OpenModeType.docNormalEdit, "张三");
    }else{
      poCtrl1.webOpen(filePath, OpenModeType.docNormalEdit, "张三");
    }
 
%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
    <title>My JSP 'index.jsp' starting page</title>
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
<a href="index.jsp">返回列表页</a>
<script type="text/javascript">
    function Save() {
        document.getElementById("PageOfficeCtrl1").WebSave();
    }
</script>
<%=poCtrl1.getHtmlCode("PageOfficeCtrl1")%>
</body>
</html>
