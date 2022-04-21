<%@ page language="java"
         import="com.zhuozhengsoft.pageoffice.DocumentOpenType,com.zhuozhengsoft.pageoffice.FileMakerCtrl, com.zhuozhengsoft.pageoffice.wordwriter.WordDocument"
         pageEncoding="utf-8" %>
<%
    FileMakerCtrl fmCtrl = new FileMakerCtrl(request);
    fmCtrl.setServerPage(request.getContextPath() + "/poserver.zz");
    String id = request.getParameter("id");
    if (id != null && id.length() > 0) {
        WordDocument doc = new WordDocument();
        //给数据区域赋值，即把数据填充到模板中相应的位置
        doc.openDataRegion("PO_company").setValue("北京卓正志远软件有限公司  " + id);
        fmCtrl.setSaveFilePage("Save.jsp?id=" + id);
        fmCtrl.setWriter(doc);
        fmCtrl.setJsFunction_OnProgressComplete("OnProgressComplete()");
        fmCtrl.fillDocument("doc/template.doc", DocumentOpenType.Word);
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
            document.getElementById("FileMakerCtrl1").PrintOut();
            window.parent.myFunc(); //调用父页面的js函数
        }
    </script>
    <!--**************   卓正 PageOffice 客户端代码结束    ************************-->
    <%=fmCtrl.getHtmlCode("FileMakerCtrl1")%>
</div>
</body>
</html>
