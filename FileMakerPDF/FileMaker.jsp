<%@ page language="java"
         import="com.zhuozhengsoft.pageoffice.DocumentOpenType,com.zhuozhengsoft.pageoffice.FileMakerCtrl,com.zhuozhengsoft.pageoffice.wordwriter.WordDocument"
         pageEncoding="utf-8" %>
<%
    FileMakerCtrl fmCtrl = new FileMakerCtrl(request);
    fmCtrl.setServerPage(request.getContextPath() + "/poserver.zz");
    WordDocument doc = new WordDocument();
    //禁用右击事件
    doc.setDisableWindowRightClick(true);
    //给数据区域赋值，即把数据填充到模板中相应的位置
    doc.openDataRegion("PO_company").setValue("北京卓正志远软件有限公司  ");
    //pdfName是传递给保存页面的pdf文件名称
    fmCtrl.setSaveFilePage("SaveMaker.jsp?pdfName=template.pdf");
    fmCtrl.setWriter(doc);
    fmCtrl.setJsFunction_OnProgressComplete("OnProgressComplete()");
    fmCtrl.fillDocumentAsPDF("doc/template.doc", DocumentOpenType.Word, "template.pdf");
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
        //转换完成后自动触发的事件
        function OnProgressComplete() {
            var type=<%=request.getParameter("type") %>;
            window.external.CallParentFunc("myFunc("+type+");"); //调用父页面的js函数
            window.external.close();//关闭POBrwoser窗口
        }
    </script>
    <!--**************   卓正 PageOffice 客户端代码结束    ************************-->
    <%=fmCtrl.getHtmlCode("FileMakerCtrl1")%>
</div>
</body>
</html>
