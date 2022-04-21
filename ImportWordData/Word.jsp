<%@ page language="java" import="com.zhuozhengsoft.pageoffice.PageOfficeCtrl,com.zhuozhengsoft.pageoffice.wordwriter.WordDocument" pageEncoding="utf-8" %>
<%
    PageOfficeCtrl poCtrl = new PageOfficeCtrl(request);
    poCtrl.addCustomToolButton("导入文件", "importData()", 15);
    poCtrl.addCustomToolButton("提交数据", "submitData()", 1);
    WordDocument doc = new WordDocument();
    poCtrl.setWriter(doc);
    poCtrl.setServerPage(request.getContextPath() + "/poserver.zz");
    poCtrl.setSaveDataPage("SaveData.jsp");
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
    <title>导入文件并提交数据</title>
    <meta http-equiv="pragma" content="no-cache">
    <meta http-equiv="cache-control" content="no-cache">
    <meta http-equiv="expires" content="0">
    <meta http-equiv="keywords" content="keyword1,keyword2,keyword3">
    <meta http-equiv="description" content="This is my page">
</head>
<body>
<!-- ***************************************PageOffice组件的使用****************************************** -->
<script type="text/javascript">
    function importData() {
        document.getElementById("PageOfficeCtrl1").WordImportDialog();
    }

    function submitData() {
        document.getElementById("PageOfficeCtrl1").WebSave();

    }
</script>
<div style="color:Red">请导入“/Samples4/ImportWordData”下的ImportWord.doc文档查看演示效果。</div>
<div style="width: auto; height: 600px;">
    <%=poCtrl.getHtmlCode("PageOfficeCtrl1")%>
</div>
<!-- ***************************************PageOffice组件的使用****************************************** -->
</body>
</html>
