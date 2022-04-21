<%@ page language="java"
         import="com.zhuozhengsoft.pageoffice.OpenModeType,com.zhuozhengsoft.pageoffice.PageOfficeCtrl"
         pageEncoding="utf-8" %>
<%
    PageOfficeCtrl poCtrl = new PageOfficeCtrl(request);
    //设置服务器页面
    poCtrl.setServerPage(request.getContextPath() + "/poserver.zz");
    //添加自定义按钮
    poCtrl.addCustomToolButton("另存HTML", "saveAsHTML", 1);
    //设置保存页面
    poCtrl.setSaveFilePage("SaveFile.jsp");
    //打开Word文档
    poCtrl.webOpen("doc/test.doc", OpenModeType.docNormalEdit, "张佚名");
%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
    <title>Word文件另存为HTML</title>
</head>
<body>
<script type="text/javascript">
    function saveAsHTML() {
        document.getElementById("PageOfficeCtrl1").WebSaveAsHTML();
        document.getElementById("PageOfficeCtrl1").Alert("HTML格式的文件已经保存到 SaveAsHTML\\doc 目录下。");
        document.getElementById("div1").innerHTML = "<a href='doc/test.htm?r=" + Math.random() + "'> 查看另存的html文件<a><br><br>";
    }
</script>
<form id="form1">
    <div id="div1"></div>
    <div style=" width:1000px; height:800px;">
        <%=poCtrl.getHtmlCode("PageOfficeCtrl1")%>
    </div>
</form>
</body>
</html>
