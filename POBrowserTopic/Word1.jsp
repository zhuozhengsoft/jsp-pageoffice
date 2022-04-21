<%@ page language="java"
         import="com.zhuozhengsoft.pageoffice.OpenModeType,com.zhuozhengsoft.pageoffice.PageOfficeCtrl"
         pageEncoding="utf-8" %>

<%
    PageOfficeCtrl poCtrl = new PageOfficeCtrl(request);
    //设置服务器页面
    poCtrl.setServerPage(request.getContextPath() + "/poserver.zz");
    //添加自定义按钮
    poCtrl.addCustomToolButton("保存", "Save", 1);

    poCtrl.setJsFunction_AfterDocumentOpened("AfterDocumentOpened()");
    //设置保存页面
    poCtrl.setSaveFilePage("SaveFile.jsp");
    //打开Word文档
    poCtrl.webOpen("doc/test.doc", OpenModeType.docNormalEdit, "张佚名");

%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
    <title>PageOfficeLink方式最简单的打开编辑保存文档</title>
</head>
<body>
<script type="text/javascript">
    function Save() {
        document.getElementById("PageOfficeCtrl1").WebSave();
    }
</script>
<script type="text/javascript">
    function AfterDocumentOpened() {
        document.getElementById("Text1").value = document.getElementById("PageOfficeCtrl1").DataRegionList.GetDataRegionByName("PO_Title").Value;
    }

    function setTitleText() {
        document.getElementById("PageOfficeCtrl1").DataRegionList.GetDataRegionByName("PO_Title").Value = document.getElementById("Text1").value;
    }
</script>
<p style=" text-indent:10px;">
    PageOffice 4.0增加了全新的文件打开方式“POBrowser 方式”，此方式提供了更完美的浏览器兼容性解决方案。
</p>
<p style=" text-indent:10px;">
    常规打开文档超链接的代码写法：&lt;a href=&quot;Word.jsp?id=12&quot;&gt;某某公司公文-12&lt;/a&gt;</p>
<p style=" text-indent:10px;">
    POBrowser打开文档超链接的代码写法：
    &lt;a href=&quot;<span style=" background-color:#D2E9FF;">javascript:POBrowser.openWindowModeless( &quot;</span>Word.jsp?id=12<span
        style=" background-color:#D2E9FF;">&quot;,&quot;width=800px;height=800px;&quot;)&gt;</span>
    某某公司公文-12&lt;/a&gt;
    &nbsp;</p>
文档标题：<input id="Text1" type="text" size="50"/>
<input type="button" value="修改" onclick="setTitleText();"/>
<form id="form1">
    <div style=" width:auto; height:700px;">
        <%=poCtrl.getHtmlCode("PageOfficeCtrl1")%>
    </div>
</form>
</body>
</html>
