<%@ page language="java"
         import="com.zhuozhengsoft.pageoffice.OpenModeType,com.zhuozhengsoft.pageoffice.PageOfficeCtrl"
         pageEncoding="utf-8" %>
<%
    PageOfficeCtrl poCtrl = new PageOfficeCtrl(request);
    //设置服务器页面
    poCtrl.setServerPage(request.getContextPath() + "/poserver.zz");
    poCtrl.setCustomToolbar(false);
    poCtrl.setOfficeToolbars(false);
    poCtrl.setAllowCopy(false);//禁止拷贝
    poCtrl.setJsFunction_AfterDocumentOpened("AfterDocumentOpened");
    //打开Word文档
    poCtrl.webOpen("doc/test.doc", OpenModeType.docReadOnly, "张佚名");
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
    <title>最简单的打开保存Word文件</title>
</head>
<body>
<script type="text/javascript">
    function AfterDocumentOpened() {
        document.getElementById("PageOfficeCtrl1").SetEnableFileCommand(3, false); // 禁止保存
        document.getElementById("PageOfficeCtrl1").SetEnableFileCommand(4, false); // 禁止另存
        document.getElementById("PageOfficeCtrl1").SetEnableFileCommand(5, false); //禁止打印
        document.getElementById("PageOfficeCtrl1").SetEnableFileCommand(6, false); // 禁止页面设置
    }
</script>
<form id="form1">
    <p>点击“文件”菜单，可以看到“保存”、“另存为”、“页面设置”、“打印”菜单项已经变灰。保存菜单项不仅变灰，Ctrl+S也被禁用。</p>
    <div style=" width:auto; height:700px;">
        <%=poCtrl.getHtmlCode("PageOfficeCtrl1")%>
    </div>
</form>
</body>
</html>
