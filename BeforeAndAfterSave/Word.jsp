<%@ page language="java"
         import="com.zhuozhengsoft.pageoffice.OpenModeType,com.zhuozhengsoft.pageoffice.PageOfficeCtrl"
         pageEncoding="utf-8" %>
<%
    PageOfficeCtrl poCtrl = new PageOfficeCtrl(request);
    poCtrl.setServerPage(request.getContextPath() + "/poserver.zz");
    // 设置文件保存之前执行的事件
    poCtrl.setJsFunction_BeforeDocumentSaved("BeforeDocumentSaved()");
    // 设置文件保存之后执行的事件
    poCtrl.setJsFunction_AfterDocumentSaved("AfterDocumentSaved()");
    poCtrl.setSaveFilePage("SaveFile.jsp");
    poCtrl.webOpen("doc/test.doc", OpenModeType.docNormalEdit, "张佚名");
%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
    <title>文档保存前和保存后执行的事件</title>
</head>
<body>
<script type="text/javascript">
    function BeforeDocumentSaved() {
        document.getElementById("PageOfficeCtrl1").Alert("BeforeDocumentSaved事件：文件就要开始保存了.");
        return true;
    }

    function AfterDocumentSaved(IsSaved) {
        if (IsSaved) {
            document.getElementById("PageOfficeCtrl1").Alert("AfterDocumentSaved事件：文件保存成功了.");
        }
    }
</script>
<form id="form1">
    演示: 文档保存前和保存后执行的事件。<br/><br/>
    <div style=" width:auto; height:700px;">
        <%=poCtrl.getHtmlCode("PageOfficeCtrl1")%>
    </div>
</form>
</body>
</html>
