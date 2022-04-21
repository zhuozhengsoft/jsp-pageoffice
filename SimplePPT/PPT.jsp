<%@ page language="java"
         import="com.zhuozhengsoft.pageoffice.OpenModeType,com.zhuozhengsoft.pageoffice.PageOfficeCtrl"
         pageEncoding="utf-8" %>
<%
    PageOfficeCtrl poCtrl = new PageOfficeCtrl(request);
    //设置服务器页面
    poCtrl.setServerPage(request.getContextPath() + "/poserver.zz");
    //添加自定义按钮
    poCtrl.addCustomToolButton("保存", "Save", 1);
    poCtrl.addCustomToolButton("关闭", "Close", 21);
    //设置保存页面
    poCtrl.setSaveFilePage("SaveFile.jsp");
    //打开Word文档
    poCtrl.webOpen("doc/test.ppt", OpenModeType.pptNormalEdit, "张佚名");
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<body>
<script type="text/javascript">
    function Save() {
        document.getElementById("PageOfficeCtrl1").WebSave();
    }

    function Close() {
        window.external.close();
    }
</script>
<div style="height:750px;width:auto;">
    <%=poCtrl.getHtmlCode("PageOfficeCtrl1")%>
</div>
</body>
</html>
