<%@ page language="java"
         import="com.zhuozhengsoft.pageoffice.OpenModeType,com.zhuozhengsoft.pageoffice.PageOfficeCtrl"
         pageEncoding="utf-8" %>
<%@ page import="com.zhuozhengsoft.pageoffice.wordwriter.WordDocument" %>
<%
    //******************************卓正PageOffice组件的使用*******************************
    //设置PageOffice服务器组件
	WordDocument doc=new WordDocument();
    doc.setDisableWindowSelection(true);//禁止选中
    doc.setDisableWindowRightClick(true);//禁止右键

    PageOfficeCtrl poCtrl1 = new PageOfficeCtrl(request);
    poCtrl1.setServerPage(request.getContextPath() + "/poserver.zz"); //此行必须

    poCtrl1.setMenubar(false);//隐藏菜单栏
    poCtrl1.setOfficeToolbars(false);//隐藏Office工具条
    poCtrl1.setCustomToolbar(false);//隐藏自定义工具栏
    poCtrl1.setJsFunction_AfterDocumentOpened("AfterDocumentOpened");
    //设置页面的显示标题
    poCtrl1.setCaption("演示：文件在线安全浏览");
    poCtrl1.setWriter(doc);//注意：千万别忘了这句，否则WordDocument设置的所有方法都不生效
    //打开文件
    poCtrl1.webOpen("doc/template.doc", OpenModeType.docReadOnly, "张三");
%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
    <title>演示：文件在线安全浏览</title>
    <meta http-equiv="pragma" content="no-cache">
    <meta http-equiv="cache-control" content="no-cache">
    <meta http-equiv="expires" content="0">
    <meta http-equiv="keywords" content="keyword1,keyword2,keyword3">
    <meta http-equiv="description" content="This is my page">
</head>
<body>
<script type="text/javascript">
    function AfterDocumentOpened() {
        document.getElementById("PageOfficeCtrl1").SetEnableFileCommand(4, false); //禁止另存
        document.getElementById("PageOfficeCtrl1").SetEnableFileCommand(5, false); //禁止打印
        document.getElementById("PageOfficeCtrl1").SetEnableFileCommand(6, false); //禁止页面设置
        document.getElementById("PageOfficeCtrl1").SetEnableFileCommand(8, false); //禁止打印预览
    }
</script>
<div style=" width:auto; height:700px;">
    <%=poCtrl1.getHtmlCode("PageOfficeCtrl1")%>
</div>
</body>
</html>
