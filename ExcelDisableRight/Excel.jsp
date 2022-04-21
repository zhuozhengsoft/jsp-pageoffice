<%@ page language="java"
         import="com.zhuozhengsoft.pageoffice.OpenModeType,com.zhuozhengsoft.pageoffice.PageOfficeCtrl,com.zhuozhengsoft.pageoffice.excelwriter.Workbook"
         pageEncoding="utf-8" %>
<%
    PageOfficeCtrl poCtrl = new PageOfficeCtrl(request);
    //设置服务器页面
    poCtrl.setServerPage(request.getContextPath() + "/poserver.zz");
    Workbook workBoook = new Workbook();
    workBoook.setDisableSheetRightClick(true);//禁止当前工作表鼠标右键
    //workBoook.setDisableSheetDoubleClick(true);//禁止当前工作表鼠标双击
    //workBoook.setDisableSheetSelection(true);//禁止在当前工作表中选择内容
    poCtrl.setWriter(workBoook);
    //打开Word文档
    poCtrl.webOpen("doc/test.xls", OpenModeType.xlsNormalEdit, "张佚名");
%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
    <title>禁止Excel文档鼠标右键</title>
</head>
<body>
<form id="form1">
    <div style="color:Red">打开Excel文档后，鼠标右键，发现右键失效。</div>
    <div style=" width:100%; height:700px;">
        <%=poCtrl.getHtmlCode("PageOfficeCtrl1")%>
    </div>
</form>
</body>
</html>
