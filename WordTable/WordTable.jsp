<%@ page language="java" import="com.zhuozhengsoft.pageoffice.OpenModeType" pageEncoding="utf-8" %>
<%@ page import="com.zhuozhengsoft.pageoffice.PageOfficeCtrl,com.zhuozhengsoft.pageoffice.wordwriter.Table" %>
<%@ page import="com.zhuozhengsoft.pageoffice.wordwriter.WordDocument" %>
<%
    PageOfficeCtrl poCtrl = new PageOfficeCtrl(request);
    WordDocument doc = new WordDocument();
    Table table1 = doc.openDataRegion("PO_T001").openTable(1);
    table1.openCellRC(1, 1).setValue("PageOffice组件");
    int dataRowCount = 5;//需要插入数据的行数
    int oldRowCount = 3;//表格中原有的行数
    // 扩充表格
    for (int j = 0; j < dataRowCount - oldRowCount; j++) {
        table1.insertRowAfter(table1.openCellRC(2, 5));  //在第2行的最后一个单元格下插入新行
    }
    // 填充数据
    int i = 1;
    while (i <= dataRowCount) {
        table1.openCellRC(i, 2).setValue("AA" + String.valueOf(i));
        table1.openCellRC(i, 3).setValue("BB" + String.valueOf(i));
        table1.openCellRC(i, 4).setValue("CC" + String.valueOf(i));
        table1.openCellRC(i, 5).setValue("DD" + String.valueOf(i));
        i++;
    }
    poCtrl.setWriter(doc);
    poCtrl.setServerPage(request.getContextPath() + "/poserver.zz");
    poCtrl.webOpen("doc/test_table.doc", OpenModeType.docNormalEdit, "张佚名");
%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
    <title>Word中的Table的数据填充</title>
    <meta http-equiv="pragma" content="no-cache">
    <meta http-equiv="cache-control" content="no-cache">
    <meta http-equiv="expires" content="0">
    <meta http-equiv="keywords" content="keyword1,keyword2,keyword3">
    <meta http-equiv="description" content="This is my page">
</head>
<body>
<div style="width: auto; height: 600px;">
    <%=poCtrl.getHtmlCode("PageOfficeCtrl1")%>
</div>
</body>
</html>
