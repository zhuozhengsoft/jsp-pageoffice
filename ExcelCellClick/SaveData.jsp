<%@ page language="java"
         import="com.zhuozhengsoft.pageoffice.excelreader.Sheet, com.zhuozhengsoft.pageoffice.excelreader.Table,com.zhuozhengsoft.pageoffice.excelreader.Workbook"
         pageEncoding="utf-8" %>
<%
    Workbook workBook = new Workbook(request, response);
    Sheet sheet = workBook.openSheet("Sheet1");
    Table table = sheet.openTable("B4:D8");
    String content = "";
    int result = 0;
    while (!table.getEOF()) {
        //获取提交的数值
        //DataFields.Count标识的是table的列数
        if (!table.getDataFields().getIsEmpty()) {
            content += "<br/>月份名称：" + table.getDataFields().get(0).getText();
            content += "<br/>计划完成量：" + table.getDataFields().get(1).getText();
            content += "<br/>实际完成量：" + table.getDataFields().get(2).getText();

            content += "<br/>*********************************************";
        }
        //循环进入下一行
        table.nextRow();
    }
    table.close();
    workBook.showPage(500, 400);
    workBook.close();
%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
    <title></title>
</head>
<body>
<form id="form1">
    <div style="border: solid 1px gray;">
        <div class="errTopArea"
             style="text-align: left; border-bottom: solid 1px gray;">
            [提示标题：这是一个开发人员可自定义的对话框]
        </div>
        <div class="errTxtArea" style="height: 88%; text-align: left">
            <b class="txt_title">
                <div style=" color:#FF0000;">
                    提交的信息如下：
                </div>
                <%=content%>
            </b>
        </div>
        <div class="errBtmArea" style="text-align: center;">
            <input type="button" class="btnFn" value=" 关闭 "
                   onclick="window.opener=null;window.open('','_self');window.close();"/>
        </div>
    </div>
</form>
</body>
</html>