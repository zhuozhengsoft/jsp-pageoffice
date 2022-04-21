<%@ page language="java"
         import="com.zhuozhengsoft.pageoffice.OpenModeType,com.zhuozhengsoft.pageoffice.PageOfficeCtrl,com.zhuozhengsoft.pageoffice.wordwriter.DataRegion,com.zhuozhengsoft.pageoffice.wordwriter.DataRegionInsertType,com.zhuozhengsoft.pageoffice.wordwriter.Table,com.zhuozhengsoft.pageoffice.wordwriter.WordDocument,java.sql.Connection,java.sql.DriverManager"
         pageEncoding="utf-8" %>
<%@ page import="java.sql.ResultSet" %>
<%@ page import="java.sql.Statement" %>
<%@ page import="java.text.NumberFormat" %>
<%@ page import="java.text.SimpleDateFormat" %>
<%@ page import="java.util.Locale" %>
<%
    if (request.getParameter("ids").equals(null)
            || request.getParameter("ids").equals("")) {
        return;
    }
    String idlist = request.getParameter("ids").trim();
    //从数据库中读取数据
    String strSql = "select * from Salary where ID in(" + idlist
            + ") order by ID";

    Class.forName("org.sqlite.JDBC");
    String strUrl = "jdbc:sqlite:"
            + this.getServletContext().getRealPath("demodata/WordSalaryBill.db");
    Connection conn = DriverManager.getConnection(strUrl);
    Statement stmt = conn.createStatement();
    ResultSet rs = stmt.executeQuery(strSql);

    SimpleDateFormat formatDate = new SimpleDateFormat("yyyy-MM-dd");
    NumberFormat formater = NumberFormat.getCurrencyInstance(Locale.CHINA);

    WordDocument doc = new WordDocument();
    DataRegion data = null;
    Table table = null;
    int i = 0;
    while (rs.next()) {
        data = doc.createDataRegion("reg" + i,
                DataRegionInsertType.Before, "[End]");
        data.setValue("[word]doc/template.doc[/word]");

        table = data.openTable(1);
        table.openCellRC(2, 1).setValue(rs.getString("ID"));

        //给单元格赋值
        table.openCellRC(2, 2).setValue(rs.getString("UserName"));
        table.openCellRC(2, 3).setValue(rs.getString("DeptName"));

        String saltotal = rs.getString("SalTotal");
        if (saltotal != null && saltotal != "") {
            table.openCellRC(2, 4).setValue(saltotal);
            //out.print(rs.getString("SalTotal"));
        } else {
            table.openCellRC(2, 4).setValue("￥0.00");
        }

        String saldeduct = rs.getString("SalDeduct");
        if (saldeduct != null && saldeduct != "") {
            table.openCellRC(2, 5).setValue(saldeduct);
        } else {
            table.openCellRC(2, 5).setValue("￥0.00");
        }
        String salcount = rs.getString("SalCount");
        if (salcount != null && salcount != "") {
            table.openCellRC(2, 6).setValue(salcount);
        } else {
            table.openCellRC(2, 6).setValue("￥0.00");
        }
        String datatime = rs.getString("DataTime");
        if (datatime != null && datatime != "") {
            table.openCellRC(2, 7).setValue(datatime);
        } else {
            table.openCellRC(2, 7).setValue("");
        }
        i++;
    }
    conn.close();
    // 设置PageOffice组件服务页面
    PageOfficeCtrl pCtrl = new PageOfficeCtrl(request);
    pCtrl.setWriter(doc);
    pCtrl.setServerPage(request.getContextPath() + "/poserver.zz");
    pCtrl.setCaption("生成工资条");
    pCtrl.setCustomToolbar(false);
    pCtrl.webOpen("doc/test.doc", OpenModeType.docNormalEdit, "somebody");
%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
    <title>生成工资条</title>
    <meta http-equiv="pragma" content="no-cache">
    <meta http-equiv="cache-control" content="no-cache">
    <meta http-equiv="expires" content="0">
    <meta http-equiv="keywords" content="keyword1,keyword2,keyword3">
    <meta http-equiv="description" content="This is my page">
</head>
<body>
<a href="#" onclick="window.external.close();">返回列表页</a>
<div style="width: 1000px; height: 800px;">
    <%=pCtrl.getHtmlCode("PageOfficeCtrl1")%>
</div>
</body>
</html>
