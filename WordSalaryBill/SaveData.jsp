<%@ page language="java"
         import="com.zhuozhengsoft.pageoffice.wordreader.WordDocument,java.sql.Connection,java.sql.DriverManager,java.sql.Statement"
         pageEncoding="utf-8" %>
<%
    String err = "";
    String id = request.getParameter("ID");
    if (id != null && id.length() > 0) {
        String strSql = "select * from Salary where id =" + id
                + " order by ID";
        Class.forName("org.sqlite.JDBC");
        String strUrl = "jdbc:sqlite:"
                + this.getServletContext().getRealPath("demodata/WordSalaryBill.db");
        Connection conn = DriverManager.getConnection(strUrl);
        Statement stmt = conn.createStatement();

        String userName = "", deptName = "", salTotoal = "0", salDeduct = "0", salCount = "0", dateTime = "";
        //-----------  PageOffice 服务器端编程开始  -------------------//
        WordDocument doc = new WordDocument(request, response);
        userName = doc.openDataRegion("PO_UserName").getValue();
        deptName = doc.openDataRegion("PO_DeptName").getValue();
        //将格式化的数据转化为String存到数据库
        salTotoal = doc.openDataRegion("PO_SalTotal").getValue();
        salDeduct = doc.openDataRegion("PO_SalDeduct").getValue();
        salCount = doc.openDataRegion("PO_SalCount").getValue();
        dateTime = doc.openDataRegion("PO_DataTime").getValue();

        String sql = "UPDATE Salary SET UserName='" + userName
                + "',DeptName='" + deptName + "',SalTotal='" + salTotoal
                + "',SalDeduct='" + salDeduct + "',SalCount='" + salCount
                + "',DataTime='" + dateTime + "' WHERE ID=" + id;

        int count = stmt.executeUpdate(sql);
        if (count > 0) {
            //向客户端控件返回以上代码执行成功的消息。
            doc.setCustomSaveResult("ok");
        }
        doc.close();
        conn.close();
    } else {
        err = "<script>alert('未获得文件的ID，保存失败！');location.href='Default.aspx'</script>";
    }
%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
    <title>My JSP 'SaveData.jsp' starting page</title>
    <meta http-equiv="pragma" content="no-cache">
    <meta http-equiv="cache-control" content="no-cache">
    <meta http-equiv="expires" content="0">
    <meta http-equiv="keywords" content="keyword1,keyword2,keyword3">
    <meta http-equiv="description" content="This is my page">
    <!--
<link rel="stylesheet" type="text/css" href="styles.css">
-->
</head>
<body>
<%=err%>
</body>
</html>
