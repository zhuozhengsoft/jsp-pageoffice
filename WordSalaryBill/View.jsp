<%@ page language="java"
         import="com.zhuozhengsoft.pageoffice.OpenModeType,com.zhuozhengsoft.pageoffice.PageOfficeCtrl,com.zhuozhengsoft.pageoffice.wordwriter.DataRegion,com.zhuozhengsoft.pageoffice.wordwriter.WordDocument,java.sql.Connection,java.sql.DriverManager,java.sql.ResultSet,java.sql.Statement"
         pageEncoding="utf-8" %>
<%@ page import="java.text.NumberFormat" %>
<%@ page import="java.text.SimpleDateFormat" %>
<%@ page import="java.util.Locale" %>
<%
    String err = "";
    String id = request.getParameter("ID").trim();
    PageOfficeCtrl poCtrl = new PageOfficeCtrl(request);
    //创建WordDocment对象
    WordDocument doc = new WordDocument();
    if (id != null && id.length() > 0) {
        String strSql = "select * from Salary where id =" + id
                + " order by ID";
        Class.forName("org.sqlite.JDBC");
        String strUrl = "jdbc:sqlite:"
                + this.getServletContext().getRealPath("demodata/WordSalaryBill.db");
        Connection conn = DriverManager.getConnection(strUrl);
        Statement stmt = conn.createStatement();
        ResultSet rs = stmt.executeQuery(strSql);

        //打开数据区域
        DataRegion datareg = doc.openDataRegion("PO_table");

        SimpleDateFormat formatDate = new SimpleDateFormat("yyyy-MM-dd");
        NumberFormat formater = NumberFormat
                .getCurrencyInstance(Locale.CHINA);
        if (rs.next()) {
            //给数据区域赋值
            doc.openDataRegion("PO_ID").setValue(id);
            doc.openDataRegion("PO_UserName").setValue(rs.getString("UserName"));
            doc.openDataRegion("PO_DeptName").setValue(rs.getString("DeptName"));

            String saltotal = rs.getString("SalTotal");
            if (saltotal != null && saltotal != "") {
                doc.openDataRegion("SalTotal").setValue(saltotal);
            } else {
                doc.openDataRegion("SalTotal").setValue("￥0.00");
            }

            String saldeduct = rs.getString("SalDeduct");
            if (saldeduct != null && saldeduct != "") {
                doc.openDataRegion("SalDeduct").setValue(saldeduct);
            } else {
                doc.openDataRegion("SalDeduct").setValue("￥0.00");
            }
            String salcount = rs.getString("SalCount");
            if (salcount != null && salcount != "") {
                doc.openDataRegion("SalCount").setValue(salcount);
            } else {
                doc.openDataRegion("SalCount").setValue("￥0.00");
            }
            String datatime = rs.getString("DataTime");
            if (datatime != null && datatime != "") {
                doc.openDataRegion("DataTime").setValue(datatime);
            } else {
                doc.openDataRegion("DataTime").setValue("");
            }

        } else {
            err = "<script>alert('未获得该员工的工资信息！');location.href='index.jsp'</script>";
        }
        rs.close();
        conn.close();
    } else {
        err = "<script>alert('未获得该工资信息的ID！');location.href='index.jsp'</script>";
    }

    poCtrl.setWriter(doc);
    poCtrl.setServerPage(request.getContextPath() + "/poserver.zz");
    poCtrl.webOpen("doc/template.doc", OpenModeType.docSubmitForm, "someBody");
%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
    <title>查看工资信息</title>
    <meta http-equiv="pragma" content="no-cache">
    <meta http-equiv="cache-control" content="no-cache">
    <meta http-equiv="expires" content="0">
    <meta http-equiv="keywords" content="keyword1,keyword2,keyword3">
    <meta http-equiv="description" content="This is my page">
</head>
<body>
<a href="#" onclick="window.external.close();">返回列表页</a>
<div style="width: auto; height: 700px;">
    <%=err%>
    <%=poCtrl.getHtmlCode("PageOfficeCtrl1")%>
</div>
</body>
</html>
