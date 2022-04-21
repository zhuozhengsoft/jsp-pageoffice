<%@ page language="java"
         import="java.sql.Connection, java.sql.DriverManager,java.sql.ResultSet,java.sql.Statement"
         pageEncoding="utf-8" %>
<%
    Class.forName("org.sqlite.JDBC");
    String strUrl = "jdbc:sqlite:"
            + this.getServletContext().getRealPath("demodata/WordSalaryBill.db");
    Connection conn = DriverManager.getConnection(strUrl);
    Statement stmt = conn.createStatement();
    ResultSet rs = stmt.executeQuery("select * from Salary  order by ID");
    boolean flg = false;//标识是否有数据
    StringBuilder strHtmls = new StringBuilder();
    //SimpleDateFormat  formatDate = new SimpleDateFormat("yyyy-MM-dd");
    //DateFormat format = new SimpleDateFormat("yyyy-MM-dd");
    //NumberFormat formater = NumberFormat.getCurrencyInstance(Locale.CHINA);
    strHtmls.append("<tr  style='height:40px;padding:0; border-right:1px solid #a2c5d9; border-bottom:1px solid #a2c5d9; background:#edf8fe; font-weight:bold; color:#666;text-align:center; text-indent:0px;'>");
    strHtmls.append("<td width='5%' >选择</td>");
    strHtmls.append("<td width='10%' >员工编号</td>");
    strHtmls.append("<td width='10%' >员工姓名</td>");
    strHtmls.append("<td width='15%' >所在部门</td>");
    strHtmls.append("<td width='10%' >应发工资</td>");
    strHtmls.append("<td width='10%' >扣除金额</td>");
    strHtmls.append("<td width='10%' >实发工资</td>");
    strHtmls.append("<td width='10%' >发放日期</td>");
    strHtmls.append("<td width='20%' >操作</td>");
    strHtmls.append("</tr>");
    while (rs.next()) {
        flg = true;
        String pID = rs.getString("ID");
        strHtmls.append("<tr  style='height:40px; text-indent:10px; padding:0; border-right:1px solid #a2c5d9; border-bottom:1px solid #a2c5d9; color:#666;'>");
        strHtmls.append("<td style=' text-align:center;'><input id='check" + pID + "'  type='checkbox' /></td>");
        strHtmls.append("<td style=' text-align:left;'>" + pID + "</td>");
        strHtmls.append("<td style=' text-align:left;'>" + rs.getString("UserName").toString() + "</td>");
        strHtmls.append("<td style=' text-align:left;'>" + rs.getString("DeptName").toString() + "</td>");

        strHtmls.append("<td style=' text-align:left;'>" + rs.getString("SalTotal").toString() + "</td>");
        strHtmls.append("<td style=' text-align:left;'>" + rs.getString("SalDeduct").toString() + "</td>");
        strHtmls.append("<td style=' text-align:left;'>" + rs.getString("SalCount").toString() + "</td>");
        strHtmls.append("<td style=' text-align:center;'>" + rs.getString("DataTime") + "</td>");
        strHtmls.append("<td style=' text-align:center;'><a href='javascript:POBrowser.openWindowModeless(\"View.jsp?ID=" + pID + "\" ,\"width=1200px;height=800px;\");'>查看</a>&nbsp;&nbsp;&nbsp;&nbsp;<a href='javascript:POBrowser.openWindowModeless(\"Openfile.jsp?ID=" + pID + "\" ,\"width=1200px;height=800px;\");'>编辑</a></td>");
        strHtmls.append("</tr>");
    }

    if (!flg) {
        strHtmls.append("<tr>\r\n");
        strHtmls.append("<td width='100%' height='100' align='center'>对不起，暂时没有可以操作的数据。\r\n");
        strHtmls.append("</td></tr>\r\n");
    }
%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
    <title>动态生成工资条</title>
    <!--pageoffice.js文件一定要引用  -->
    <script type="text/javascript" src="../pageoffice.js"></script>
    <script type="text/javascript">
        function getID() {
            var ids = "";
            for (var i = 1; i < 11; i++) {
                if (document.getElementById("check" + i.toString()).checked) {
                    ids += i.toString() + ",";
                }
            }

            if (ids == "")
                alert("请先选择要生成工资条的记录");
            else
                location.href = "javascript:POBrowser.openWindowModeless('Compose.jsp?ids=" + ids.substr(0, ids.length - 1) + " ', 'width=1200px;height=800px;');";
        }

    </script>
</head>
<body>
<div style="text-align: center;">
    <p style="color: Red; font-size: 14px;">
        1.可以点击“操作”栏中的“编辑”来编辑各个员工的工作条&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<br/>
        2.可以点击“操作”栏中的“查看”来查看各个员工的工作条&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<br/>
        3.可选择多条工资记录，点“生成工资条”按钮，生成工资条&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</p>
    <table
            style="width: 60%; font-size: 12px; text-align: center; border: solid 1px #a2c5d9;">
        <%=strHtmls %>
    </table>
    <br/>
    <input type="button" value="生成工资条" onclick="getID()"/>
</div>
</body>
</html>
