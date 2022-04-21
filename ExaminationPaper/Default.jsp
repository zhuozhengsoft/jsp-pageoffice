<%@ page language="java" import="java.sql.Connection,java.sql.DriverManager"
         pageEncoding="utf-8" %>
<%@ page import="java.sql.ResultSet" %>
<%@ page import="java.sql.Statement" %>
<%
    Class.forName("org.sqlite.JDBC");
    String strUrl = "jdbc:sqlite:"
            + this.getServletContext().getRealPath("demodata/ExaminationPaper.db");
    Connection conn = DriverManager.getConnection(strUrl);
    Statement stmt = conn.createStatement();
    ResultSet rs = stmt.executeQuery("Select * from stream");
    boolean flg = false;//标识是否有数据
    StringBuilder strHtmls = new StringBuilder();
    strHtmls.append("<tr  style='background-color:#FEE;'>");
    strHtmls.append("<td style='text-align:center;width=10%' >选择</td>");
    strHtmls.append("<td style='text-align:center;width=30%'>题库编号</td>");
    strHtmls.append("<td style='text-align:center;width=60%'>操作</td>");
    strHtmls.append("</tr>");
    while (rs.next()) {
        flg = true;
        String pID = rs.getString("ID");
        strHtmls.append("<tr  style='background-color:white;'>");
        strHtmls.append("<td style='text-align:center'><input id='check" + pID + "'  type='checkbox' /></td>");
        strHtmls.append("<td style='text-align:center'>选择题-" + pID + "</td>");
        strHtmls.append("<td style='text-align:center'><a href='javascript:POBrowser.openWindowModeless(\"Edit.jsp?id=" + pID + "\" ,\"width=1200px;height=800px;\");'>编辑</a></td>");
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
    <title>无标题页</title>
    <!--pageoffice.js文件一定要引用-->
    <script type="text/javascript" src="../pageoffice.js"></script>
</head>
<body style="text-align: center">
<script type="text/javascript">
    // JS方式生成试卷
    function button1Click() {
        var ids = "";
        for (var i = 1; i < 11; i++) {
            if (document.getElementById("check" + i.toString()).checked) {
                ids += i.toString() + ",";
            }
        }
        if (ids == "")
            alert("请先选择要生成的试题");
        else
            location.href = "javascript:POBrowser.openWindowModeless('Compose.jsp?ids=" + ids.substr(0, ids.length - 1) + "', 'width=1200px;height=800px;');";
    }

    // 后台编程生成试卷
    function button2Click() {
        var ids = "";
        for (var i = 1; i < 11; i++) {
            if (document.getElementById("check" + i.toString()).checked) {
                ids += i.toString() + ",";
            }
        }
        if (ids == "")
            alert("请先选择要生成的试题");
        else
            location.href = "javascript:POBrowser.openWindowModeless('Compose2.jsp?ids=" + ids.substr(0, ids.length - 1) + "', 'width=1200px;height=800px;');";
    }
</script>
<div>
    <table align="center" style="color:red;">
        <tr>
            <td>1.可以点"操作"栏中的"编辑"来编辑各个试题</td>
        </tr>
        <tr>
            <td>2.可以选择多个试题，点"生成试卷"按钮，生成试卷</td>
        </tr>
    </table>
</div>
<div style="text-align: center;">
    <form id="form1" name="form1" method="post" action="Compose.jsp">
        <table style="background-color: #999; width: 600px; margin-left:28%; align:center;">
            <%=strHtmls%>
        </table>
        <br/>
        <input type="button" onclick="button1Click();" id="Button1" value="JS方式生成试卷"/><span>（所有版本）</span>
        <input type="button" onclick="button2Click();" id="Button2" value="后台编程生成试卷"/><span
            style="color:Red;">（专业版、企业版）</span>
    </form>
</div>
</body>
</html>
