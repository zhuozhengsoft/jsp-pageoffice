<%@ page language="java"
         import="java.sql.Connection"
         pageEncoding="utf-8" %>
<%@ page import="java.sql.DriverManager" %>
<%@ page import="java.sql.ResultSet" %>
<%@ page import="java.sql.Statement" %>
<%@ page import="java.text.SimpleDateFormat" %>
<%@ page import="java.util.Date" %>
<%
    if (request.getParameter("op") != null && request.getParameter("op").length() > 0) {
        Class.forName("org.sqlite.JDBC");
        String strUrl = "jdbc:sqlite:"
                + this.getServletContext().getRealPath("demodata/CreateWord.db");
        Connection conn = DriverManager.getConnection(strUrl);
        Statement stmt = conn.createStatement();
        ResultSet rs = stmt.executeQuery("select Max(ID) from word");
        int newID = 1;
        if (rs.next()) {
            newID = Integer.parseInt(rs.getString(1)) + 1;
        }
        rs.close();

        String fileName = "aabb" + newID + ".doc";

        String FileSubject = "请输入文档主题";
        String getFile = (String) request.getParameter("FileSubject");
        if (getFile != null && getFile.length() > 0)
            FileSubject = new String(getFile.getBytes("iso-8859-1"), "utf-8");
        out.print(FileSubject);
        SimpleDateFormat df = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");//设置日期格式
        // new Date()为获取当前系统时间
        String strsql = "Insert into word(ID,FileName,Subject,SubmitTime) values(" + newID
                + ",'" + fileName + "','" + FileSubject + "','" + df.format(new Date()) + "')";
        stmt.executeUpdate(strsql);
        stmt.close();
        conn.close();

        response.sendRedirect("word-lists.jsp");
    }


%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
    <title></title>
    <link href="css/style.css" rel="stylesheet" type="text/css"/>
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8"/>
    <!--pageoffice.js文件一定要引用-->
    <script type="text/javascript" src="../pageoffice.js"></script>
    <script type="text/javascript">
        function onColor(td) {
            td.style.backgroundColor = '#D7FFEE';
        }

        function offColor(td) {
            td.style.backgroundColor = '';
        }


        function freshIndex() {
            window.location.href = "word-lists.jsp";
        }
    </script>

</head>
<body style=" text-align:center;">
<!--header-->
<div class="zz-headBox br-5 clearfix">
    <div class="zz-head mc">
        <!--logo-->
        <div class="logo fl">
            <a href="home.html">
                <img src="css/logo.png" alt=""/></a></div>
        <!--logo end-->
        <ul class="head-rightUl fr">
            <li><a href="http://www.zhuozhengsoft.com">卓正网站</a></li>
            <li class="bor-0"><a href="http://www.zhuozhengsoft.com">联系我们</a></li>
        </ul>
    </div>
</div>
<!--header end-->
<!--content-->
<div class="zz-content mc clearfix pd-28">
    <div class="demo mc">
        <h2 class="fs-16">
            PageOffice使用webCreateNew方式创建新文档</h2>
    </div>
    <div class="demo mc" style="text-align:left;">
        <h3 class="fs-12">
            新建文件</h3>
        <form id="form1" method="post" action="word-lists.jsp?op=create">
            <table class="text" cellspacing="0" cellpadding="0" border="0">
                <tr>
                    <td colspan="3">&nbsp;</td>
                </tr>
                <tr>
                    <td>
                        利用属性WebCreateNew创建新文件&nbsp;&nbsp;
                    </td>
                    <td>
                        &nbsp;<a
                            href="javascript:POBrowser.openWindowModeless('CreateWord.jsp' , 'width=1200px;height=800px;')"
                            style=" color:Blue; text-decoration:underline;">新建文件</a>
                    </td>
                    <td style="width: 221px;">
                        &nbsp;&nbsp;
                    </td>
                </tr>
            </table>
        </form>
    </div>
    <div class="zz-talbeBox mc">
        <h2 class="fs-12">
            文档列表</h2>
        <table class="zz-talbe">
            <thead>
            <tr>
                <th width="22%">
                    文档名称
                </th>
                <th width="10%">
                    创建日期
                </th>
            </tr>
            </thead>
            <tbody>
            <%
                Class.forName("org.sqlite.JDBC");
                String strUrl = "jdbc:sqlite:"
                        + this.getServletContext().getRealPath("demodata/CreateWord.db");
                Connection conn = DriverManager.getConnection(strUrl);
                Statement stmt = conn.createStatement();
                ResultSet rs = stmt.executeQuery("select * from word order by id desc");
                String fileName = "";
                String subject = "";
                String submitTime = "";
                while (rs.next()) {
                    int id = rs.getInt("ID");
                    fileName = rs.getString("FileName");
                    subject = rs.getString("Subject");
                    submitTime = rs.getString("SubmitTime");
            %>
            <tr onmouseover='onColor(this)' onmouseout='offColor(this)'>
                <td>
                    <a
                            href="javascript:POBrowser.openWindowModeless('OpenWord.jsp?id=<%=id%>' , 'width=1200px;height=800px;')"><%=subject %>
                    </a>
                </td>
                <%
                    if (submitTime != null && submitTime.length() > 0) {
                %>
                <td><%=new SimpleDateFormat("yyyy/MM/dd")
                        .format(new SimpleDateFormat("yyyy-MM-dd")
                                .parse(submitTime)) %>
                </td>
                <%
                } else {
                %>
                <td>&nbsp;</td>
                <%
                    }
                %>
            </tr>
            <%
                }
                stmt.close();
                conn.close();
            %>
            </tbody>
        </table>
    </div>
</div>
<!--content end-->
<!--footer-->
<div class="login-footer clearfix">
    Copyright &copy 2019 北京卓正志远软件有限公司
</div>
<!--footer end-->
</body>
</html>