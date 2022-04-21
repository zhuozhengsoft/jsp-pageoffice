<%@ page language="java" import="java.sql.Connection,java.sql.DriverManager,java.sql.ResultSet,java.sql.Statement"
         pageEncoding="utf-8" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head id="Head1">
    <title></title>
    <link href="css/style.css" rel="stylesheet" type="text/css"/>
    <!--pageoffice.js一定要引用  -->
    <script type="text/javascript" src="../pageoffice.js"></script>
    <script type="text/javascript">

        function onColor(td) {
            td.style.backgroundColor = '#D7FFEE';
        }

        function offColor(td) {
            td.style.backgroundColor = '';
        }

        function getFocus() {
            var str = document.getElementById("Input_KeyWord").value;
            if (str == "请输入关键字") {
                document.getElementById("Input_KeyWord").value = "";
            }

        }

        function lostFocus() {
            var str = document.getElementById("Input_KeyWord").value;
            if (str.length <= 0) {
                document.getElementById("Input_KeyWord").value = "请输入关键字";
            }
        }

        function copyKeyToInput(key) {
            document.getElementById("Input_KeyWord").value = key;
        }

        function mySubmit() {
            document.getElementById("form1").submit();

        }
    </script>
</head>
<body>
<!--header-->
<div class="zz-headBox br-5 clearfix" align="center">
    <div class="zz-head mc">
        <!--logo-->
        <div class="logo fl">
            <a href="#">
                <img src="images/logo.png" alt=""/></a></div>
        <!--logo end-->
        <ul class="head-rightUl fr">
            <li><a href="http://www.zhuozhengsoft.com">卓正网站</a></li>
            <li><a href="http://www.zhuozhengsoft.com/Technical/">技术支持</a></li>
            <li class="bor-0"><a href="http://www.zhuozhengsoft.com/about/about/">联系我们</a></li>
        </ul>
    </div>
</div>
<!--header end-->
<!--content-->
<div class="zz-content mc clearfix pd-28" align="center">
    <div class="demo mc">
        <h2 class="fs-16">
            PageOffice 实现Word文档的在线编辑保存和全文关键字搜索</h2>
    </div>
    <div class="demo mc">
        <h3 class="fs-12">
            搜索文件</h3>
        <form id="form1" action="FileManage.jsp" method="post">
            <table class="text" cellspacing="0" cellpadding="0" border="0">
                <tr>
                    <td style="font-size: 9pt" align="left">
                        通过文档内容中的关键字搜索文档&nbsp;&nbsp;&nbsp;
                    </td>
                    <td align="center">
                        <input name="Input_KeyWord" id="Input_KeyWord" type="text" onfocus="getFocus()"
                               onblur="lostFocus()"
                               class="boder" style="width: 180px;" value="请输入关键字"/>
                    </td>
                    <td style="width: 221px;">
                        &nbsp;
                        <input type="button" value="搜索" onclick="mySubmit();" style=" width:86px;"/>
                    </td>
                </tr>
                <tr>
                    <td>&nbsp;</td>
                    <td colspan="2">&nbsp;<span style="color:Gray;">热门搜索：</span>
                        <a href="#" style="color:#00217d;" onclick="copyKeyToInput('网站');">网站</a>
                        <a href="#" style="color:#00217d;" onclick="copyKeyToInput('软件');">软件</a>
                        <a href="#" style="color:#00217d;" onclick="copyKeyToInput('字体');">字体</a></td>
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
                <th width="90%">
                    文档名称
                </th>
                <th width="100%">
                    操作
                </th>
            </tr>
            </thead>
            <tbody>
            <%
                request.setCharacterEncoding("utf-8");
                String key = request.getParameter("Input_KeyWord");
                String sql = "";
                if (key != null && key.trim().length() > 0) {
                    request.setAttribute("key", key);
                    sql = "select * from word  where Content like '%" + key
                            + "%' order by ID desc";
                    System.out.println("--key--" + key);
                } else {
                    sql = "select * from word order by ID desc";
                }
                Class.forName("org.sqlite.JDBC");
                String strUrl = "jdbc:sqlite:"
                        + this.getServletContext().getRealPath("demodata/SaveAndSearch.db");
                Connection conn = DriverManager.getConnection(strUrl);
                Statement stmt = conn.createStatement();
                ResultSet rs = stmt.executeQuery(sql);
                int id;
                String FileName = "";
                String Content = "";
                while (rs.next()) {
                    FileName = rs.getString("FileName");
                    Content = rs.getString("Content");
                    id = rs.getInt("ID");
            %>
            <tr onmouseover='onColor(this)' onmouseout='offColor(this)'>
                <td>
                    <%=FileName%>
                </td>

                <td style='text-align:center;'>
                    <%
                        if (key != null && key.trim().length() > 0) { %>
                    <a style="color:#00217d;"
                       href="javascript:POBrowser.openWindowModeless('Edit.jsp?id=<%=id %>&key='+encodeURI('<%=key%>')+'','width=1200px;height=800px;')">编辑</a>

                    <%

                        }
                    %>
                    <%
                        if (key == null || key == "") { %>
                    <a style="color:#00217d;"
                       href="javascript:POBrowser.openWindowModeless('Edit.jsp?id=<%=id %>' ,'width=1200px;height=800px;');">编辑</a>

                    <%
                        }
                    %>

                </td>

                <%
                    }
                %>
            </tr>
            <%
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
