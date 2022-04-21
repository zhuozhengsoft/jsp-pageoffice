<%@ page language="java"
         import="com.zhuozhengsoft.pageoffice.OpenModeType,com.zhuozhengsoft.pageoffice.PageOfficeCtrl,java.net.URLDecoder,java.sql.Connection,java.sql.DriverManager"
         pageEncoding="utf-8" %>
<%@ page import="java.sql.ResultSet" %>
<%@ page import="java.sql.Statement" %>

<%
    int id = Integer.parseInt(request.getParameter("id"));
    //根据id查询数据库中对应的文档名称
    Class.forName("org.sqlite.JDBC");
    String strUrl = "jdbc:sqlite:"
            + this.getServletContext().getRealPath("demodata/SaveAndSearch.db");
    Connection conn = DriverManager.getConnection(strUrl);
    Statement stmt = conn.createStatement();
    String sql = "select * from word where id=" + id;
    ResultSet rs = stmt.executeQuery(sql);
    String FileName = "";
    while (rs.next()) {
        FileName = rs.getString("FileName");
    }
    stmt.close();
    conn.close();

    /******************************PageOffice编程开始**************************/
    PageOfficeCtrl poCtrl1 = new PageOfficeCtrl(request);
    poCtrl1.setServerPage(request.getContextPath() + "/poserver.zz"); //此行必须
    //隐藏菜单栏
    poCtrl1.setMenubar(false);
    poCtrl1.addCustomToolButton("保存", "Save()", 1);
    //设置保存页面
    poCtrl1.setSaveFilePage("SaveFile.jsp?id=" + id);
    String filePath = request.getSession().getServletContext().getRealPath("SaveAndSearch/doc/" + FileName + ".doc");
    //打开文件
    //这里打开的是中文名称的文件，所以必须用磁盘路径方式打开文件，为了支持linx服务器，文档路径前面必须加上“file://”前缀
    if(System.getProperty("os.name").toLowerCase().contains("linux")){    //判断该当前系统是否是linux操作系统
         poCtrl1.webOpen("file://"+filePath, OpenModeType.docNormalEdit, "张三");
    }else{
     poCtrl1.webOpen(filePath, OpenModeType.docNormalEdit, "张三");
    }

%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
    <title>编辑文档页面</title>
    <meta http-equiv="pragma" content="no-cache">
    <meta http-equiv="cache-control" content="no-cache">
    <meta http-equiv="expires" content="0">
    <meta http-equiv="keywords" content="keyword1,keyword2,keyword3">
    <meta http-equiv="description" content="This is my page">
    <script type="text/javascript">
        var strKey = '<%=URLDecoder.decode(request.getParameter("key"),"utf-8")%>';
        function Save() {
            document.getElementById("PageOfficeCtrl1").WebSave();
            //document.getElementById("PageOfficeCtrl1").CustomSaveResult获取的是保存页面的返回值
            if (document.getElementById("PageOfficeCtrl1").CustomSaveResult == "ok")
                document.getElementById("PageOfficeCtrl1").Alert("保存成功");
            else
                document.getElementById("PageOfficeCtrl1").Alert(document.getElementById("PageOfficeCtrl1").CustomSaveResult);
        }

        function SetKeyWord(key, visible) {
            if (key == "null" || "" == key) {
                document.getElementById("PageOfficeCtrl1").Alert("关键字为空。");
                return;
            }
            var sMac = "function myfunc()" + "\r\n"
                + "Application.Selection.HomeKey(6) \r\n"
                + "Application.Selection.Find.ClearFormatting \r\n"
                + "Application.Selection.Find.Replacement.ClearFormatting \r\n"
                + "Application.Selection.Find.Text = \"" + key + "\" \r\n"
                + "While (Application.Selection.Find.Execute()) \r\n"
                + "If (" + visible + ") Then \r\n"
                + "Application.Selection.Range.HighlightColorIndex = 7 \r\n"
                + "Else \r\n"
                + "Application.Selection.Range.HighlightColorIndex = 0 \r\n"
                + "End If \r\n"
                + "Wend \r\n"
                + "Application.Selection.HomeKey(6) \r\n"
                + "End function";
            document.getElementById("PageOfficeCtrl1").RunMacro("myfunc", sMac);
        }
    </script>
</head>
<body>
<form id="form1">
    <input name="button" id="Button1" type="button" onclick="SetKeyWord(strKey,true)" value="高亮显示关键字"/>
    <input name="button" id="Button2" type="button" onclick="SetKeyWord(strKey,false)" value="取消关键字显示"/>
    <div style="width: auto; height: 700px;">
        <%=poCtrl1.getHtmlCode("PageOfficeCtrl1")%>
    </div>
</form>
</body>
</html>
