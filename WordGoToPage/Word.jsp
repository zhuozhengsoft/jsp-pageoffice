<%@ page language="java"
         import="com.zhuozhengsoft.pageoffice.OpenModeType,com.zhuozhengsoft.pageoffice.PageOfficeCtrl"
         pageEncoding="utf-8" %>
<%
    PageOfficeCtrl poCtrl = new PageOfficeCtrl(request);
    //设置服务器页面
    poCtrl.setServerPage(request.getContextPath() + "/poserver.zz");
    //添加自定义按钮
    poCtrl.addCustomToolButton("保存", "Save", 1);
    //设置保存页面
    //poCtrl.setSaveFilePage("SaveFile.jsp");
    // 打开Word文档
    poCtrl.webOpen("doc/test.doc", OpenModeType.docNormalEdit, "张佚名");
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head id="Head1" runat="server">
    <title>演示：js实现word中页面跳转</title>
    <style>
        html, body {
            height: 100%;
        }

        .main {
            height: 100%;
        }
    </style>
</head>
<body>
<script type="text/javascript">
    function gotoPage(num) {
        var sMac = "function myfunc()" + "\r\n"
            + "  If " + num + " > Application.Selection.Information(4) Then " + "\r\n"
            + "     Msgbox \"超出文档范围，本文共\" & Application.Selection.Information(4) & \"页\"" + "\r\n"
            + "  End If " + "\r\n"
            + "  Selection.GoTo What:=wdGoToPage, Which:=wdGoToAbsolute, Name:= " + num + "\r\n"
            + "End function";
        return document.getElementById("PageOfficeCtrl1").RunMacro("myfunc", sMac);
    }

    function Button1_onclick() {
        var num = document.getElementById("pageNum").value;
        if ("" == Trim(num)) {
            document.getElementById("PageOfficeCtrl1").Alert("请输入页码");
            return;
        }
        if (isNaN(num)) {
            document.getElementById("PageOfficeCtrl1").Alert("只能输入数字");
            return;
        }
        gotoPage(num);
    }
</script>
<script type="text/javascript">
    function LTrim(str) {
        var i;
        for (i = 0; i < str.length; i++) {
            if (str.charAt(i) != " " && str.charAt(i) != " ") break;
        }
        str = str.substring(i, str.length);
        return str;
    }

    function RTrim(str) {
        var i;
        for (i = str.length - 1; i >= 0; i--) {
            if (str.charAt(i) != " " && str.charAt(i) != " ") break;
        }
        str = str.substring(0, i + 1);
        return str;
    }

    function Trim(str) {
        return LTrim(RTrim(str));
    }
</script>
<div style="font-size:12px; line-height:20px; border-bottom:dotted 1px #ccc;border-top:dotted 1px #ccc; padding:5px;">
    <span style="color:red;">操作说明：</span>请输入页码后点跳转按钮。页码：
    <input id="pageNum" type="text" value="3"/>
    <input id="Button1" type="button" value="跳转" onclick="return Button1_onclick()"/>
    <br/>

    关键代码：点右键，选择“查看源文件”，看js函数<span style="background-color:Yellow;">gotoPage(num)</span></div>
<br/>
<form id="form1" style="height:100%;">
    <div class="main">
        <!--**************   PageOffice 客户端代码开始    ************************-->
        <div style=" width:auto; height:700px;">
            <%=poCtrl.getHtmlCode("PageOfficeCtrl1")%>
        </div>
        <!--**************   PageOffice 客户端代码结束    ************************-->
    </div>
</form>
</body>
</html>

