<%@ page language="java"
         import="com.zhuozhengsoft.pageoffice.OpenModeType,com.zhuozhengsoft.pageoffice.PageOfficeCtrl"
         pageEncoding="utf-8" %>
<%
    //******************************卓正PageOffice组件的使用*******************************
    PageOfficeCtrl poCtrl1 = new PageOfficeCtrl(request);
    poCtrl1.setServerPage(request.getContextPath() + "/poserver.zz"); //此行必须
    //隐藏菜单栏
    poCtrl1.setMenubar(false);
    //添加自定义按钮
    poCtrl1.addCustomToolButton("获取word选中的文字", "getSelectionText", 5);
    poCtrl1.webOpen("doc/test.doc", OpenModeType.docNormalEdit, "张三");
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
    <title>演示：js获取word选中的文字</title>
    <script type="text/javascript">
        function getSelectionText() {
            if (document.getElementById("PageOfficeCtrl1").Document.Application.Selection.Range.Text != "") {
                document.getElementById("PageOfficeCtrl1").Alert(document.getElementById("PageOfficeCtrl1").Document.Application.Selection.Range.Text);
            }
            else {
                document.getElementById("PageOfficeCtrl1").Alert("您没有选择任何文本。");
            }
        }
    </script>
</head>
<body>
<div style="font-size: 12px; line-height: 20px; border-bottom: dotted 1px #ccc; border-top: dotted 1px #ccc;
        padding: 5px;">
    <span style="color: red;">操作说明：</span>选中word文件中的一段文字，然后点“获取选中文字”按钮。<br/>
    关键代码：点右键，选择“查看源文件”，看js函数<span style="background-color: Yellow;">getSelectionText()</span></div>

<input id="Button1" type="button" onclick="getSelectionText();" value="js获取word选中的文字"/><br/>
<form id="form1">
    <div style=" width:auto; height:700px;">
        <!--**************   PageOffice 客户端代码开始    ************************-->
        <%=poCtrl1.getHtmlCode("PageOfficeCtrl1")%>
        <!--**************   PageOffice 客户端代码结束    ************************-->
    </div>
</form>
</body>
</html>

