<%@ page language="java"
         import="com.zhuozhengsoft.pageoffice.OpenModeType,com.zhuozhengsoft.pageoffice.PageOfficeCtrl"
         pageEncoding="utf-8" %>
<%
    //******************************卓正PageOffice组件的使用*******************************
    //设置PageOffice服务器组件
    PageOfficeCtrl poCtrl1 = new PageOfficeCtrl(request);
    poCtrl1.setServerPage(request.getContextPath() + "/poserver.zz"); //此行必须
    //隐藏菜单栏
    poCtrl1.setMenubar(false);
    //隐藏自定义工具栏
    poCtrl1.setCustomToolbar(false);
    //打开文件
    poCtrl1.webOpen("doc/test.doc", OpenModeType.docNormalEdit, "张三");
%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
    <title>执行宏命令</title>
    <script type="text/javascript">
        function Button1_onclick() {
            document.getElementById("PageOfficeCtrl1").RunMacro("myfunc", document.getElementById("textarea1").value);
        }
    </script>
</head>
<body>
<form id="form1">
    <div style="font-size: 12px; line-height: 20px; border-bottom: dotted 1px #ccc; border-top: dotted 1px #ccc;
        padding: 5px;">
        注意：<span style="background-color: Yellow;">执行“执行宏myfunc”按钮之前需先设置好MS Word的关于执行宏命令的设置。
        <br/>设置步骤如下：打开一个Word文档，点击“文件”→“选项”→“信任中心”→“信任中心设置”→“宏设置”→勾选上“信任对VBA工程对象模型的访问（V）”</span>
    </div>
    <textarea id="textarea1" name="textarea1" style=" height:87px; width:486px;" rows="" cols="">
    sub myfunc() 
msgbox "123" 
end sub
    </textarea>
    <input id="Button1" type="button" value="执行宏myfunc" onclick="return Button1_onclick()"/>
    <div style=" height:800px;">
        <%=poCtrl1.getHtmlCode("PageOfficeCtrl1")%>
    </div>
</form>
</body>
</html>

