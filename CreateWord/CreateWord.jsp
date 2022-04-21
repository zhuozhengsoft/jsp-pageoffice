<%@ page language="java"
         import="com.zhuozhengsoft.pageoffice.DocumentVersion,com.zhuozhengsoft.pageoffice.PageOfficeCtrl"
         pageEncoding="utf-8" %>
<%
    //******************************卓正PageOffice组件的使用*******************************
    PageOfficeCtrl poCtrl1 = new PageOfficeCtrl(request);
    poCtrl1.setServerPage(request.getContextPath() + "/poserver.zz"); //此行必须
    //隐藏菜单栏
    poCtrl1.setMenubar(false);
    //隐藏工具栏
    poCtrl1.setCustomToolbar(false);
    poCtrl1.setJsFunction_BeforeDocumentSaved("BeforeDocumentSaved()");
    //设置保存页面
    poCtrl1.setSaveFilePage("SaveNewFile.jsp");
    //新建Word文件，webCreateNew方法中的两个参数分别指代“操作人”和“新建Word文档的版本号”
    poCtrl1.webCreateNew("张佚名", DocumentVersion.Word2003);
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
    <title></title>
    <link href="css/csstg.css" rel="stylesheet" type="text/css"/>
    <script type="text/javascript">
        function Save() {
            document.getElementById("PageOfficeCtrl1").WebSave();
            if (document.getElementById("PageOfficeCtrl1").CustomSaveResult == "ok") {
                alert('保存成功！');
                //返回列表页面
                if (window.external.CallParentFunc("freshIndex()") == 'poerror:parentlost') {
                    alert('父页面关闭或跳转刷新了，导致父页面函数没有调用成功！');
                }
                window.external.close();//关闭当前POBrower弹出的窗口
            } else {
                alert('保存失败！');
            }
        }

        function Cancel() {
            window.external.close();
        }

        function getFocus() {
            var str = document.getElementById("FileSubject").value;
            if (str == "请输入文档主题") {
                document.getElementById("FileSubject").value = "";
            }
        }

        function lostFocus() {
            var str = document.getElementById("FileSubject").value;
            if (str.length <= 0) {
                document.getElementById("FileSubject").value = "请输入文档主题";
            }
        }

        function BeforeDocumentSaved() {
            var str = document.getElementById("FileSubject").value;
            if (str.length <= 0) {
                document.getElementById("PageOfficeCtrl1").Alert("请输入文档主题");
                return false
            }
            else
                return true;
        }
    </script>
</head>
<body>
<form id="form2" action="">
    <div id="header">
        <div style="float: left; margin-left: 20px;">
            <img src="css/logo.jpg" height="30"/>
        </div>
        <ul>
            <li>
                <a href="#">卓正网站</a>
            </li>
            <li>
                <a href="#">客户问吧</a>
            </li>
            <li>
                <a href="#">在线帮助</a>
            </li>
            <li>
                <a href="#">联系我们</a>
            </li>
        </ul>
    </div>
    <div id="content">
        <div id="textcontent" style="width: 1000px; height: 800px;">
            <div class="flow4">
                <span style="width: 100px;"> &nbsp; </span>
            </div>
            <div>
                文档主题：
                <input name="FileSubject" id="FileSubject" type="text"
                       onfocus="getFocus()" onblur="lostFocus()" class="boder"
                       style="width: 180px;" value="请输入文档主题"/>
                <input type="button" onclick="Save()" value="保存并关闭"/>
                <input type="button" onclick="Cancel()" value="取消"/>
            </div>
            <div>
                &nbsp;
            </div>
            <%=poCtrl1.getHtmlCode("PageOfficeCtrl1")%>
        </div>
    </div>
    <div id="footer">
        <hr width="1000px;"/>
        <div>
            Copyright (c) 2019 北京卓正志远软件有限公司
        </div>
    </div>
</form>
</body>
</html>
