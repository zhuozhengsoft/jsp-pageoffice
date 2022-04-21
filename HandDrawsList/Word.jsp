<%@ page language="java"
         import="com.zhuozhengsoft.pageoffice.OpenModeType,com.zhuozhengsoft.pageoffice.PageOfficeCtrl"
         pageEncoding="utf-8" %>
<%
    //******************************卓正PageOffice组件的使用*******************************
    //设置PageOffice服务器组件
    PageOfficeCtrl poCtrl1 = new PageOfficeCtrl(request);
    poCtrl1.setServerPage(request.getContextPath() + "/poserver.zz"); //此行必须
    poCtrl1.setJsFunction_AfterDocumentOpened("AfterDocumentOpened()");
    poCtrl1.addCustomToolButton("保存", "Save()", 1);
    poCtrl1.addCustomToolButton("开始手写", "StartHandDraw()", 3);
    poCtrl1.addCustomToolButton("设置线宽", "SetPenWidth()", 5);
    poCtrl1.addCustomToolButton("设置颜色", "SetPenColor()", 5);
    poCtrl1.addCustomToolButton("设置笔型", "SetPenType()", 5);
    poCtrl1.addCustomToolButton("设置缩放", "SetPenZoom()", 5);
    poCtrl1.setOfficeToolbars(false);//隐藏office工具栏
    poCtrl1.setSaveFilePage("SaveFile.jsp");
    //打开文件
    poCtrl1.webOpen("doc/test.doc", OpenModeType.docHandwritingOnly, "John");
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
    <title></title>
</head>
<body>
<script type="text/javascript">
    function Save() {
        document.getElementById("PageOfficeCtrl1").WebSave();
    }

    function StartHandDraw() {
        document.getElementById("PageOfficeCtrl1").HandDraw.Start();
    }

    //设置线宽
    function SetPenWidth() {
        document.getElementById("PageOfficeCtrl1").HandDraw.SetPenWidth(5);
    }

    //设置颜色
    function SetPenColor() {

        document.getElementById("PageOfficeCtrl1").HandDraw.SetPenColor(5292104);
    }

    //设置笔型
    function SetPenType() {

        document.getElementById("PageOfficeCtrl1").HandDraw.SetPenType(1);
    }

    //设置缩放
    function SetPenZoom() {

        document.getElementById("PageOfficeCtrl1").HandDraw.SetPenZoom(50);
    }

    //撤销最近一次手写
    function UndoHandDraw() {

        document.getElementById("PageOfficeCtrl1").HandDraw.Undo();
    }

    //退出手写
    function ExitHandDraw() {

        document.getElementById("PageOfficeCtrl1").HandDraw.Exit();
    }


    function AfterDocumentOpened() {
        refreshList();
    }


    function refreshList() {
        document.getElementById("PageOfficeCtrl1").HandDraw.Refresh();
        var i;
        document.getElementById("ul_Comments").innerHTML = "";
        if (document.getElementById("PageOfficeCtrl1").HandDraw.Count != 0) {
            for (i = 0; i < document.getElementById("PageOfficeCtrl1").HandDraw.Count; i++) {
                handDraw = document.getElementById("PageOfficeCtrl1").HandDraw.Item(i);
                var str = "";
                str = str + "第" + handDraw.PageNumber + "页" + "," + handDraw.UserName + ", " + handDraw.DateTime;
                document.getElementById("ul_Comments").innerHTML += "<li><a href='#' onclick='goToHandDraw(" + i + ")'>" + str + "</a></li>";
            }
        } else {
            document.getElementById("PageOfficeCtrl1").Alert("当前文档没有手写批注!");
        }

    }

    //定位到当前手写批注
    function goToHandDraw(index) {
        document.getElementById("PageOfficeCtrl1").HandDraw.Item(index).Locate();
    }

    function refresh_click() {
        //刷新手写批注集合
        refreshList();
    }
</script>
<form id="form1">
    <div style=" width:1380px; height:700px;">
        <div id="Div_Comments" style=" float:left; width:300px; height:700px; border:solid 1px red;">
            <h3>手写批注列表</h3>
            <input type="button" name="refresh" value="刷新" onclick="return refresh_click()"/>
            <ul id="ul_Comments">

            </ul>
        </div>
        <div style=" width:1000px; height:700px; float:right;">
            <%=poCtrl1.getHtmlCode("PageOfficeCtrl1")%>
        </div>
    </div>
</form>
</body>
</html>

