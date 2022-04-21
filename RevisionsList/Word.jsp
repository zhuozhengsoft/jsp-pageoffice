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
    poCtrl1.setOfficeToolbars(false);//隐藏office工具栏
    poCtrl1.setSaveFilePage("SaveFile.jsp");
    //打开文件
    poCtrl1.webOpen("doc/test.doc", OpenModeType.docRevisionOnly, "Tom");
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

    function AfterDocumentOpened() {
        refreshList();
    }

    //获取当前痕迹列表
    function refreshList() {
        var i;
        document.getElementById("ul_Comments").innerHTML = "";
        for (i = 1; i <= document.getElementById("PageOfficeCtrl1").Document.Revisions.Count; i++) {
            var str = "";
            str = str + document.getElementById("PageOfficeCtrl1").Document.Revisions.Item(i).Author;
            var revisionDate = document.getElementById("PageOfficeCtrl1").Document.Revisions.Item(i).Date;
            //转换为标准时间
            str = str + " " + dateFormat(revisionDate, "yyyy-MM-dd HH:mm:ss");

            if (document.getElementById("PageOfficeCtrl1").Document.Revisions.Item(i).Type == "1") {
                str = str + ' 插入：' + document.getElementById("PageOfficeCtrl1").Document.Revisions.Item(i).Range.Text;
            }
            else if (document.getElementById("PageOfficeCtrl1").Document.Revisions.Item(i).Type == "2") {
                str = str + ' 删除：' + document.getElementById("PageOfficeCtrl1").Document.Revisions.Item(i).Range.Text;
            }
            else {
                str = str + ' 调整格式或样式。';
            }
            document.getElementById("ul_Comments").innerHTML += "<li><a href='#' onclick='goToRevision(" + i + ")'>" + str + "</a></li>"
        }

    }

    //GMT时间格式转换为CST
    dateFormat = function (date, format) {
        date = new Date(date);
        var o = {
            'M+': date.getMonth() + 1, //month
            'd+': date.getDate(), //day
            'H+': date.getHours(), //hour
            'm+': date.getMinutes(), //minute
            's+': date.getSeconds(), //second
            'q+': Math.floor((date.getMonth() + 3) / 3), //quarter
            'S': date.getMilliseconds() //millisecond
        };

        if (/(y+)/.test(format))
            format = format.replace(RegExp.$1, (date.getFullYear() + '').substr(4 - RegExp.$1.length));

        for (var k in o)
            if (new RegExp('(' + k + ')').test(format))
                format = format.replace(RegExp.$1, RegExp.$1.length == 1 ? o[k] : ('00' + o[k]).substr(('' + o[k]).length));

        return format;
    }

    //定位到当前痕迹
    function goToRevision(index) {
        document.getElementById("PageOfficeCtrl1").Document.Revisions.Item(index).Range.Select();
    }

    //刷新列表
    function refresh_click() {
        refreshList();
    }

</script>
<div style=" width:1300px; height:700px;">
    <div id="Div_Comments" style=" float:left; width:200px; height:700px; border:solid 1px red;">
        <h3>痕迹列表</h3>
        <input type="button" name="refresh" value="刷新" onclick=" return refresh_click()"/>
        <ul id="ul_Comments">

        </ul>
    </div>
    <div style=" width:1050px; height:700px; float:right;">
        <%=poCtrl1.getHtmlCode("PageOfficeCtrl1")%>
    </div>
    </form>
</div>
</body>
</html>

