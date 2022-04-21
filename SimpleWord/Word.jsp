<%@ page language="java"
         import="com.zhuozhengsoft.pageoffice.OpenModeType,com.zhuozhengsoft.pageoffice.PageOfficeCtrl"
         pageEncoding="utf-8" %>
<%
    PageOfficeCtrl poCtrl = new PageOfficeCtrl(request);
    //设置服务器页面
    poCtrl.setServerPage(request.getContextPath() + "/poserver.zz");
    //添加自定义按钮
    poCtrl.addCustomToolButton("保存", "Save", 1);
    poCtrl.addCustomToolButton("另存到本地", "SaveAs", 12);
    poCtrl.addCustomToolButton("打印设置", "PrintSet", 0);
    poCtrl.addCustomToolButton("打印", "PrintFile", 6);
    poCtrl.addCustomToolButton("全屏/还原", "IsFullScreen", 4);
    poCtrl.addCustomToolButton("-", "", 0);
    poCtrl.addCustomToolButton("关闭", "Close", 21);
    poCtrl.setFileTitle("test");//设置另存为到本地的文件名称
    //设置保存页面
    poCtrl.setSaveFilePage("SaveFile.jsp");
    //打开Word文档
    poCtrl.webOpen("doc/test.doc", OpenModeType.docNormalEdit, "张佚名");
%>
<!DOCTYPE html>
<html>
<head lang="en">
    <meta charset="UTF-8">
    <title> 选项卡</title>
    <style type="text/css">
        /* CSS样式制作 */
        * {
            padding: 0px;
            margin: 0px;
        }

        a {
            text-decoration: none;
            color: black;
        }

        a:hover {
            text-decoration: none;
            color: #336699;
        }

        #tab {
            width: auto;
            padding: 5px;
            height: 150px;
            margin: 20px;
        }

        #tab ul {
            list-style: none;
            height: 30px;
            line-height: 30px;
            border-bottom: 2px #C88 solid;
        }

        #tab ul li {
            background: #FFF;
            cursor: pointer;
            float: left;
            list-style: none
            height: 29px;
            line-height: 29px;
            padding: 0px 10px;
            margin: 0px 3px;
            border: 1px solid #BBB;
            border-bottom: 2px solid #C88;
        }

        #tab ul li.on {
            border-top: 2px solid Saddlebrown;
            border-bottom: 2px solid #FFF;
        }

        #tab div {
            height: 700px;
            width: auto;
            line-height: 24px;
            border-top: none;
            padding: 1px;
            border: 1px solid #336699;
            padding: 10px;
        }

        .hide {
            display: none;
        }

        .show {
            display: block;
        }

        .page {
        }
    </style>
    <script type="text/javascript" src="../js/jquery-1.8.2.min.js"></script>
    <script type="text/javascript">
        // JS实现选项卡切换
        window.onload = function () {
            var myTab = document.getElementById("tab");    //整个div
            var myUl = myTab.getElementsByTagName("ul")[0]; //一个节点
            var myLi = myUl.getElementsByTagName("li");    //数组
            var myDiv = $(".page"); //myTab.getElementsByTagName("div"); //数组

            for (var i = 0; i < myLi.length; i++) {
                myLi[i].index = i;
                myLi[i].onclick = function () {
                    for (var j = 0; j < myLi.length; j++) {
                        myLi[j].className = "off";
                        myDiv[j].className = "hide";
                    }
                    this.className = "on";
                    myDiv[this.index].className = "show";
                }
            }
        }
    </script>
</head>
<body>
<script type="text/javascript">
    function Save() {
        document.getElementById("PageOfficeCtrl1").WebSave();
    }
   function SaveAs() {
         document.getElementById("PageOfficeCtrl1").ShowDialog(3);
    }

    function PrintSet() {
        document.getElementById("PageOfficeCtrl1").ShowDialog(5);
    }

    function PrintFile() {
        document.getElementById("PageOfficeCtrl1").ShowDialog(4);
    }

    function Close() {
        window.external.close();
    }

    function IsFullScreen() {
        document.getElementById("PageOfficeCtrl1").FullScreen = !document.getElementById("PageOfficeCtrl1").FullScreen;
    }

    //文档关闭前先提示用户是否保存
    function BeforeBrowserClosed() {
        if (document.getElementById("PageOfficeCtrl1").IsDirty) {
            if (confirm("提示：文档已被修改，是否继续关闭放弃保存 ？")) {
                return true;

            } else {

                return false;
            }
        }
    }
</script>
<div style=" text-align:center; margin:10px; background-color:#00BFFF; color:White; padding:8px;"><h1>XX单位XX办公系统</h1>
</div>
<!-- HTML页面布局 -->
<div id="tab">
    <ul>
        <li class="on">Word文件</li>
        <li class="off">文档属性</li>
        <li class="off">AA选项卡</li>
        <li class="off">BB选项卡</li>
        <li class="off">CC选项卡</li>
    </ul>
    <div id="firstPage" class="page show">
        <%=poCtrl.getHtmlCode("PageOfficeCtrl1")%>
    </div>
    <div id="secondPage" class="page hide" style="  ">
        <p style=" line-height:40px;">
            文件格式：word<br/>
            作者：张三<br/>
            XXXX：-----<br/>
            XXXX：---<br/>
            XXXX：------------------<br/>
            XXXX：xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
            xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
            xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
        </p>
    </div>
    <div id="AAPage" class="page hide">
        AA
    </div>
    <div id="BBPage" class="page hide">
        BB
    </div>
    <div id="CCPage" class="page hide">
        CC
    </div>
</div>
</body>
</html>
