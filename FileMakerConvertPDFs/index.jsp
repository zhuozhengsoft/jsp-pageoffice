<%@ page language="java" pageEncoding="utf-8" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
    <title></title>
    <script type="text/javascript">
        var checkit = true;
        function selectall() {
            if (checkit) {
                var obj = document.all.check;
                for (var i = 0; i < obj.length; i++) {
                    obj[i].checked = true;
                    checkit = false;
                }
            } else {
                var obj = document.all.check;
                for (var i = 0; i < obj.length; i++) {
                    obj[i].checked = false;
                    checkit = true;
                }

            }

        }

        var count = 0;
        var strLength = 0;
        var str = new Array();

        function getID() {
            var ids = "";
            //循环获取选中checkbox的值
            var check = document.getElementsByName("check");
            for (var i = 0; i < check.length; i++) {
                if (check[i].checked) {
                    ids += check[i].value + ",";
                }
            }

            if (ids == "" || ids == null || ids == "on,") {
                alert("请至少选择一个要转换的文档！");
                return;

            }
            //去掉最后一个逗号
            ids = ids.substring(0, ids.length - 1);
            var i = ids.indexOf("on");
            if (i == 0 || i > 0) {
                //判读是否包含"on"，如果包含，则去掉
                ids = ids.replace("on,", "");
            }
            str = ids.split(",");
            if (count > 0) {
                alert("请刷新当前页面重新选择要转换的文档!");
                return;
            }
            //第一次自刷新
            document.getElementById("iframe1").src = "Convert.jsp?id=" + str[count];
            //设置进度条
            document.getElementById("ProgressBarSide").style.visibility = "visible";
            document.getElementById("ProgressBar").style.width = count+1+ "0%";
            strLength = str.length;
        }
        function convert() {
            count = count + 1;
            if (count < strLength) {
                document.getElementById("iframe1").src = "Convert.jsp?id=" + str[count];
                //设置进度条
                document.getElementById("ProgressBarSide").style.visibility = "visible";
                document.getElementById("ProgressBar").style.width = count+1+"0%";
                //document.body.insertBefore(document.getElementById("ProgressBarSide"), document.body.childNode[0];
            } else {
                //隐藏进度条div
                document.getElementById("ProgressBarSide").style.visibility = "hidden";
                count = 0;
                //重置进度条
                document.getElementById("ProgressBar").style.width = "0%";
                document.getElementById("aDiv").style.display = "";
                //alert("转换完成！");
            }
        }
    </script>
</head>
<body>
<%
    String url = request.getSession().getServletContext().getRealPath("FileMakerConvertPDFs/doc/" + "/");
%>
<div style="margin:100px" align="center">
    <form id="form1" method="post">
        <table id="table1" style="border-spacing:20px;border:1px;background-color: darkgray; width: 600px;"
               cellpadding="3" cellspacing="1">
            <h2>文件列表</h2>
            <tr>
                <td><input name="check" type="checkbox" onclick="selectall()"/></td>
                <td>序号</td>
                <td>文件名</td>
                <td>编辑</td>
            </tr>
            <tr>
                <td><input name="check" type="checkbox" value="1"/></td>
                <td>01</td>
                <td>PageOffice产品简介</td>
                <td><a href="Edit.jsp?id=1">编辑</a></td>
            </tr>
            <tr>
                <td><input name="check" type="checkbox" value="2"/></td>
                <td>02</td>
                <td>PageOffice产品安装步骤</td>
                <td><a href="Edit.jsp?id=2">编辑</a></td>
            </tr>
            <tr>
                <td><input name="check" type="checkbox" value="3"/></td>
                <td>03</td>
                <td>PageOffice产品应用领域</td>
                <td><a href="Edit.jsp?id=3">编辑</a></td>
            </tr>
            <tr>
                <td><input name="check" type="checkbox" value="4"/></td>
                <td>04</td>
                <td>PageOffice产品对环境的要求</td>
                <td><a href="Edit.jsp?id=4">编辑</a></td>
            </tr>
        </table>
        </br></br>
        <input type="button" value="批量转换成PDF文档" onclick="getID()"/>
    </form>
</div>
<div id="ProgressBarSide" style="color: Silver; width:500px; visibility: hidden;
        position: relative;  left: 40%; top:30%; margin-top: 22px">
    <span style="color: gray; font-size: 12px; text-align: left;">正在转换请稍后...</span><br/>
    <div id="ProgressBar" style="background-color: Green; height: 16px; width: 0%; border-width: 1px;
            border-style: Solid;">
    </div>
</div>
<div id="aDiv" style="display: none; color: Red; font-size: 15px;text-align:center;">
    <span>转换完成，可在下面的地址中查看扩展名为“.pdf”文件，具体的地址为:<%=url %></span>
</div>
<div style="width: 1px; height: 1px; overflow: hidden;">
    <iframe id="iframe1" name="iframe1" src=""></iframe>
</div>
</body>
</html>
