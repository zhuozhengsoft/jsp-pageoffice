<%@ page language="java"
	import="java.util.*,com.zhuozhengsoft.pageoffice.*,com.zhuozhengsoft.pageoffice.wordwriter.*"
	pageEncoding="utf-8"%>
<%
//******************************卓正PageOffice组件的使用*******************************
	PageOfficeCtrl poCtrl1 = new PageOfficeCtrl(request);
	poCtrl1.setServerPage(request.getContextPath()+"/poserver.zz"); //此行必须
	//隐藏菜单栏
	poCtrl1.setMenubar(false);
	poCtrl1.setCustomToolbar(false);
	poCtrl1.webOpen("doc/template.doc", OpenModeType.docNormalEdit, "张三");
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
  <head>
    <title>js 合并Word中的单元格</title>
    <script type="text/javascript">
//        // 向右合并一个单元格
//        function mergeCellRight() {
//            var docObj = document.getElementById("PageOfficeCtrl1").Document;
//            docObj.Tables(1).Cell(1, 1).Select();   // 选择单元格（1，1）
//            docObj.Application.Selection.MoveRight(1, 1, 1); // 向右选择一个单元格
//            docObj.Application.Selection.Cells.Merge(); // 合并
//        }
//        
//        // 向下合并一个单元格
//        function mergeCellDown() {
//            var docObj = document.getElementById("PageOfficeCtrl1").Document;
//            docObj.Tables(1).Cell(2, 2).Select();   // 选择单元格（2，2）
//            docObj.Application.Selection.MoveDown(5, 1, 1); // 向下移动1个单元格
//            docObj.Application.Selection.Cells.Merge(); // 合并
//        }

        // 向右合并一个单元格
        function mergeCellRight() {
            var mac = "Function myfunc()" + " \r\n"
                    + "ActiveDocument.Tables(1).Cell(1, 1).Select " + " \r\n" // 选择单元格（1，1）
                    + "Application.Selection.MoveRight 1, 1, 1 " + " \r\n"    // 向右选择一个单元格
                    + "Application.Selection.Cells.Merge " + " \r\n"          // 合并
                    + "End Function " + " \r\n";
            document.getElementById("PageOfficeCtrl1").RunMacro("myfunc", mac);
        }

        // 向下合并一个单元格
        function mergeCellDown() {
            var mac = "Function myfunc()" + " \r\n"
                    + "ActiveDocument.Tables(1).Cell(2, 2).Select " + " \r\n"// 选择单元格（2，2）
                    + "Application.Selection.MoveDown 5, 1, 1 " + " \r\n"    // 向下移动1个单元格
                    + "Application.Selection.Cells.Merge " + " \r\n"         // 合并
                    + "End Function " + " \r\n";
            document.getElementById("PageOfficeCtrl1").RunMacro("myfunc", mac);
        }
    </script>
</head>
<body>
    <form id="form1">
    <div style="font-size:12px; line-height:20px; border-bottom:dotted 1px #ccc;border-top:dotted 1px #ccc; padding:5px;">
    关键代码：点右键，选择“查看源文件”，看js函数<span style="background-color:Yellow;">mergeCellRight()和mergeCellDown()</span>
    </div>
    <div style=" font-size:small">
        <input id="Button1" type="button" value="向右合并一个单元格" onclick="mergeCellRight()"/>&nbsp;&nbsp;
        <input id="Button2" type="button" value="向下合并一个单元格" onclick="mergeCellDown()" /><br />
    </div>
    <div style="width: 1200px; height: 700px;">
       	        <%=poCtrl1.getHtmlCode("PageOfficeCtrl1")%>
    </div>
    </form>
</body>
</html>
