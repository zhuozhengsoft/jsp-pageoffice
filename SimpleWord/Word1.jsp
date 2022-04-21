<%@ page language="java"
         import="com.zhuozhengsoft.pageoffice.OpenModeType,com.zhuozhengsoft.pageoffice.PageOfficeCtrl"
         pageEncoding="utf-8" %>
<%
    // 获取文件磁盘路径
    String filePath = request.getSession().getServletContext().getRealPath("SimpleWord/doc/test.doc");
    PageOfficeCtrl poCtrl = new PageOfficeCtrl(request);
    //设置服务器页面
    poCtrl.setServerPage(request.getContextPath() + "/poserver.zz");
    //添加自定义按钮
    poCtrl.addCustomToolButton("保存", "Save", 1);
    //设置保存页面
    poCtrl.setSaveFilePage("SaveFile.jsp");
    //打开文件
    //这里用磁盘路径方式打开文件，为了支持linx服务器，文档路径前面必须加上“file://”前缀
    if(System.getProperty("os.name").toLowerCase().contains("linux")){    //判断该当前系统是否是linux操作系统
        poCtrl.webOpen("file://"+filePath, OpenModeType.docNormalEdit, "张佚名");
    }else{
        poCtrl.webOpen(filePath, OpenModeType.docNormalEdit, "张佚名");
    }
%>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html>
<head>
    <meta charset="utf-8">
    <title>XX文档系统</title>
    <link rel="stylesheet" href="css/style.css" type="text/css">
    <script type="text/javascript">
        document.createElement("section");
        document.createElement("article");
        document.createElement("footer");
        document.createElement("header");
        document.createElement("hgroup");
        document.createElement("nav");
        document.createElement("menu");
    </script>
</head>
<body>
<script type="text/javascript">
    function Save() {
        document.getElementById("PageOfficeCtrl1").WebSave();
    }
</script>
<header>
    <div class="w12 header">
        <a class="db logo fl"><img src="images/logo.jpg" width="327" height="94" alt=""/></a>
        <div class="fr logofr"><a href="#" class="blk">返回首页</a> |<a href="#" class="blk">客服中心</a><br>
            如注册遇到问题请拨打：<strong style="font-size:14px;">400-000-0000</strong></div>
    </div>
</header>
<div class="head_border"></div>
<section class="w12 login">
    <em class="fr">当前用户：张三 </em>
</section>
<section class="main w12">
    <div class="title"><a class="title1 db fl">文档内容</a><a class="title2 db fl">其他信息</a></div>
    <div class="fr tit2"><span class="arr"></span></div>
    <div style=" width:auto; height:700px;">
        <%=poCtrl.getHtmlCode("PageOfficeCtrl1")%>
    </div>
</section>
<br/><br/>
<div style=" text-align:center; height:80px; border-top: solid 1px #666; line-height:70px;">Copyright &copy 2019
    北京卓正志远软件有限公司
</div>
</body>
</html>

