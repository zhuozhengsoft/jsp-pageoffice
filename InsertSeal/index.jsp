<%@ page language="java" pageEncoding="UTF-8" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
    <title>盖章和签字专题</title>
    <meta http-equiv="pragma" content="no-cache">
    <meta http-equiv="cache-control" content="no-cache">
    <meta http-equiv="expires" content="0">
    <meta http-equiv="keywords" content="keyword1,keyword2,keyword3">
    <meta http-equiv="description" content="This is my page">
    <link href="css/style.css" rel="stylesheet" type="text/css"/>
    <!-- pageoffice.js文件一定要引用-->
    <script type="text/javascript" src="../pageoffice.js"></script>
    <script type="text/javascript">
        function refreshDemo() {
            if (confirm("确认复位所有示例的文档吗？")) {
                location.href = "refresh.jsp";
            }
        }
    </script>
</head>
<body>
<!--content-->
<div class="zz-content mc clearfix pd-28" align="center">
    <div class="demo mc">
        <h2 class="fs-16">
            盖章和(手写)签字专题
        </h2>
    </div>
    <div style="margin: 10px" align="center">
        <p>
            点击
            <a href="../adminseal.zz" target="_blank">电子印章简易管理平台</a> 添加、删除印章。管理员:admin 密码:111111
            <span  style="font-size:12px;color: red">(为了系统使用印章的安全性，强烈建议修改此登录密码！！)</span>
        </p>
        <p>
            点击右侧的“全部复位”连接，即可重新演示盖章和签字效果：<a href="javascript:refreshDemo();">全部复位</a>
        </p>
        <p style="color: red;">
            印章测试用户名：李志，密码：111111或123456
        </p>
        <p style="color: red;">
          补充： HTML网页签章请参考<a style="color: red" href="http://www.zhuozhengsoft.com/down5/Java/BigDemo/huiqiandan2.rar">三、16、“汇签单2”效果（专业版、企业版）</a>
        </p>
    </div>
    <div style="margin: 10px" align="center">
        <h3>
            一、Word盖章
        </h3>
        <table style="border-collapse: separate; border-spacing: 20px;border: 1px solid #9CF " align="center"
               width="1200px">
            <tr>
                <th style="border-bottom: 1px solid #9CF;width:85%;">功能演示</th>
                <th style="border-bottom: 1px solid #9CF ">文件目录</th>
            </tr>
            <tr>
                <td>
                    1.常规盖章：
                    <a
                            href="javascript:POBrowser.openWindowModeless('Word/AddSeal1/Word1.jsp','width=1200px;height=800px;');">Word1.jsp</a>
                </td>
                <td>
                    (Word/AddSeal1)
                </td>
            </tr>
            <tr>
                <td>
                    2.无需输入用户名、密码盖章：
                    <a
                            href="javascript:POBrowser.openWindowModeless('Word/AddSeal2/Word2.jsp?id=1','width=1200px;height=800px;');">Word2.jsp</a>
                </td>
                <td>
                    (Word/AddSeal2)
                </td>
            </tr>
            <tr>
                <td>
                    3.无需输入用户名、密码，并且不弹出印章选择框盖章：
                    <a
                            href="javascript:POBrowser.openWindowModeless('Word/AddSeal3/Word3.jsp','width=1200px;height=800px;');">Word3.jsp</a>
                </td>
                <td>
                    (Word/AddSeal3)
                </td>
            </tr>
            <tr>
                <td>
                    4.编辑模版 - 在模版中添加盖章位置：
                    <a
                            href="javascript:POBrowser.openWindowModeless('Word/AddSeal4/Word4.jsp','width=1200px;height=800px;');">Word4.jsp</a>
                </td>
                <td>
                    (Word/AddSeal4)
                </td>
            </tr>
            <tr>
                <td>
                    5.常规指定位置盖章，加盖印章到模板中的指定位置：
                    <a
                            href="javascript:POBrowser.openWindowModeless('Word/AddSeal5/Word5.jsp','width=1200px;height=800px;');">Word5.jsp</a>
                </td>
                <td>
                    (Word/AddSeal5)
                </td>
            </tr>
            <tr>
                <td>
                    6.无需输入用户名、密码加盖印章到模板中的指定位置：
                    <a
                            href="javascript:POBrowser.openWindowModeless('Word/AddSeal6/Word6.jsp','width=1200px;height=800px;');">Word6.jsp</a>
                </td>
                <td>
                    (Word/AddSeal6)
                </td>
            </tr>
            <tr>
                <td>7.无需输入用户名、密码，并且不弹出印章选择框加盖印章到模板中的指定位置：
                    <a
                            href="javascript:POBrowser.openWindowModeless('Word/AddSeal7/Word7.jsp','width=1200px;height=800px;');">Word7.jsp</a>
                </td>
                <td>
                    (Word/AddSeal7)
                </td>
            </tr>

            <tr>
                <td>
                    8.特殊盖章需求实现：盖章后印章不保护文档内容，用户仍可编辑修改，印章不会出现失效字样：
                    <a
                            href="javascript:POBrowser.openWindowModeless('Word/AddSeal8/Word8.jsp','width=1200px;height=800px;');">Word8.jsp</a>
                </td>
                <td>
                    (Word/AddSeal8)
                </td>
            </tr>

            <tr>
                <td>
                    9.批量盖章：
                    <a
                            href="javascript:POBrowser.openWindowModeless('Word/BatchAddSeal/Default.jsp','width=1200px;height=800px;');">Default.jsp</a>
                </td>
                <td>
                    (Word/BatchAddSeal)
                </td>
            </tr>
            <tr>
                <td>
                    10.删除印章：
                    <a
                            href="javascript:POBrowser.openWindowModeless('Word/AddSeal9/Word9.jsp','width=1200px;height=800px;');">Word9.jsp</a>
                </td>
                <td>
                    (Word/AddSeal9)
                </td>
            </tr>
             <tr>
                <td>
                    11.加盖骑缝章：
                    <a
                            href="javascript:POBrowser.openWindowModeless('Word/AddSeal10/Word10.jsp','width=1200px;height=800px;');">Word10.jsp</a>
                </td>
                <td>
                    (Word/AddSeal10)
                </td>
            </tr>

        </table>
        <div style="margin: 10px" align="center">
            <h3>二、Word(手写)签字</h3>
            <table style="border-collapse: separate; border-spacing: 20px;border: 1px solid #9CF" align="center"
                   width="1200px">
                <tr>
                    <th style="border-bottom: 1px solid #9CF; width:85%;">功能演示</th>
                    <th style="border-bottom: 1px solid #9CF ">文件目录</th>
                </tr>
                <tr>
                    <td>
                        1.常规(手写)签字：
                        <a
                                href="javascript:POBrowser.openWindowModeless('Word/AddSign1/Word1.jsp','width=1200px;height=800px;');">Word1.jsp</a>
                    </td>
                    <td>
                        (Word/AddSign1)
                    </td>
                </tr>
                <tr>
                    <td>
                        2.无需输入用户名密码签字：
                        <a
                                href="javascript:POBrowser.openWindowModeless('Word/AddSign2/Word2.jsp','width=1200px;height=800px;');">Word2.jsp</a>
                    </td>
                    <td>
                        (Word/AddSign2)
                    </td>
                </tr>
                <tr>
                    <td>
                        3.(手写)签字到文档指定位置：
                        <a
                                href="javascript:POBrowser.openWindowModeless('Word/AddSign3/Word3.jsp','width=1200px;height=800px;');">Word3.jsp</a>
                    </td>
                    <td>
                        (Word/AddSign3)
                    </td>
                </tr>
                <tr>
                    <td>
                        4.无需输入用户名、密码，(手写)签字到模板中的指定位置（参考一、4示例在模板中添加签字位置）：
                        <a
                                href="javascript:POBrowser.openWindowModeless('Word/AddSign4/Word4.jsp','width=1200px;height=800px;');">Word4.jsp</a>
                    </td>
                    <td>
                        (Word/AddSign4)
                    </td>
                </tr>
                <tr>
                    <td>
                        5.特殊(手写)签字需求实现：签批后(手写)签字不保护文档内容，用户仍可编辑修改，(手写)签字不会出现失效字样：
                        <a
                                href="javascript:POBrowser.openWindowModeless('Word/AddSign5/Word5.jsp','width=1200px;height=800px;');">Word5.jsp</a>
                    </td>
                    <td>
                        (Word/AddSign5)
                    </td>
                </tr>
            </table>
        </div>
        <div style="margin: 10px" align="center">
            <h3>三、PDF盖章、签字<span style=" color:Red;">（专业版、企业版）</span></h3>
            <table style="border-collapse: separate; border-spacing: 20px;border: 1px solid #9CF" align="center" width="1200px">
                <tr>
                    <th style="border-bottom: 1px solid #9CF; width:85%;">功能演示</th>
                    <th style="border-bottom: 1px solid #9CF">文件目录</th>
                </tr>
            <tr>
                <td>
                    1.常规盖章：
                    <a
                            href="javascript:POBrowser.openWindowModeless('PDF/AddSeal1/PDF1.jsp','width=1200px;height=850px;');">PDF1.jsp</a>
                </td>
                <td>
                    (PDF/AddSeal1)
                </td>
            </tr>
            <tr>
                <td>
                    2.常规签字：
                    <a
                            href="javascript:POBrowser.openWindowModeless('PDF/AddSign1/PDF2.jsp' ,'width=1200px;height=850px;');">PDF2.jsp</a>
                </td>
                <td>
                    (PDF/AddSign1)
                </td>
            </tr>
        </table>
        <div style="margin: 10px" align="center">
            <h3>
                四、Excel盖章
            </h3>
            <table style="border-collapse: separate; border-spacing: 20px;border: 1px solid #9CF " align="center"
                   width="1200px">
                <tr>
                    <th style="border-bottom: 1px solid #9CF; width:85%;">功能演示</th>
                    <th style="border-bottom: 1px solid #9CF ">文件目录</th>
                </tr>
                <tr>
                    <td>
                        1.常规盖章：
                        <a
                                href="javascript:POBrowser.openWindowModeless('Excel/AddSeal1/Excel1.jsp','width=1200px;height=850px;');">Excel1.jsp</a>
                    </td>
                    <td>
                        (Excel/AddSeal1)
                    </td>
                </tr>

                <tr>
                    <td>
                        2.无需输入用户名密码盖章：
                        <a
                                href="javascript:POBrowser.openWindowModeless('Excel/AddSeal2/Excel2.jsp','width=1200px;height=800px;');">Excel2.jsp</a>
                    </td>
                    <td>
                        (Excel/AddSeal2)
                    </td>
                </tr>
                <tr>
                    <td>
                        3.无需输入用户名密码，并且不显示印章选择对话框：
                        <a
                                href="javascript:POBrowser.openWindowModeless('Excel/AddSeal3/Excel3.jsp','width=1200px;height=800px;');">Excel3.jsp</a>
                    </td>
                    <td>
                        (Excel/AddSeal3)
                    </td>
                </tr>
                <tr>
                    <td>
                        4.特殊盖章需求实现：盖章后印章不保护文档内容，用户仍可编辑修改，印章不会出现失效字样：
                        <a
                                href="javascript:POBrowser.openWindowModeless('Excel/AddSeal4/Excel4.jsp?id=1','width=1200px;height=800px;');">Excel4.jsp</a>
                    </td>
                    <td>
                        (Excel/AddSeal4)
                    </td>
                </tr>
                <tr>
                    <td>
                        5.删除印章：
                        <a
                                href="javascript:POBrowser.openWindowModeless('Excel/AddSeal5/Excel5.jsp','width=1200px;height=800px;');">Excel5.jsp</a>
                    </td>
                    <td>
                        (Excel/AddSeal5)
                    </td>
                </tr>

            </table>
        </div>
        <div style="margin: 10px" align="center">
            <h3>五、Excel(手写)签字</h3>
            <table style="border-collapse: separate; border-spacing: 20px;border: 1px solid #9CF" align="center"
                   width="1200px">
                <tr>
                    <th style="border-bottom: 1px solid #9CF; width:85%;">功能演示</th>
                    <th style="border-bottom: 1px solid #9CF ">文件目录</th>
                </tr>
                <tr>
                    <td>
                        1.常规(手写)签字：
                        <a
                                href="javascript:POBrowser.openWindowModeless('Excel/AddSign1/Excel1.jsp','width=1200px;height=850px;');">Excel1.jsp</a>
                    </td>
                    <td>
                        (Excel/AddSign1)
                    </td>
                </tr>
                <tr>
                    <td>
                        2.无需输入用户名密码(手写)签字：
                        <a
                                href="javascript:POBrowser.openWindowModeless('Excel/AddSign2/Excel2.jsp','width=1200px;height=850px;');">Excel2.jsp</a>
                    </td>
                    <td>
                        (Excel/AddSign2)
                    </td>
                </tr>
                <tr>
                    <td>
                        3.特殊(手写)签字需求实现：(手写)签批后签字不保护文档内容，用户仍可编辑修改，签字不会出现失效字样：
                        <a
                                href="javascript:POBrowser.openWindowModeless('Excel/AddSign3/Excel3.jsp','width=1200px;height=800px;');">Excel3.jsp</a>
                    </td>
                    <td>
                        (Excel/AddSign3)
                    </td>
                </tr>
            </table>
        </div>
    </div>
</div>
</body>
</html>
