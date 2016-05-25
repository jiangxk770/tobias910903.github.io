<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%
' FileName="Connection_ado_conn_string.htm"
' Type="ADO" 
' DesigntimeType="ADO"
' HTTP="true"
' Catalog=""
' Schema=""
Dim MM_beijing_STRING
MM_beijing_STRING = "Driver={Microsoft Access Driver (*.mdb)};DBQ="&Server.MapPath("beijing.mdb")
%>
<%
Dim MM_editAction
MM_editAction = CStr(Request.ServerVariables("SCRIPT_NAME"))
If (Request.QueryString <> "") Then
  MM_editAction = MM_editAction & "?" & Server.HTMLEncode(Request.QueryString)
End If

' boolean to abort record edit
Dim MM_abortEdit
MM_abortEdit = false
%>
<%
If (CStr(Request("MM_insert")) = "form1") Then
  If (Not MM_abortEdit) Then
    ' execute the insert
    Dim MM_editCmd

    Set MM_editCmd = Server.CreateObject ("ADODB.Command")
    MM_editCmd.ActiveConnection = MM_beijing_STRING
    MM_editCmd.CommandText = "INSERT INTO pl (plcont) VALUES (?)" 
    MM_editCmd.Prepared = true
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param1", 203, 1, 1073741823, Request.Form("plcont")) ' adLongVarWChar
    MM_editCmd.Execute
    MM_editCmd.ActiveConnection.Close

    ' append the query string to the redirect URL
    Dim MM_editRedirectUrl
    MM_editRedirectUrl = "index.asp"
    If (Request.QueryString <> "") Then
      If (InStr(1, MM_editRedirectUrl, "?", vbTextCompare) = 0) Then
        MM_editRedirectUrl = MM_editRedirectUrl & "?" & Request.QueryString
      Else
        MM_editRedirectUrl = MM_editRedirectUrl & "&" & Request.QueryString
      End If
    End If
    Response.Redirect(MM_editRedirectUrl)
  End If
End If
%>
<%
Dim pl
Dim pl_cmd
Dim pl_numRows

Set pl_cmd = Server.CreateObject ("ADODB.Command")
pl_cmd.ActiveConnection = MM_beijing_STRING
pl_cmd.CommandText = "SELECT * FROM pl ORDER BY pltime DESC" 
pl_cmd.Prepared = true

Set pl = pl_cmd.Execute
pl_numRows = 0
%>
<%
Dim Repeat1__numRows
Dim Repeat1__index

Repeat1__numRows = -1
Repeat1__index = 0
pl_numRows = pl_numRows + Repeat1__numRows
%>
<!doctype html>
<html>
<head>
<meta charset="utf-8">
<title>漂在北京 - 记录我的北漂生活</title>
<meta name="keywords" content="北漂生活,晓雨个人网站,漂在北京" />
<meta name="description" content="每个人都有自己的梦想，无论梦想是大是小，是尊贵还是卑微，它都是人们心中最崇高的向往。" />
<meta http-equiv="X-UA-Compatible" content="IE=edge" />
<meta name="viewport" content="width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no, minimal-ui">
<link href="style.css" rel="stylesheet" type="text/css" />
<link rel="shortcut icon" href="/images/dlz.ico" />
</head>

<body>
<!-- nav -->
<div class="nav">
		<div class="nav_box">
				<div class="logo" title="网站首页"><a href="http://www.ml22.cn/"><img src="/images/logo-v2.png" width="100" alt="logo"></a></div>
				<ul>
						<li>
								<select id="checked_time">
										<!-- 2017年
										<option disabled>2017年</option>
										<option value="201712">2017年12月</option>
										<option value="201711">2017年11月</option>
										<option value="201710">2017年10月</option>
										<option value="201709">2017年09月</option>
										<option value="201708">2017年08月</option>
										<option value="201707">2017年07月</option>
										<option value="201706">2017年06月</option>
										<option value="201705">2017年05月</option>
										<option value="201704">2017年04月</option>
										<option value="201703">2017年03月</option>
										<option value="201702">2017年02月</option>
										<option value="201701">2017年01月</option>
										-->
										<option disabled>—2016年—</option>
										<!--
										<option value="201612">2016年12月</option>
										<option value="201611">2016年11月</option>
										<option value="201610">2016年10月</option>
										<option value="201609">2016年09月</option>
										<option value="201608">2016年08月</option>
										<option value="201607">2016年07月</option>
										<option value="201606">2016年06月</option>
										-->
										<option value="201605">2016年05月</option>
										<option value="201604">2016年04月</option>
										<option value="201603">2016年03月</option>
										<option value="201602">2016年02月</option>
										<option value="201601">2016年01月</option>
								</select>
						</li>
				</ul>
		</div>
</div>
<!-- //nav --> 
<!-- 视频背景 -->
<div class="indexbg">
		<video autoplay muted loop poster="images/indexbg.jpg">
				<source src="indexbg.mp4" type="video/mp4">
		</video>
		<div class="popover"></div>
</div>
<!-- // 视频背景 -->
<div class="open_cont">
		<div class="open_list" id="open_cont"><img src="images/index_nav_01.jpg"><p>&nbsp;</p></div>
		<div class="open_list"><a href="http://www.ml22.cn/"><img src="images/index_nav_02.jpg"><p>晓雨个人网站</p></a></div>
		<div class="open_list"><a href="#"><img src="images/index_nav_03.jpg"><p>更多推荐</p></a></div>
</div>
<!-- content -->
<div class="content">
		<div class="content_box"></div>
		<ul class="pageturn">
				<li id="next"></li>
				<li id="pre"></li>
				<li id="comment_switch">参与评论</li>
		</ul>
</div>
<!-- // content --> 
<!-- comment -->
<div class="comment" id="comment">
		<form ACTION="<%=MM_editAction%>" name="form1" METHOD="POST">
				<input type="text" id="plcont" name="plcont" value="" placeholder="请输入评论内容，不超过100个字符。" maxlength="100">
				<input type="submit" id="bid" value="评 论">
				<input type="hidden" name="MM_insert" value="form1">
		</form>
		<ul class="pllist">
				<% 
While ((Repeat1__numRows <> 0) AND (NOT pl.EOF)) 
%>
						<li>游客<%=(pl.Fields.Item("plid").Value)%>&nbsp;[<%=(pl.Fields.Item("pltime").Value)%>]：<%=(pl.Fields.Item("plcont").Value)%></li>
						<% 
  Repeat1__index=Repeat1__index+1
  Repeat1__numRows=Repeat1__numRows-1
  pl.MoveNext()
Wend
%>
		</ul>
</div>
<!-- // comment --> 
<!-- footer -->
<div class="footer">
		<p> <a href="http://www.ml22.cn/dlz_about.asp" target="_blank">关于我们</a> | <a href="http://www.ml22.cn/dlz_about.asp" target="_blank">联系方式</a> | <a href="http://www.ml22.cn/dlz_note.asp" target="_blank">心情随笔</a> | <a href="http://www.ml22.cn/" target="_blank">网站首页</a> </p>
		<p>漂在北京：用照片、文字、视频等等记录自己的北漂生活。<span>有一天，你一定会告别这样的生活,因为你在荒凉的地方仍开出了一朵花，我看到了。</span></p>
		<p><span>Copyright 2015-2016 - </span>All Rights Reserved.©市仔村晓雨<span> 琼ICP备15001175号-1</span></p>
		<p>
				<iframe frameborder="no" border="0" marginwidth="0" marginheight="0" width=310 height=52 src="http://music.163.com/outchain/player?type=0&id=383844672&auto=1&height=32"></iframe>
		</p>
</div>
<!-- //footer --> 
<script src="jquery-1.7.2.min.js"></script> 
<script>
var $tvalue = $("#checked_time").val();

$.ajax({
	type: "GET",
	url: "json_page.js",
	dataType: "json",
	success: function(data) {

		$("#checked_time").val(data.page[1].time); // 索引为0的是默认没有数据，所以这里从第一个开始
		function pageList() {
			var $tvalue = $("#checked_time").val();
			var $dlength = data.page.length;
			for (var i = 1; i < $dlength-1; i++) {
				if (data.page[i].time == $tvalue) {
					//alert(i);
					$("#pre").text(data.page[i + 1].time);
					$("#next").text(data.page[i - 1].time);
				}
			};
			$("html,body").animate({
				scrollTop: 0
			});
		};

		// 翻页效果
		$("#pre,#next").click(function() {
			$(".content_box").load("page/" + $(this).text() + '.html');
			$("#checked_time").val($(this).text());
			pageList();
		});

		// 顶部发生变化时加载文件
		$("#checked_time").on('change',
		function() {
			var $tvalue = $("#checked_time").val();
			
			$(".open_cont,.indexbg").fadeOut(function() {
				$(".indexbg").remove();
			});
			$(".content").fadeIn(function() {
				$(".content_box").load("page/" + $tvalue + ".html");
			});
			pageList();
		});

		// 按钮点击时加载JSON文件中最新的一个 索引从1开始 0已经被占用
		$("#open_cont").click(function() {
			$(".open_cont,.indexbg").fadeOut(function() {
				$(".indexbg").remove();
			});
			$(".content").fadeIn(function() {
				$(".content_box").load("page/" + data.page[1].time + '.html',
				function() {
					pageList();
				});
			});
		});
	},
	error: function(msd) {
		alert("发生错误：" + msd.status)
	}
});

$("#comment_switch").toggle(function() {
	$("#comment").slideDown();
	$("#comment_switch").text("收起回复").css({
		"background-color": "#660000",
		"color": "#ffffff"
	});
},
function() {
	$("#comment").slideUp();
	$("#comment_switch").text("参与评论").css({
		"background-color": "",
		"color": ""
	});
});

function btncss() {
	if ($("#plcont").val().length == 0) {
		$("#bid").attr("disabled", true);
		$("#bid").css({
			"color": "#999999",
			"border-color": "#999999"
		});
	} else if ($("#plcont").val().length >= 100) {
		alert("不超过100个字符！");
	} else {
		$("#bid").attr("disabled", false);
		$("#bid").css({
			"color": "#660000",
			"border-color": "#660000"
		});
	}
};

$(function() {
	btncss();
	$("#open_cont p").text("漂在北京 v." + $tvalue);
});

$("#plcont").keyup(function() {
	btncss();
});

// 百度统计
var _hmt = _hmt || []; (function() {
	var hm = document.createElement("script");
	hm.src = "//hm.baidu.com/hm.js?b3ce7d44b7cc964167120d5c2f4be5d5";
	var s = document.getElementsByTagName("script")[0];
	s.parentNode.insertBefore(hm, s);
})();
</script>
</body>
</html>
<%
pl.Close()
Set pl = Nothing
%>
