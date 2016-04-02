<%@ page language="java" import="java.util.*" pageEncoding="utf-8"%>
<%@ taglib prefix="c" uri="http://java.sun.com/jsp/jstl/core"%>
<%@ taglib prefix="fmt" uri="http://java.sun.com/jsp/jstl/fmt"%>
<%@ taglib prefix="fn" uri="http://java.sun.com/jsp/jstl/functions"%>
<%
    String path = request.getContextPath();
    String basePath = request.getScheme() + "://"
            + request.getServerName() + ":" + request.getServerPort()
            + path + "/";
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<base href="<%=basePath%>">
<html>
<head>
    <title>Title</title>
    <script type="text/javascript" src="${pageContext.request.contextPath}/res/jquery/jquery-1.8.1.min.js" ></script>
    <script type="text/javascript" src="${pageContext.request.contextPath}/res/uploadify/jquery.uploadify.min.js"></script>
    <script type="text/javascript" src="${pageContext.request.contextPath}/uploadfile.js"></script>
    <link rel="stylesheet" type="text/css" href="${pageContext.request.contextPath}/res/uploadify/uploadify.css">
</head>
<body>
<div id="queue"></div>
<input type="file" id="file_upload">
<input type="button" value="开始上传" onclick="javascript:$('#file_upload').uploadify('upload','*')">
<input type="button" value="取消上传" onclick="javascript:$('#file_upload').uploadify('cancel','*')">
</body>
</html>
