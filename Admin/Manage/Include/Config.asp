<%
'///////////////////////////////////////////////////////////////////////////////////////////
'//
'//                                     系统配制信息
'//                                 (2003-06-01)
'///////////////////////////////////////////////////////////////////////////////////////////
'//////////////////////////////////////////
'//资源管理模块配制信息

'//FileSystemObject对象名称
Const FsoObjectStr="Scripting.FileSystemObject"

'//静态文件扩展名
Const ExNameOfNewsFile=".htm"

'//Tsys系统主目录地址
Const TsysRootPath = "/Admin"

'//一个管理员帐户是否允许同时多个不同IP使用
Const DubleOnlineUser=False

'//后台资源列表的每页显示数目
Const Sys_PageSize=20

'//生成资源时允许的服务器脚本超时时间
Const CreateNewsFiles_ScriptTimeOut=900

'//资源系统允许的关键词数目
Const NewsKeyWordListNum=3

'//生成的相关资源数目
Const RelateNewsNumber=10

'//在生成资源时是否使用模板缓冲(将提高生成速度及稳定性；同时会牺牲小部分内存,依据模板内容大小而定；)
Const Buffer_WhenCreatingFile=True

'//////////////////////////////////////////
'//评论模块配制信息

'//每次发表评论的间隔时间（秒）
Const Comment_SubmitTime=10

'//默认评论管理页显示大小
Const Def_Comment_PageSize=30

'//////////////////////////////////////////
'//其它配制信息

'//站点名称
Const Def_MySiteTitle="广东省银监局信息发布系统（）"
'//系统标题
Const Def_SysTitle="广东省银监局信息发布系统 V1.0"

'//系统标题
Const Def_Developer="<a href='#'><font color=""#FFFFFF""><strong>银监局信息统计处开发完成</strong></font></a>"

'//版权信息
Const SoftCopyright_Script="\n银监局版权所有\n"

'//是否为下次登录自动记录管理员帐户
Const IsAutoRemberLoginName=True

'//自动记录管理员帐户名Cookie的超时时间（天）
Const AutoRemberLoginName_ExpiresTime=5

'//默认资源操作面板是否展开
Const Def_ShowNewsContorlPlane=False

'//资源修改后是否需重新审枋
Const Def_ReCheckAfterModify=False

'//退出信息发布系统时是否提示
Const ConfirmWhenExitNewsSystem=True

'//是否启用管理员登录安全登记
Const Def_UseLoginPolliceMan=True
'//监视的时间范围（秒）
Const Def_StakeoutTimeRange=60
'//允许登录错误的次数
Const Def_EnableLoginWrong_Number=6
'//被封时间（秒）
Const Def_LoginWrongLockTimeRange=1000
'//过期时间
server.ScriptTimeout =30000

'
'//////////////////////////////////////////
'//文件系统配制文件

'//虚拟文件目录
'Const DirectoryRoot="../../UpLoadFiles"
Const DirectoryRoot="../../UpLoadFiles"

'//允许上传的文件类型
Const FileSystem_EnableFileExt="|TXT|HTML|WMV|TXT|GIF|JPG|JPEG|BMP|PNG|DOC|XLS|TIF|RAR|ZIP|EXE|MHT|PPT|MP3|CHM|SWF|RM|PPS|"
%>