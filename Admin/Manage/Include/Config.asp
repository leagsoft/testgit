<%
'///////////////////////////////////////////////////////////////////////////////////////////
'//
'//                                     ϵͳ������Ϣ
'//                                 (2003-06-01)
'///////////////////////////////////////////////////////////////////////////////////////////
'//////////////////////////////////////////
'//��Դ����ģ��������Ϣ

'//FileSystemObject��������
Const FsoObjectStr="Scripting.FileSystemObject"

'//��̬�ļ���չ��
Const ExNameOfNewsFile=".htm"

'//Tsysϵͳ��Ŀ¼��ַ
Const TsysRootPath = "/Admin"

'//һ������Ա�ʻ��Ƿ�����ͬʱ�����ͬIPʹ��
Const DubleOnlineUser=False

'//��̨��Դ�б��ÿҳ��ʾ��Ŀ
Const Sys_PageSize=20

'//������Դʱ����ķ������ű���ʱʱ��
Const CreateNewsFiles_ScriptTimeOut=900

'//��Դϵͳ����Ĺؼ�����Ŀ
Const NewsKeyWordListNum=3

'//���ɵ������Դ��Ŀ
Const RelateNewsNumber=10

'//��������Դʱ�Ƿ�ʹ��ģ�建��(����������ٶȼ��ȶ��ԣ�ͬʱ������С�����ڴ�,����ģ�����ݴ�С������)
Const Buffer_WhenCreatingFile=True

'//////////////////////////////////////////
'//����ģ��������Ϣ

'//ÿ�η������۵ļ��ʱ�䣨�룩
Const Comment_SubmitTime=10

'//Ĭ�����۹���ҳ��ʾ��С
Const Def_Comment_PageSize=30

'//////////////////////////////////////////
'//����������Ϣ

'//վ������
Const Def_MySiteTitle="�㶫ʡ�������Ϣ����ϵͳ����"
'//ϵͳ����
Const Def_SysTitle="�㶫ʡ�������Ϣ����ϵͳ V1.0"

'//ϵͳ����
Const Def_Developer="<a href='#'><font color=""#FFFFFF""><strong>�������Ϣͳ�ƴ��������</strong></font></a>"

'//��Ȩ��Ϣ
Const SoftCopyright_Script="\n����ְ�Ȩ����\n"

'//�Ƿ�Ϊ�´ε�¼�Զ���¼����Ա�ʻ�
Const IsAutoRemberLoginName=True

'//�Զ���¼����Ա�ʻ���Cookie�ĳ�ʱʱ�䣨�죩
Const AutoRemberLoginName_ExpiresTime=5

'//Ĭ����Դ��������Ƿ�չ��
Const Def_ShowNewsContorlPlane=False

'//��Դ�޸ĺ��Ƿ�����������
Const Def_ReCheckAfterModify=False

'//�˳���Ϣ����ϵͳʱ�Ƿ���ʾ
Const ConfirmWhenExitNewsSystem=True

'//�Ƿ����ù���Ա��¼��ȫ�Ǽ�
Const Def_UseLoginPolliceMan=True
'//���ӵ�ʱ�䷶Χ���룩
Const Def_StakeoutTimeRange=60
'//�����¼����Ĵ���
Const Def_EnableLoginWrong_Number=6
'//����ʱ�䣨�룩
Const Def_LoginWrongLockTimeRange=1000
'//����ʱ��
server.ScriptTimeout =30000

'
'//////////////////////////////////////////
'//�ļ�ϵͳ�����ļ�

'//�����ļ�Ŀ¼
'Const DirectoryRoot="../../UpLoadFiles"
Const DirectoryRoot="../../UpLoadFiles"

'//�����ϴ����ļ�����
Const FileSystem_EnableFileExt="|TXT|HTML|WMV|TXT|GIF|JPG|JPEG|BMP|PNG|DOC|XLS|TIF|RAR|ZIP|EXE|MHT|PPT|MP3|CHM|SWF|RM|PPS|"
%>