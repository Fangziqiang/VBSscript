'����� QTP ֱ�Ӵ��� AOM �����ǻᱨ��ģ���Ϊ QTP ֻ������һ��ʵ�����󣬵�����
'�Ѿ�������һ�� QTP ֮��Ͳ������ٴ�������һ�� QTP �ˣ�������ʱ�����ǿ���ֱ��
'ʹ�� GETOBJECT("","quicktest.application")���Ե�ǰ������ QTP ���в���,�����������
'ͼ�ű������ǵ�����к�ǰ QTP �ͻ��Զ����أ���������Զ��ָ��ɼ�.
'Set qtapp=GetObject("","QuickTest.Application")
'qtapp.visible = false

Dim qtApp

'��������
Set qtApp = CreateObject("QuickTest.Application")

'Start QuickTest
qtApp.Launch

'����Ϊ�ɼ�
qtApp.Visible = True

'�򿪲��Խű�
qtApp.Open("E:\UFTworkstation\GUITest1")

'���в��Խű�
qtApp.Test.Run

qtApp.Quit
Set qtApp=Nothing
