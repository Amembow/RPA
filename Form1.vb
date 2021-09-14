


Option Strict Off
Option Explicit On
Imports System.Runtime.InteropServices
Imports System.Diagnostics
Imports System.Text
Imports System
Imports System.Collections.Generic
Imports System.Windows.Automation
Imports System.Linq
Imports System.Threading
Imports System.IO
Imports System.Collections
Imports System.Windows.Automation.Provider
Imports System.Windows.Automation.Text





Public Class Form1
    'クラス名、キャプションから子ウィンドウのハンドルを取得
    Private Declare Function FindWindowEx Lib "user32.dll" Alias "FindWindowExA" (ByVal hwndParent As Integer, ByVal hwndChildAfter As Integer, ByVal lpClassName As String, ByVal lpWindowName As String) As Integer
    Private Declare Function SendMessage Lib "user32.dll" Alias "SendMessageA" (ByVal hWnd As Integer, ByVal Msg As Integer, ByVal wParam As Integer, ByVal lParam As String) As Integer
    Public Declare Function FindWindow Lib "user32.dll" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Integer
    Public Declare Function GetDlgItem Lib "user32.dll" Alias "GetDlgItem" (ByVal hDlg As Integer, ByVal nIDDlgItem As Integer) As Integer
    Private Declare Function PostMessage Lib "user32.dll" Alias "PostMessageA" (ByVal hwnd As IntPtr, ByVal wMsg As Integer, ByVal wParam As Integer, ByVal lParam As IntPtr) As IntPtr


    Public Const WM_LBUTTONDOWN As Integer = &H201
    Public Const WM_LBUTTONUP As Integer = &H202
    Public Const CONTROLID As Integer = 1
    Private Const WM_IME_CHAR As Integer = &H286S '文字コード送信
    Private Const WM_SETTEXT As Integer = &HCS '文字列送信
    Private Const VK_RETURN As Integer = &HD
    Private Const WM_KEYDOWN As Integer = &H100
    Private Const WM_KEYUP As Integer = &H101
    Private Const WM_CHAR As UInteger = &H102
    Private Const CB_SELECTSTRING As Integer = &H14D
    Private Const WM_COMMAND As Integer = &H111
    Private Const CBN_SELCHANGE As Integer = &H10000
    Private Const SC_CLOSE As Integer = &HF060
    Private Const WM_SYSCOMMAND As Integer = &H112

    Private uiAuto As UIAutomationClient.CUIAutomation
    Private Const UIA_WindowControlTypeId As Integer = 50032
    Private Const UIA_ControlTypePropertyId As Integer = 30003
    Private Const UIA_ClassNamePropertyId As Integer = 30012
    Private Const UIA_TextControlTypeId As Integer = 50020
    Private Const TreeScope_Subtree As Integer = 7
    Private Const UIA_InvokePatternId As Integer = 10000
    Private Const UIA_AutomationPropertyID As Integer = 30011
    Private Const UIA_NamePropertyID As Integer = 30005

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click


        Dim fso As Object
        Dim inputfile As Object
        Dim LineStr As String
        Dim SP() As String
        Dim SP1() As String
        Dim i As Long

        fso = CreateObject("Scripting.FileSystemObject")
        Dim ofd As New OpenFileDialog
        ofd.FileName = ""
        ofd.InitialDirectory = "C:\"
        ofd.Filter = "csvファイル(*.csv;*.csv)|*.csv;*.csv|全てのファイル(*.*)|*.*"
        ofd.FilterIndex = 1
        ofd.Title = "開くCSVファイルを選択してください"

        If ofd.ShowDialog() = DialogResult.OK Then
            inputfile = fso.OpenTextFile(ofd.FileName)
        Else
            Exit Sub
        End If

        Dim logPath As String = ""

        SP1 = ofd.FileName.Split("\")

        For i = 0 To UBound(SP1) - 1

            If i = 0 Then
                logPath = SP1(i)
            Else
                logPath = logPath & "\" & SP1(i)
            End If

        Next i

        Dim sw As New System.IO.StreamWriter(logPath & "\MitsubishiRPALog” & DateTime.Now.ToString("yyMMddhhmm") & ".LOG", False, System.Text.Encoding.GetEncoding("shift_jis"))

        sw.WriteLine(DateTime.Now)

        sw.WriteLine("　")

        LineStr = inputfile.ReadLine




        Dim lnghWnd As Integer 'トップレベル（親）のウィンドウハンドル
        Dim lnghWndTarget As Integer 'ターゲット（子）のウィンドウハンドル

        Dim j As Long

        Dim Z As Integer

        i = 0

        Do Until inputfile.AtEndOfStream
            i = i + 1

            Label1.Text = i.ToString("00") & " 路線目"
            sw.WriteLine(i.ToString("00") & " 路線目 開始")

            LineStr = inputfile.ReadLine
            SP = LineStr.Split(",")

            '-------------------------------------
            ' ターゲットウィンドウのハンドルを取得
            '-------------------------------------
            lnghWnd = FindWindowEx(0, 0, "WindowsForms10.Window.8.app.0.2780b98_r7_ad1", "路面性状データ解析") '操作パネルのハンドル取得
            If lnghWnd = 0 Then
                sw.WriteLine("路面性状データ解析　取得失敗")
                Console.WriteLine("路面性状データ解析　取得失敗")
                Exit Do
            End If
            sw.WriteLine(lnghWnd & "：路面性状データ解析　取得")
            Console.WriteLine(lnghWnd & "：路面性状データ解析　取得")

            lnghWndTarget = FindWindowEx(lnghWnd, 0, "WindowsForms10.Window.8.app.0.2780b98_r7_ad1", "toolStripContainer1") '子ウィンドウのEdit
            If lnghWndTarget = 0 Then
                sw.WriteLine("toolStripContainer1　取得失敗")
                Console.WriteLine("toolStripContainer1　取得失敗")
                Exit Do
            End If
            sw.WriteLine(lnghWndTarget & "：toolStripContainer1　取得")
            Console.WriteLine(lnghWndTarget & "：toolStripContainer1　取得")

            lnghWndTarget = FindWindowEx(lnghWndTarget, 0, "WindowsForms10.Window.8.app.0.2780b98_r7_ad1", "") '子ウィンドウのEdit 
            If lnghWndTarget = 0 Then
                sw.WriteLine("子ウィンドウ1名無し　取得失敗")
                Console.WriteLine("子ウィンドウ1名無し　取得失敗")
                Exit Do
            End If
            sw.WriteLine(lnghWndTarget & "：子ウィンドウ1名無し　取得")
            Console.WriteLine(lnghWndTarget & "：子ウィンドウ1名無し　取得")

            Z = lnghWndTarget

            lnghWndTarget = FindWindowEx(Z, 0, "WindowsForms10.Window.8.app.0.2780b98_r7_ad1", "") '子ウィンドウのEdit  
            lnghWndTarget = FindWindowEx(Z, lnghWndTarget, "WindowsForms10.Window.8.app.0.2780b98_r7_ad1", "") '子ウィンドウのEdit    
            If lnghWndTarget = 0 Then
                sw.WriteLine("子ウィンドウ2名無し　取得失敗")
                Console.WriteLine("子ウィンドウ2名無し　取得失敗")
                Exit Do
            End If
            sw.WriteLine(lnghWndTarget & "：子ウィンドウ2名無し　取得")
            Console.WriteLine(lnghWndTarget & "：子ウィンドウ2名無し　取得")

            Z = lnghWndTarget

            lnghWndTarget = FindWindowEx(Z, 0, "WindowsForms10.Window.8.app.0.2780b98_r7_ad1", "") '子ウィンドウのEdit  
            lnghWndTarget = FindWindowEx(Z, lnghWndTarget, "WindowsForms10.Window.8.app.0.2780b98_r7_ad1", "")
            lnghWndTarget = FindWindowEx(Z, lnghWndTarget, "WindowsForms10.Window.8.app.0.2780b98_r7_ad1", "")
            If lnghWndTarget = 0 Then
                sw.WriteLine("路線一覧　取得失敗")
                Console.WriteLine("路線一覧　取得失敗")
                Exit Do
            End If
            sw.WriteLine(lnghWndTarget & "：路線一覧　取得")
            Console.WriteLine(lnghWndTarget & "：路線一覧　取得")

            Z = lnghWndTarget

            lnghWndTarget = FindWindowEx(Z, 0, "WindowsForms10.Window.8.app.0.2780b98_r7_ad1", "")
            lnghWndTarget = FindWindowEx(Z, lnghWndTarget, "WindowsForms10.Window.8.app.0.2780b98_r7_ad1", "")
            If lnghWndTarget = 0 Then
                sw.WriteLine("子ウィンドウ3名無し　取得失敗")
                Console.WriteLine("子ウィンドウ3名無し　取得失敗")
                Exit Do
            End If
            sw.WriteLine(lnghWndTarget & "：子ウィンドウ3名無し　取得")
            Console.WriteLine(lnghWndTarget & "：子ウィンドウ3名無し　取得")

            lnghWndTarget = FindWindowEx(lnghWndTarget, 0, "WindowsForms10.BUTTON.app.0.2780b98_r7_ad1", "登録")
            If lnghWndTarget = 0 Then
                sw.WriteLine("登録　取得失敗")
                Console.WriteLine("登録　取得失敗")
                Exit Do
            End If
            sw.WriteLine(lnghWndTarget & "：登録　取得")
            Console.WriteLine(lnghWndTarget & "：登録　取得")

            Call PostMessage(lnghWndTarget, WM_LBUTTONDOWN, 0, 0)
            Call PostMessage(lnghWndTarget, WM_LBUTTONUP, 0, 0)
            System.Threading.Thread.Sleep(1000)
            sw.WriteLine("登録ボタン実行済み")
            Console.WriteLine("登録ボタン実行済み")

            '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~登録ボタンを押すまで

            sw.WriteLine("　")



            lnghWnd = FindWindowEx(0, 0, "WindowsForms10.Window.8.app.0.2780b98_r7_ad1", "路線情報登録") '親のハンドル取得
            If lnghWnd = 0 Then
                sw.WriteLine("路線情報登録　取得失敗")
                Console.WriteLine("路線情報登録　取得失敗")
                Exit Do
            End If
            sw.WriteLine(lnghWnd & "：路線情報登録　取得")
            Console.WriteLine(lnghWnd & "：路線情報登録　取得")

            lnghWndTarget = FindWindowEx(lnghWnd, 0, "WindowsForms10.Window.8.app.0.2780b98_r7_ad1", "路線情報") '子ウィンドウの路線情報登録
            If lnghWndTarget = 0 Then
                sw.WriteLine("路線情報　取得失敗")
                Console.WriteLine("路線情報　取得失敗")
                Exit Do
            End If
            sw.WriteLine(lnghWndTarget & "：路線情報　取得")
            Console.WriteLine(lnghWndTarget & "：路線情報　取得")

            Z = lnghWndTarget

            lnghWndTarget = FindWindowEx(Z, 0, "WindowsForms10.COMBOBOX.app.0.2780b98_r7_ad1", "") '上下
            If lnghWndTarget = 0 Then
                sw.WriteLine("上下　取得失敗")
                Console.WriteLine("上下　取得失敗")
                Exit Do
            End If
            sw.WriteLine(lnghWndTarget & "：上下　取得")
            Console.WriteLine(lnghWndTarget & "：上下　取得")

            Call SendMessage(lnghWndTarget, CB_SELECTSTRING, -1, SP(2)) '上下はエクセルから参照する OK
            Call PostMessage(lnghWnd, WM_COMMAND, CBN_SELCHANGE, lnghWndTarget)

            lnghWndTarget = FindWindowEx(Z, 0, "WindowsForms10.EDIT.app.0.2780b98_r7_ad1", vbNullString)
            lnghWndTarget = FindWindowEx(Z, lnghWndTarget, "WindowsForms10.EDIT.app.0.2780b98_r7_ad1", vbNullString)
            lnghWndTarget = FindWindowEx(Z, lnghWndTarget, "WindowsForms10.EDIT.app.0.2780b98_r7_ad1", vbNullString)
            lnghWndTarget = FindWindowEx(Z, lnghWndTarget, "WindowsForms10.EDIT.app.0.2780b98_r7_ad1", vbNullString)
            lnghWndTarget = FindWindowEx(Z, lnghWndTarget, "WindowsForms10.EDIT.app.0.2780b98_r7_ad1", vbNullString) '路線名称
            If lnghWndTarget = 0 Then
                sw.WriteLine("路線名称　取得失敗")
                Console.WriteLine("路線名称　取得失敗")
                Exit Do
            End If
            sw.WriteLine(lnghWndTarget & "：路線名称　取得")
            Console.WriteLine(lnghWndTarget & "：路線名称　取得")

            Call SendMessage(lnghWndTarget, WM_SETTEXT, 0, SP(1)) '路線名称はエクセルから参照する OK

            lnghWndTarget = FindWindowEx(Z, lnghWndTarget, "WindowsForms10.EDIT.app.0.2780b98_r7_ad1", vbNullString) '路線管理番号
            If lnghWndTarget = 0 Then
                sw.WriteLine("路線管理番号　取得失敗")
                Console.WriteLine("路線管理番号　取得失敗")
                Exit Do
            End If
            sw.WriteLine(lnghWndTarget & "：路線管理番号　取得")
            Console.WriteLine(lnghWndTarget & "：路線管理番号　取得")

            Call SendMessage(lnghWndTarget, WM_SETTEXT, 0, SP(0)) '路線管理番号はエクセルから参照する

            sw.WriteLine("　")

            lnghWndTarget = FindWindowEx(lnghWnd, 0, "WindowsForms10.Window.8.app.0.2780b98_r7_ad1", vbNullString)
            lnghWndTarget = FindWindowEx(lnghWnd, lnghWndTarget, "WindowsForms10.Window.8.app.0.2780b98_r7_ad1", vbNullString)
            lnghWndTarget = FindWindowEx(lnghWnd, lnghWndTarget, "WindowsForms10.Window.8.app.0.2780b98_r7_ad1", vbNullString) '子ウィンドウの測定データ連携
            If lnghWndTarget = 0 Then
                sw.WriteLine("測定データ連携　取得失敗")
                Console.WriteLine("測定データ連携　取得失敗")
                Exit Do
            End If
            sw.WriteLine(lnghWndTarget & "：測定データ連携　取得")
            Console.WriteLine(lnghWndTarget & "：測定データ連携　取得")

            Z = lnghWndTarget


            lnghWndTarget = FindWindowEx(Z, 0, "WindowsForms10.EDIT.app.0.2780b98_r7_ad1", "") '走行データID
            If lnghWndTarget = 0 Then
                sw.WriteLine("走行データID　取得失敗")
                Console.WriteLine("走行データID　取得失敗")
                Exit Do
            End If
            sw.WriteLine(lnghWndTarget & "：走行データID　取得")
            Console.WriteLine(lnghWndTarget & "：走行データID　取得")

            Call SendMessage(lnghWndTarget, WM_SETTEXT, 0, SP(4)) '草稿データIDはエクセルから参照する OK
            System.Threading.Thread.Sleep(1000)

            lnghWndTarget = FindWindowEx(Z, lnghWndTarget, "WindowsForms10.COMBOBOX.app.0.2780b98_r7_ad1", "") '連携するシーン
            If lnghWndTarget = 0 Then
                sw.WriteLine("連携するシーン　取得失敗")
                Console.WriteLine("連携するシーン　取得失敗")
                Exit Do
            End If

            sw.WriteLine(lnghWndTarget & "：連携するシーン　取得")
            Console.WriteLine(lnghWndTarget & "：連携するシーン　取得")

            Call SendMessage(lnghWndTarget, CB_SELECTSTRING, 0, SP(5)) 'シーンはエクセルから参照する OK
            Call PostMessage(Z, WM_COMMAND, CBN_SELCHANGE, lnghWndTarget)

            sw.WriteLine("　")

            lnghWndTarget = FindWindowEx(lnghWnd, 0, "WindowsForms10.Window.8.app.0.2780b98_r7_ad1", vbNullString)
            lnghWndTarget = FindWindowEx(lnghWnd, lnghWndTarget, "WindowsForms10.Window.8.app.0.2780b98_r7_ad1", vbNullString) '子ウィンドウの画像形成パラメータ
            If lnghWndTarget = 0 Then
                sw.WriteLine("画像形成パラメータ　取得失敗")
                Exit Do
            End If
            sw.WriteLine(lnghWndTarget & "：画像形成パラメータ　取得")

            lnghWndTarget = FindWindowEx(lnghWndTarget, 0, "WindowsForms10.EDIT.app.0.2780b98_r7_ad1", vbNullString) '平面直角座標系番号
            If lnghWndTarget = 0 Then
                sw.WriteLine("平面直角座標系番号　取得失敗")
                Exit Do
            End If
            sw.WriteLine(lnghWndTarget & "：平面直角座標系番号　取得")
            '平面直角座標はテキストボックスから入力する。

            Call SendMessage(lnghWndTarget, WM_SETTEXT, 0, TextBox2.Text)

            sw.WriteLine("　")

            lnghWndTarget = FindWindowEx(lnghWnd, 0, "WindowsForms10.Window.8.app.0.2780b98_r7_ad1", vbNullString)
            If lnghWndTarget = 0 Then
                sw.WriteLine("子ウィンドウ1a　取得失敗")
                Exit Do
            End If
            sw.WriteLine(lnghWndTarget & "：子ウィンドウ1a　取得")

            lnghWndTarget = FindWindowEx(lnghWndTarget, 0, "WindowsForms10.BUTTON.app.0.2780b98_r7_ad1", "登録")
            If lnghWndTarget = 0 Then
                sw.WriteLine("登録　取得失敗")
                Exit Do
            End If
            sw.WriteLine(lnghWndTarget & "：登録　取得")

            Call PostMessage(lnghWndTarget, WM_LBUTTONDOWN, 0, 0)
            Call PostMessage(lnghWndTarget, WM_LBUTTONUP, 0, 0)

            Do
                lnghWnd = FindWindowEx(0, 0, "WindowsForms10.Window.8.app.0.2780b98_r7_ad1", "路線情報登録") '路線情報登録ウィンドウが消えるまで待つ
                System.Threading.Thread.Sleep(400)
            Loop Until lnghWnd = 0

            sw.WriteLine("登録完了")
            sw.WriteLine("　")

            '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~登録するまで

            Dim elmDesktop As UIAutomationClient.IUIAutomationElement
            Dim elmEdge As UIAutomationClient.IUIAutomationElement

            uiAuto = New UIAutomationClient.CUIAutomation
            elmDesktop = uiAuto.GetRootElement

            Dim cndWindowControls As UIAutomationClient.IUIAutomationCondition
            Dim aryWindowControls As UIAutomationClient.IUIAutomationElement

            cndWindowControls = uiAuto.CreatePropertyCondition(UIA_ClassNamePropertyId, "WindowsForms10.Window.8.app.0.2780b98_r7_ad1")
            aryWindowControls = elmDesktop.FindFirst(TreeScope_Subtree, cndWindowControls)

            Dim cndCoreWindow As UIAutomationClient.IUIAutomationCondition
            Dim elmCoreWindow As UIAutomationClient.IUIAutomationElement

            cndCoreWindow = uiAuto.CreatePropertyCondition(UIA_AutomationPropertyID, "toolStripContainer1")
            elmCoreWindow = aryWindowControls.FindFirst(TreeScope_Subtree, cndCoreWindow)
            If elmCoreWindow Is Nothing Then
                sw.WriteLine("toolStripContainer1　取得失敗")
                Exit Do
            End If
            sw.WriteLine(elmCoreWindow.CurrentName & " 取得")

            Dim cnd3 As UIAutomationClient.IUIAutomationCondition
            Dim elm3 As UIAutomationClient.IUIAutomationElement

            cnd3 = uiAuto.CreatePropertyCondition(UIA_ClassNamePropertyId, "WindowsForms10.Window.8.app.0.2780b98_r7_ad1")
            elm3 = elmCoreWindow.FindFirst(TreeScope_Subtree, cnd3)
            If elm3 Is Nothing Then
                sw.WriteLine("elm3無し")
                Exit Do
            End If
            sw.WriteLine(elm3.CurrentName & " 取得")

            Dim cnd4 As UIAutomationClient.IUIAutomationCondition
            Dim elm4 As UIAutomationClient.IUIAutomationElement

            cnd4 = uiAuto.CreatePropertyCondition(UIA_AutomationPropertyID, "toolStrip1")
            elm4 = elm3.FindFirst(TreeScope_Subtree, cnd4)
            If elm4 Is Nothing Then
                sw.WriteLine("elm4無し")
                Exit Do
            End If
            sw.WriteLine(elm4.CurrentName & " 取得")

            Dim cnd5 As UIAutomationClient.IUIAutomationCondition
            Dim elm5 As UIAutomationClient.IUIAutomationElement

            cnd5 = uiAuto.CreatePropertyCondition(UIA_NamePropertyID, "路面画像成形")
            elm5 = elm4.FindFirst(TreeScope_Subtree, cnd5)
            If elm5 Is Nothing Then
                sw.WriteLine("elm5無し")
                Exit Do
            End If
            sw.WriteLine(elm5.CurrentName & " 取得")

            Dim ptnInvk As UIAutomationClient.IUIAutomationInvokePattern = elm5.GetCurrentPattern(UIA_InvokePatternId)
            ptnInvk.Invoke()

            Dim cnd6 As UIAutomationClient.IUIAutomationCondition
            Dim elm6 As UIAutomationClient.IUIAutomationElement

            j = 0

            Do
                j = j + 1
                System.Threading.Thread.Sleep(3000)

                cnd6 = uiAuto.CreatePropertyCondition(UIA_ClassNamePropertyId, "WindowsForms10.Window.8.app.0.34f5582_r35_ad1")
                elm6 = elmDesktop.FindFirst(TreeScope_Subtree, cnd6)
                If elm6 Is Nothing Then
                    sw.WriteLine("路面画像形成終了")
                    Exit Do
                End If
            Loop

            Dim cnd7 As UIAutomationClient.IUIAutomationCondition
            Dim elm7 As UIAutomationClient.IUIAutomationElement

            Do
                cnd7 = uiAuto.CreatePropertyCondition(UIA_ClassNamePropertyId, "#32770")
                elm7 = elmDesktop.FindFirst(TreeScope_Subtree, cnd7)
            Loop While elm7 Is Nothing



            Dim cnd8 As UIAutomationClient.IUIAutomationCondition
            Dim elm8 As UIAutomationClient.IUIAutomationElement

            cnd8 = uiAuto.CreatePropertyCondition(UIA_ClassNamePropertyId, "Button")
            elm8 = elm7.FindFirst(TreeScope_Subtree, cnd8)
            If elm8 Is Nothing Then
                sw.WriteLine("elm8無し")
                Exit Do
            End If
            sw.WriteLine(elm8.CurrentName & " 取得")

            ptnInvk = elm8.GetCurrentPattern(UIA_InvokePatternId)
            ptnInvk.Invoke()

            sw.WriteLine(i & ”番目終了：" & DateTime.Now)
            sw.WriteLine("　")

        Loop

        sw.WriteLine("終了")
        sw.Close()
        MsgBox("終了です")

    End Sub

    Function CutRight(s, i)
        Dim iLen

        If VarType(s) <> vbString Then
            Exit Function
        End If

        iLen = Len(s)

        '// 文字列長より指定文字数が大きい場合
        If iLen < i Then
            Exit Function
        End If

        '// 指定文字数を削除して返す
        CutRight = Strings.Left(s, iLen - i)
    End Function


End Class
