Attribute VB_Name = "ModuleMoningPrm"
' 모닝프레임 주문 자동화 - 엑셀 실행 버튼용 VBA 매크로

Public Sub RunMoningPrm()
    Dim projectPath As String
    Dim batPath As String
    Dim resultPath As String
    Dim wsh As Object
    Dim fso As Object
    Dim ts As Object
    Dim msg As String

    ' ★ 프로젝트 폴더 경로 (본인 PC 경로에 맞게 수정)
    projectPath = "C:\Users\82103\Desktop\moning_prm"
    batPath = projectPath & "\run.bat"
    resultPath = projectPath & "\output\최신_처리결과.txt"

    If Dir(batPath) = "" Then
        MsgBox "run.bat을 찾을 수 없습니다:" & vbCrLf & batPath, vbCritical, "모닝프레임 자동화"
        Exit Sub
    End If

    Set wsh = CreateObject("WScript.Shell")
    wsh.CurrentDirectory = projectPath

    ' run.bat 실행 (창 표시, 완료까지 대기)
    wsh.Run """" & batPath & """", 1, True

    ' 처리 결과 파일 읽어서 안내
    msg = "처리가 끝났습니다." & vbCrLf & vbCrLf
    msg = msg & "생산지시서, 월결산 엑셀이 자동으로 열렸을 것입니다." & vbCrLf
    msg = msg & "안 열렸다면 아래 폴더를 확인하세요:" & vbCrLf
    msg = msg & projectPath & "\output" & vbCrLf & vbCrLf

    Set fso = CreateObject("Scripting.FileSystemObject")
    If fso.FileExists(resultPath) Then
        Set ts = fso.OpenTextFile(resultPath, 1, False, -1)
        msg = msg & ts.ReadAll
        ts.Close
    End If

    MsgBox msg, vbInformation, "모닝프레임 자동화"

    ' output 폴더 탐색기 열기
    If fso.FolderExists(projectPath & "\output") Then
        wsh.Run "explorer """ & projectPath & "\output""", 1, False
    End If
End Sub
