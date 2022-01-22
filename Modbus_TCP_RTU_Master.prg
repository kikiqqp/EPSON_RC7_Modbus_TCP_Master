'**********************
' Author: hewenyuan
' Data: 2018_08_08
' Description: 在Espon機器人上實現 Modbus TCP RTU的 Master功能
'**********************
' New Feature and Enhancements
' Author: HaoYen.Wei (YAU-YIH ENTERPRISE CO., LTD. http://www.yauyih.com.tw/)
' 2021_01_05: 重新編寫及最佳化程式流程，本程式實作讀寫設備方式，實作功能碼 1, 3, 4, 15, 16
'**********************
'Modbus功能碼
'功能碼	功能名稱		描述
'1		讀取線圈狀態	讀可讀性Bit
'2		讀取輸入狀態	讀只讀Bit
'3		讀取保持寄存器	讀可讀寫Word '
'4		讀取輸入寄存器	讀只讀Word
'5		強制單線圈		寫入可讀寫Bit
'6		預置單寄存器	寫入可讀寫Word
'15		強制多線圈		批量寫入多個連續的可讀寫Bit
'16		預置多寄存器	批量寫入多個連續的可讀寫Word
'
'TCP Header	Address	Function Code	Start register addr	data
'6bytes		1byte	1byte			2byte				N bytes
'
'00 00 00 00 00 06 01 01 00 20 00 20
'
'00 00 通訊的識別碼 隨機號碼。基於網路時序及可靠性建議重新編寫，MODBUS格式中送出資料後，收到相同通訊識別碼表示由該命令所響應
'00 00 協議辨識碼 0表示MODBUS TCP
'00 06 長度(Header之後的資料長度)
'01    站號
'01    功能碼
'00 20 起始位置
'00 20 長度
'///////////////////////////////////////////////////////////////////////////
'標記主站是否已連線
Global Boolean ModbusTcpMaster_Connected(10)
'事務元識別符號
UShort Flags(10)

'///////////////////////////////////////////////////////////////////////////
'*****************
' Description: 啟動Modbus TCP主站
' Param handle: 操作的的網路控制編號(0 ~ 10)
'*****************
Function ModbusTcpMaster_Start(handle As UShort)
    ModbusTcpMaster_Connected(handle) = False    '先假設無連線
    Call ModbusTcpMaster_CheckConnection(handle) '檢查連線狀況,若處於斷開狀態則嘗試重新建立連線
    Do
        Wait 2
        Call ModbusTcpMaster_CheckConnection(handle)
    Loop
Fend

'///////////////////////////////////////////////////////////////////////////
'*****************
' Description: 檢查連線狀況,若處於斷開狀態則嘗試重新建立連線
' Param handle: 操作的的網路控制編號(0 ~ 10)
'*****************
Function ModbusTcpMaster_CheckConnection(handle As UShort)
    UShort port
    port = 200 + handle
    Integer chkNetResult
    chkNetResult = ChkNet(port)
 
    If chkNetResult < 0 Then
    '連線從機
        OpenNet #port As Client
        WaitNet #port, 10000
        Print "moudbus tcp handle:", handle, " open (port ", port, ")"
        ModbusTcpMaster_Connected(handle) = True '連線成功
    EndIf
Fend

'///////////////////////////////////////////////////////////////////////////
'*****************
' Description: 讀取從站的Word 0x03 起始範圍 40000開始
' Param handle: 操作的的網路控制編號
' Param inputIndex: 起始Word編號
' Param num: 讀取Word的數量 512以內
' Param readData: 讀取到的結果放入到readData陣列中，readData陣列的長度應大於(num x 2) + 10
' Return: 返回是否讀取成功
'*****************
Function ModbusTcpMaster_ReadWord(handle As UShort, inputIndex As UShort, num As UShort, ByRef readWordData() As Integer) As Boolean

    OnErr GoTo TcpErrHandle '錯誤發生跳轉
 
    '檢查是否已連線
    If ModbusTcpMaster_Connected(handle) = False Then
        ModbusTcpMaster_ReadWord = False
        UShort portNum
        portNum = 200 + handle
        Print "prot:", portNum, " not open."
        Exit Function
    EndIf
 
    UByte sendData(11)
    '*******************************
    '生成MBAP 報文頭描述
    '*******************************
    '事務元識別符號
    sendData(0) = Int(Rnd(&HFF))
    sendData(1) = Int(Rnd(&HFF))
    '協議識別符號
    sendData(2) = &H0
    sendData(3) = &H0
    '長度
    sendData(4) = &H0
    sendData(5) = 6
    '站號
    sendData(6) = &H1
    '*******************************
    '生成查詢指令
    '*******************************
    '功能碼
    sendData(7) = &H03;
    '輸入口編號
    sendData(8) = RShift(&HFF00 And inputIndex, 8)
    sendData(9) = &H00FF And inputIndex
    '長度
    sendData(10) = RShift(&HFF00 And num, 8)
    sendData(11) = &H00FF And num
 
    '*******************************
    '傳送指令
    '*******************************
    UShort port
    port = 200 + handle
    WriteBin #port, sendData(), 12
 
    '*******************************
    '接收返回資料
    '*******************************
    '等待有資料接收到
    UShort waitCount
    Integer numOfChars
    waitCount = 0
    Do Until (waitCount > 50)
        numOfChars = ChkNet(port)
 
        If numOfChars >= ((num * 2) + 9) Then
            Exit Do
        Else
            Wait 0.01
        EndIf
 
        waitCount = waitCount + 1
    Loop
 
    If numOfChars < 1 Then '超時
        ModbusTcpMaster_ReadWord = False
        Print "prot:", port, " modbus funcion 04 error. Timeout"
        Exit Function '跳出函式
    EndIf
 
    UByte reciveData(1034)
    ReadBin #port, reciveData(), numOfChars
 
    '*******************************
    '判斷識別符號是否正確
    '*******************************
    '省略

    '*******************************
    '解析資料
    '結果資料中 第8位--功能碼 第9位--資料長度 第10位開始是資料
    '*******************************
    If reciveData(7) = &H03 Then
        Integer i, j
        For i = 0 To num - 1
            If BTst(reciveData(9 + i * 2), 7) Then
                j = LShift((reciveData(9 + i * 2) Xor &H00FF), 8) + &H00FF
                j = ((reciveData(9 + i * 2 + 1) Xor &H00FF) Or &HFF00) And j
                j = (j + 1) * -1
            Else
                j = LShift(reciveData(9 + i * 2), 8) + &H00FF
                j = (reciveData(9 + i * 2 + 1) Or &HFF00) And j
            EndIf
            readWordData(i) = j
        Next
    EndIf
    Exit Function
 
    TcpErrHandle: '連接錯誤時顯示
    Print "Error Number# ", Err
    Print "port:", port, "ModbusTcpMaster_ReadWord read write error"
    CloseNet #port
    ModbusTcpMaster_Connected(handle) = False '連線失敗
Fend

'///////////////////////////////////////////////////////////////////////////
'*****************
' Description: 讀取從站的Word 0x03 起始範圍 40000開始
' Param handle: 操作的的網路控制編號
' Param inputIndex: 起始Word編號
' Param num: 讀取Double Word的數量 256以下
' Param readData: 讀取到的結果放入到readData陣列中，readData陣列的長度應大於(num x 2) + 10
' Return: 返回是否讀取成功
'*****************
Function ModbusTcpMaster_ReadDoubleWord(handle As UShort, inputIndex As UShort, num As UShort, ByRef readDoubleWordData() As Int32) As Boolean

    OnErr GoTo TcpErrHandle '錯誤發生跳轉
 
    '檢查是否已連線
    If ModbusTcpMaster_Connected(handle) = False Then
        ModbusTcpMaster_ReadDoubleWord = False
        UShort portNum
        portNum = 200 + handle
        Print "prot:", portNum, " not open."
        Exit Function
    EndIf
 
    UByte sendData(11)
    '*******************************
    '生成MBAP 報文頭描述
    '*******************************
    '事務元識別符號
    sendData(0) = Int(Rnd(&HFF))
    sendData(1) = Int(Rnd(&HFF))
    '協議識別符號
    sendData(2) = &H0
    sendData(3) = &H0
    '長度
    sendData(4) = &H0
    sendData(5) = 6
    '站號
    sendData(6) = &H1
    '*******************************
    '生成查詢指令
    '*******************************
    '功能碼
    sendData(7) = &H03;
    '輸入口編號
    sendData(8) = RShift(&HFF00 And inputIndex, 8)
    sendData(9) = &H00FF And inputIndex
    '長度
    UShort DataLength
    DataLength = num * 2
    sendData(10) = RShift(&HFF00 And DataLength, 8)
    sendData(11) = &H00FF And DataLength
 
    '*******************************
    '傳送指令
    '*******************************
    UShort port
    port = 200 + handle
    WriteBin #port, sendData(), 12
 
    '*******************************
    '接收返回資料
    '*******************************
    '等待有資料接收到
    UShort waitCount
    Integer numOfChars
    waitCount = 0
    Do Until (waitCount > 50)
        numOfChars = ChkNet(port)
 
        If numOfChars >= ((num * 4) + 9) Then
            Exit Do
        Else
            Wait 0.01
        EndIf

        waitCount = waitCount + 1
    Loop
 
    If numOfChars < 1 Then '超時
        ModbusTcpMaster_ReadDoubleWord = False
        Print "prot:", port, " modbus funcion 04 error. Timeout"
        Exit Function
    EndIf
 
    UByte reciveData(1034)
    ReadBin #port, reciveData(), numOfChars
 
    '*******************************
    '判斷識別符號是否正確
    '*******************************
    '省略

    '*******************************
    '解析資料
    '結果資料中 第8位--功能碼 第9位--資料長度 第10位開始是資料
    '*******************************
    If reciveData(7) = &H03 Then
        Integer i
        Int32 j
        For i = 0 To num - 1
            If BTst(reciveData(9 + i * 4 + 3), 7) Then
                j = reciveData(9 + i * 4 + 1) Xor &HFF
                j = LShift((reciveData(9 + i * 4) Xor &HFF), 8) + j
                j = LShift((reciveData(9 + i * 4 + 3) Xor &HFF), 16) + j
                j = LShift((reciveData(9 + i * 4 + 2) Xor &HFF), 24) + j
                j = (j + 1) * -1
            Else
                j = reciveData(9 + i * 4 + 1)
                j = LShift(reciveData(9 + i * 4), 8) + j
                j = LShift(reciveData(9 + i * 4 + 3), 16) + j
                j = LShift(reciveData(9 + i * 4 + 2), 24) + j
            EndIf
            readDoubleWordData(i) = j
        Next
    EndIf
    Exit Function
 
    TcpErrHandle: '連接錯誤時顯示
    Print "Error Number# ", Err
    Print "port:", port, "ModbusTcpMaster_ReadDoubleWord read error"
    CloseNet #port
    ModbusTcpMaster_Connected(handle) = False '連線失敗
Fend

'///////////////////////////////////////////////////////////////////////////
'*****************
' Description: 讀取從站的Input量 0x02
' Param handle: 操作的的網路控制編號
' Param inputIndex: 起始輸入口編號
' Param num: 讀取輸入口的數量
' Param readData: 讀取到的結果放入到readData陣列中，readData陣列的長度應大於num
' Return: 返回是否讀取成功
'*****************
Function ModbusTcpMaster_ReadInput(handle As UShort, inputIndex As UShort, num As UShort, ByRef readData() As Boolean) As Boolean

    OnErr GoTo TcpErrHandle '錯誤發生跳轉
 
    '檢查是否已連線
    If ModbusTcpMaster_Connected(handle) = False Then
        ModbusTcpMaster_ReadInput = False
    EndIf
 
    UByte sendData(12)
    UByte reciveData(1024)
    '*******************************
    '生成MBAP 報文頭描述
    '*******************************
    '事務元識別符號
    sendData(0) = Int(Rnd(&HFF))
    sendData(1) = Int(Rnd(&HFF))
    '協議識別符號
    sendData(2) = &H0
    sendData(3) = &H0
    '長度
    sendData(4) = &H0
    sendData(5) = 6
    '站號
    sendData(6) = &H1
    '*******************************
    '生成查詢指令
    '*******************************
    '功能碼
    sendData(7) = &H02;
    '輸入口編號
    sendData(8) = RShift(&HFF00 And inputIndex, 8)
    sendData(9) = &H00FF And inputIndex
    '長度
    sendData(10) = RShift(&HFF00 And num, 8)
    sendData(11) = &H00FF And num
 
    '*******************************
    '傳送指令
    '*******************************
    UShort port
    port = 200 + handle
    WriteBin #port, sendData(), 12
 
    '*******************************
    '接收返回資料
    '*******************************
    '等待有資料接收到
    UShort waitCount
    Integer numOfChars
    waitCount = 0
    Do
        numOfChars = ChkNet(port)
 	
        If waitCount > 10 Then
            Exit Do
        EndIf

        If numOfChars > 0 Then
            Wait 0.02
            Exit Do
        Else
            Wait 0.01
        EndIf
        waitCount = waitCount + 1
    Loop

    If numOfChars < 1 Then '接收超時
        ModbusTcpMaster_ReadInput = False
        Print "prot:", port, " modbus funcion 02 error. Timeout"
        Exit Function
    EndIf
 
    ReadBin #port, reciveData(), numOfChars

    '*******************************
    '解析資料
    '結果資料中 第8位--功能碼 第9位--資料長度 第10位開始是資料
    '*******************************
    If reciveData(7) = &H02 Then
        UShort byteIndex, bitIndex
        Integer i
        For i = 0 To num - 1
            byteIndex = Int(i / 8)
            bitIndex = i Mod 8
            readData(i) = BTst(reciveData(9 + byteIndex), bitIndex)
        Next
    EndIf
    Exit Function '強制終止函式
 
    TcpErrHandle: '連接錯誤時顯示
    Print "Error Number# ", Err
    Print "port:", port, "ModbusTcpMaster_ReadInput read write error"
    CloseNet #port
    ModbusTcpMaster_Connected(handle) = False '連線失敗
Fend

'///////////////////////////////////////////////////////////////////////////
'*****************
' Description: 讀取從站的Coil量 0x01
' Param handle: 操作的的網路控制編號
' Param inputIndex: 起始Coil編號
' Param num: 讀取Coil的數量
' Param readData: 讀取到的結果放入到readData陣列中，readData陣列的長度應大於num
' Return: 返回是否讀取成功
'*****************
Function ModbusTcpMaster_ReadCoil(handle As UShort, inputIndex As UShort, num As UShort, ByRef readData() As Boolean) As Boolean

    OnErr GoTo TcpErrHandle '錯誤發生跳轉
 
    '檢查是否已連線
    If ModbusTcpMaster_Connected(handle) = False Then
        ModbusTcpMaster_ReadCoil = False
    EndIf
 
    UByte sendData(12)
    UByte reciveData(1024)
 
    '*******************************
    '生成MBAP 報文頭描述
    '*******************************
    '事務元識別符號
    sendData(0) = &H0
    sendData(1) = &H0
    '協議識別符號
    sendData(2) = &H0
    sendData(3) = &H0
    '長度
    sendData(4) = &H0
    sendData(5) = 6
    '站號
    sendData(6) = &H1
    '*******************************
    '生成查詢指令
    '*******************************
    '功能碼
    sendData(7) = &H01;
    '輸入口編號
    sendData(8) = RShift(&HFF00 And inputIndex, 8)
    sendData(9) = &H00FF And inputIndex
    '長度
    sendData(10) = RShift(&HFF00 And num, 8)
    sendData(11) = &H00FF And num
 
    '*******************************
    '傳送指令
    '*******************************
    UShort port
    port = 200 + handle
    WriteBin #port, sendData(), 12
 
    '*******************************
    '接收返回資料
    '*******************************
    '等待有資料接收到
    Integer numOfChars
    UShort waitCount
    waitCount = 0
    Do
        numOfChars = ChkNet(port)
 	
        If waitCount > 10 Then
            Exit Do
        EndIf
 
        If numOfChars > 0 Then
            Wait 0.02
            Exit Do
        Else
            Wait 0.01
        EndIf
        
        waitCount = waitCount + 1
    Loop
 
    If numOfChars < 1 Then '超時
        ModbusTcpMaster_ReadCoil = False
        Print "prot:", port, " modbus funcion 01 error. Timeout"
        Exit Function
    EndIf
 
    ReadBin #port, reciveData(), numOfChars

    '*******************************
    '解析資料
    '結果資料中 第8位--功能碼 第9位--資料長度 第10位開始是資料
    '*******************************
    If reciveData(7) = &H01 Then
        UShort byteIndex, bitIndex
        Integer i
        For i = 0 To num - 1
            byteIndex = Int(i / 8)
            bitIndex = i Mod 8
            readData(i) = BTst(reciveData(9 + byteIndex), bitIndex)
        Next
    EndIf
    Exit Function
 
    TcpErrHandle: '連接錯誤時顯示
    Print "Error Number# ", Err
    Print "port:", port, "ModbusTcpMaster_ReadCoil read write error"
    CloseNet #port
    ModbusTcpMaster_Connected(handle) = False '連線失敗
Fend

'///////////////////////////////////////////////////////////////////////////
'*****************
' Description: 批量寫入從站的Coil線圈資料 0x0F
' Param handle: 操作的的網路控制編號
' Param startIndex: 起始Coil編號
' Param num: 寫入輸出Coil的數量
' Param coilsData: 寫入資料
' Return: 返回是否讀取成功
' 資料幀: MBAP(7byte) + 功能碼(1byte) + 起始地址(2byte) + Coil數量(2byte) +資料長度(1byte) + 資料(num/8)
'*****************
Function ModbusTcpMaster_WriteMultiCoil(handle As UShort, startIndex As UShort, num As UShort, ByRef coilsData() As Boolean) As Boolean
 
    OnErr GoTo TcpErrHandle '錯誤發生跳轉
 
    '檢查是否已連線
    If ModbusTcpMaster_Connected(handle) = False Then
        ModbusTcpMaster_WriteMultiCoil = False
    EndIf
 
    UByte sendData(256)
    UByte reciveData(1024)
    Integer numOfChars, length1, length2, length3
 
    '計算髮送資料的長度
    length1 = Int((num + 7) / 8)
    length2 = 7 + length1
    length3 = 13 + length1
 
    '*******************************
    '生成MBAP 報文頭描述
    '*******************************
    '事務元識別符號
    sendData(0) = Int(Rnd(&HFF))
    sendData(1) = Int(Rnd(&HFF))
    '協議識別符號
    sendData(2) = &H0
    sendData(3) = &H0
    '長度
    sendData(4) = RShift(&HFF00 And length2, 8)
    sendData(5) = &H00FF And length2
    '站號
    sendData(6) = &H1
    '*******************************
    '生成查詢指令
    '*******************************
    '功能碼
    sendData(7) = &H0F;
    '起始地址
    sendData(8) = RShift(&HFF00 And startIndex, 8)
    sendData(9) = &H00FF And startIndex
    '長度
    sendData(10) = RShift(&HFF00 And num, 8)
    sendData(11) = &H00FF And num
    '資料長度
    sendData(12) = &H00FF And length1
    'coil資料
    UShort byteIndex, bitIndex
    Integer i
    For i = 0 To num - 1
        byteIndex = Int(i / 8)
        bitIndex = i Mod 8

        If coilsData(i) = False Then
            sendData(13 + byteIndex) = BClr(sendData(13 + byteIndex), bitIndex)
        Else
            sendData(13 + byteIndex) = BSet(sendData(13 + byteIndex), bitIndex)
        EndIf
    Next
 
    '*******************************
    '傳送指令
    '*******************************
    UShort port
    port = 200 + handle
    WriteBin #port, sendData(), length3
 
    '*******************************
    '接收返回資料
    '*******************************
    '等待有資料接收到
    UShort waitCount
    waitCount = 0
    Do
        numOfChars = ChkNet(port)
        If waitCount > 10 Then '逾時設定 0.1秒 
            Exit Do
        EndIf
 
        If numOfChars > 0 Then '有收到
            Wait 0.02
            Exit Do
        Else
            Wait 0.01
        EndIf
 	
        waitCount = waitCount + 1
    Loop
 
    If numOfChars < 1 Then '檢查資料，是否超時什麼東西都沒收到
        ModbusTcpMaster_WriteMultiCoil = False
        Print "prot:", port, " modbus funcion 15 error. Timeout"
        Exit Function
    EndIf
 
    ReadBin #port, reciveData(), numOfChars

    '*******************************
    '解析資料
    '結果資料中 第8位--功能碼
    '*******************************
    If reciveData(7) = &H0F Then
        ModbusTcpMaster_WriteMultiCoil = True
    Else
        ModbusTcpMaster_WriteMultiCoil = False
    EndIf
    Exit Function
 
    TcpErrHandle: '連接錯誤時顯示
    Print "Error Number# ", Err
    Print "port:", port, "ModbusTcpMaster_WriteMultiCoil read write error"
    CloseNet #port
    ModbusTcpMaster_Connected(handle) = False '連線失敗

Fend

'///////////////////////////////////////////////////////////////////////////
'*****************
' Description: 批量寫入從站的Word資料 0x10
' Param handle: 操作的的網路控制編號
' Param startIndex: 起始Word編號
' Param num: 寫入Word的數量 256以下
' Param WordData: 寫入資料
' Return: 返回是否讀取成功
' 資料幀: MBAP(7byte) + 功能碼(1byte) + 起始地址(2byte) + word數量(2byte) +資料長度(1byte) + 資料(num * 2)
'*****************
Function ModbusTcpMaster_WriteMultiWord(handle As UShort, startIndex As UShort, num As UShort, ByRef WordData() As Int32) As Boolean
 
    OnErr GoTo TcpErrHandle '錯誤發生跳轉
 
    '檢查是否已連線
    If ModbusTcpMaster_Connected(handle) = False Then
        ModbusTcpMaster_WriteMultiWord = False
        UShort portNum
        portNum = 200 + handle
        Print "prot:", portNum, " not open."
        Exit Function
    EndIf
 
    UByte sendData(384)
    UShort length1, length2, length3
 
    '計算發送資料的長度
    length1 = num * 2
    length2 = 7 + length1
    length3 = 13 + length1
 
    '*******************************
    '生成MBAP 報文頭描述
    '*******************************
    '事務元識別符號
    sendData(0) = Int(Rnd(&HFF))
    sendData(1) = Int(Rnd(&HFF))
    '協議識別符號
    sendData(2) = &H0
    sendData(3) = &H0
    '長度
    sendData(4) = RShift(&HFF00 And length2, 8)
    sendData(5) = &H00FF And length2
    '站號
    sendData(6) = &H01
    '*******************************
    '生成查詢指令
    '*******************************
    '功能碼
    sendData(7) = &H10;
    '起始地址
    sendData(8) = RShift(&HFF00 And startIndex, 8)
    sendData(9) = &H00FF And startIndex
    '長度
    sendData(10) = RShift(&HFF00 And num, 8)
    sendData(11) = &H00FF And num
    '資料長度
    sendData(12) = &H00FF And length1
    'Word資料
    UShort HighData, LowData
    Integer i
    For i = 0 To num - 1
        If WordData(i) Then
            LowData = WordData(i) And &H00FF
            HighData = RShift(&HFF00 And WordData(i), 8)
        Else
            HighData = &H00
            LowData = &H00
        EndIf
        sendData(13 + i * 2) = HighData
        sendData(13 + i * 2 + 1) = LowData
    Next
 
    '*******************************
    '傳送指令
    '*******************************
    UShort port
    port = 200 + handle
    WriteBin #port, sendData(), length3
 
    '*******************************
    '接收返回資料
    '*******************************
    '等待有資料接收到
    UShort waitCount
    UShort numOfChars

    waitCount = 0
    Do Until (waitCount > length3)
        numOfChars = ChkNet(port)
 
        If numOfChars >= 12 Then
            Exit Do
        Else
            Wait 0.01
        EndIf
 
        waitCount = waitCount + 1
    Loop
 
    If numOfChars < 1 Then '檢查資料，是否超時什麼東西都沒收到
        ModbusTcpMaster_WriteMultiWord = False
        Print "prot:", port, " modbus funcion 16 error. Timeout"
        Exit Function
    EndIf
 
    UByte reciveData(16)
    ReadBin #port, reciveData(), numOfChars
    '*******************************
    '判斷識別符號是否正確
    '*******************************
    '省略
    
    '*******************************
    '解析資料
    '結果資料中 第8位--功能碼
    '*******************************
    If reciveData(7) = &H10 Then
        ModbusTcpMaster_WriteMultiWord = True
    Else
        ModbusTcpMaster_WriteMultiWord = False
    EndIf
    Exit Function
 
    TcpErrHandle: '連接錯誤時顯示
    Print "Error Number# ", Err
    Print "port:", port, "ModbusTcpMaster_WriteMultiWord read write error"
    CloseNet #port
    ModbusTcpMaster_Connected(handle) = False '連線失敗
Fend

'///////////////////////////////////////////////////////////////////////////
'*****************
' Description: 批量寫入從站的Double Word資料 0x10
' Param handle: 操作的的網路控制編號
' Param startIndex: 起始位址編號
' Param num: 寫入Double Word的數量 256以下
' Param WordData: 寫入資料
' Return: 返回是否讀取成功
' 資料幀: MBAP(7byte) + 功能碼(1byte) + 起始地址(2byte) + word數量(2byte) +資料長度(1byte) + 資料(num * 2)
'*****************
Function ModbusTcpMaster_WriteMultiDoubleWord(handle As UShort, startIndex As UShort, num As UShort, ByRef DoubleWordData() As Int64) As Boolean
 
    OnErr GoTo TcpErrHandle '錯誤發生跳轉
 
    '檢查是否已連線
    If ModbusTcpMaster_Connected(handle) = False Then
        ModbusTcpMaster_WriteMultiDoubleWord = False
        UShort portNum
        portNum = 200 + handle
        Print "prot:", portNum, " not open."
        Exit Function
    EndIf
 
    UByte sendData(384)
    UShort length1, length2, length3
 
    '計算發送資料的長度
    length1 = num * 4
    length2 = 7 + length1
    length3 = 13 + length1
 
    '*******************************
    '生成MBAP 報文頭描述
    '*******************************
    '事務元識別符號
    sendData(0) = Int(Rnd(&HFF))
    sendData(1) = Int(Rnd(&HFF))
    '協議識別符號
    sendData(2) = &H0
    sendData(3) = &H0
    '長度
    sendData(4) = RShift(&HFF00 And length2, 8)
    sendData(5) = &H00FF And length2
    '站號
    sendData(6) = &H01
    '*******************************
    '生成查詢指令
    '*******************************
    '功能碼
    sendData(7) = &H10;
    '起始地址
    sendData(8) = RShift(&HFF00 And startIndex, 8)
    sendData(9) = &H00FF And startIndex
    '長度
    UShort DataLength
    DataLength = num * 2
    sendData(10) = RShift(&HFF00 And DataLength, 8)
    sendData(11) = &H00FF And DataLength
    '資料長度
    sendData(12) = &H00FF And length1
    'Word資料
    UShort HHighData, HLowData, LHighData, LLowData
    Integer i
    For i = 0 To num - 1
        If DoubleWordData(i) Then
            LLowData = DoubleWordData(i) And &H000000FF
            LHighData = RShift(&H0000FF00 And DoubleWordData(i), 8)
            HLowData = RShift(&H00FF0000 And DoubleWordData(i), 16)
            HHighData = (RShift(DoubleWordData(i), 24)) And &H00FF
        Else
            HHighData = &H00
            HLowData = &H00
            LHighData = &H00
            LLowData = &H00
        EndIf
        sendData(13 + i * 4) = LHighData
        sendData(13 + i * 4 + 1) = LLowData
        sendData(13 + i * 4 + 2) = HHighData
        sendData(13 + i * 4 + 3) = HLowData
    Next
 
    '*******************************
    '傳送指令
    '*******************************
    UShort port
    port = 200 + handle
    WriteBin #port, sendData(), length3
 
    '*******************************
    '接收返回資料
    '*******************************
    '等待有資料接收到
    UShort waitCount, numOfChars
    waitCount = 0
    Do Until (waitCount > length3)
        numOfChars = ChkNet(port)
 
        If numOfChars >= 12 Then
            Exit Do
        Else
            Wait 0.01
        EndIf
 
        waitCount = waitCount + 1
    Loop
 
    If numOfChars < 1 Then '檢查資料，是否超時什麼東西都沒收到
        ModbusTcpMaster_WriteMultiDoubleWord = False
        Print "prot:", port, " modbus funcion 16 error. Timeout"
        Exit Function
    EndIf
 
    UByte reciveData(16)
    ReadBin #port, reciveData(), numOfChars
    '*******************************
    '判斷識別符號是否正確
    '*******************************
    '省略
    '*******************************
    '解析資料
    '結果資料中 第8位--功能碼
    '*******************************
    If reciveData(7) = &H10 Then
        ModbusTcpMaster_WriteMultiDoubleWord = True
    Else
        ModbusTcpMaster_WriteMultiDoubleWord = False
    EndIf
    Exit Function
 
    TcpErrHandle: '連接錯誤時顯示
    Print "Error Number# ", Err
    Print "port:", port, "ModbusTcpMaster_WriteMultiDoubleWord read write error"
    CloseNet #port
    ModbusTcpMaster_Connected(handle) = False '連線失敗
Fend


