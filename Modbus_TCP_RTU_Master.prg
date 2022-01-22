'**********************
' Author: hewenyuan
' Data: 2018_08_08
' Description: �bEspon�����H�W��{ Modbus TCP RTU�� Master�\��
'**********************
' New Feature and Enhancements
' Author: HaoYen.Wei (YAU-YIH ENTERPRISE CO., LTD. http://www.yauyih.com.tw/)
' 2021_01_05: ���s�s�g�γ̨ΤƵ{���y�{�A���{����@Ū�g�]�Ƥ覡�A��@�\��X 1, 3, 4, 15, 16
'**********************
'Modbus�\��X
'�\��X	�\��W��		�y�z
'1		Ū���u�骬�A	Ū�iŪ��Bit
'2		Ū����J���A	Ū�uŪBit
'3		Ū���O���H�s��	Ū�iŪ�gWord '
'4		Ū����J�H�s��	Ū�uŪWord
'5		�j���u��		�g�J�iŪ�gBit
'6		�w�m��H�s��	�g�J�iŪ�gWord
'15		�j��h�u��		��q�g�J�h�ӳs�򪺥iŪ�gBit
'16		�w�m�h�H�s��	��q�g�J�h�ӳs�򪺥iŪ�gWord
'
'TCP Header	Address	Function Code	Start register addr	data
'6bytes		1byte	1byte			2byte				N bytes
'
'00 00 00 00 00 06 01 01 00 20 00 20
'
'00 00 �q�T���ѧO�X �H�����X�C�������ɧǤΥi�a�ʫ�ĳ���s�s�g�AMODBUS�榡���e�X��ƫ�A����ۦP�q�T�ѧO�X��ܥѸөR�O���T��
'00 00 ��ĳ���ѽX 0���MODBUS TCP
'00 06 ����(Header���᪺��ƪ���)
'01    ����
'01    �\��X
'00 20 �_�l��m
'00 20 ����
'///////////////////////////////////////////////////////////////////////////
'�аO�D���O�_�w�s�u
Global Boolean ModbusTcpMaster_Connected(10)
'�ưȤ��ѧO�Ÿ�
UShort Flags(10)

'///////////////////////////////////////////////////////////////////////////
'*****************
' Description: �Ұ�Modbus TCP�D��
' Param handle: �ާ@������������s��(0 ~ 10)
'*****************
Function ModbusTcpMaster_Start(handle As UShort)
    ModbusTcpMaster_Connected(handle) = False    '�����]�L�s�u
    Call ModbusTcpMaster_CheckConnection(handle) '�ˬd�s�u���p,�Y�B���_�}���A�h���խ��s�إ߳s�u
    Do
        Wait 2
        Call ModbusTcpMaster_CheckConnection(handle)
    Loop
Fend

'///////////////////////////////////////////////////////////////////////////
'*****************
' Description: �ˬd�s�u���p,�Y�B���_�}���A�h���խ��s�إ߳s�u
' Param handle: �ާ@������������s��(0 ~ 10)
'*****************
Function ModbusTcpMaster_CheckConnection(handle As UShort)
    UShort port
    port = 200 + handle
    Integer chkNetResult
    chkNetResult = ChkNet(port)
 
    If chkNetResult < 0 Then
    '�s�u�q��
        OpenNet #port As Client
        WaitNet #port, 10000
        Print "moudbus tcp handle:", handle, " open (port ", port, ")"
        ModbusTcpMaster_Connected(handle) = True '�s�u���\
    EndIf
Fend

'///////////////////////////////////////////////////////////////////////////
'*****************
' Description: Ū���q����Word 0x03 �_�l�d�� 40000�}�l
' Param handle: �ާ@������������s��
' Param inputIndex: �_�lWord�s��
' Param num: Ū��Word���ƶq 512�H��
' Param readData: Ū���쪺���G��J��readData�}�C���AreadData�}�C���������j��(num x 2) + 10
' Return: ��^�O�_Ū�����\
'*****************
Function ModbusTcpMaster_ReadWord(handle As UShort, inputIndex As UShort, num As UShort, ByRef readWordData() As Integer) As Boolean

    OnErr GoTo TcpErrHandle '���~�o�͸���
 
    '�ˬd�O�_�w�s�u
    If ModbusTcpMaster_Connected(handle) = False Then
        ModbusTcpMaster_ReadWord = False
        UShort portNum
        portNum = 200 + handle
        Print "prot:", portNum, " not open."
        Exit Function
    EndIf
 
    UByte sendData(11)
    '*******************************
    '�ͦ�MBAP �����Y�y�z
    '*******************************
    '�ưȤ��ѧO�Ÿ�
    sendData(0) = Int(Rnd(&HFF))
    sendData(1) = Int(Rnd(&HFF))
    '��ĳ�ѧO�Ÿ�
    sendData(2) = &H0
    sendData(3) = &H0
    '����
    sendData(4) = &H0
    sendData(5) = 6
    '����
    sendData(6) = &H1
    '*******************************
    '�ͦ��d�߫��O
    '*******************************
    '�\��X
    sendData(7) = &H03;
    '��J�f�s��
    sendData(8) = RShift(&HFF00 And inputIndex, 8)
    sendData(9) = &H00FF And inputIndex
    '����
    sendData(10) = RShift(&HFF00 And num, 8)
    sendData(11) = &H00FF And num
 
    '*******************************
    '�ǰe���O
    '*******************************
    UShort port
    port = 200 + handle
    WriteBin #port, sendData(), 12
 
    '*******************************
    '������^���
    '*******************************
    '���ݦ���Ʊ�����
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
 
    If numOfChars < 1 Then '�W��
        ModbusTcpMaster_ReadWord = False
        Print "prot:", port, " modbus funcion 04 error. Timeout"
        Exit Function '���X�禡
    EndIf
 
    UByte reciveData(1034)
    ReadBin #port, reciveData(), numOfChars
 
    '*******************************
    '�P�_�ѧO�Ÿ��O�_���T
    '*******************************
    '�ٲ�

    '*******************************
    '�ѪR���
    '���G��Ƥ� ��8��--�\��X ��9��--��ƪ��� ��10��}�l�O���
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
 
    TcpErrHandle: '�s�����~�����
    Print "Error Number# ", Err
    Print "port:", port, "ModbusTcpMaster_ReadWord read write error"
    CloseNet #port
    ModbusTcpMaster_Connected(handle) = False '�s�u����
Fend

'///////////////////////////////////////////////////////////////////////////
'*****************
' Description: Ū���q����Word 0x03 �_�l�d�� 40000�}�l
' Param handle: �ާ@������������s��
' Param inputIndex: �_�lWord�s��
' Param num: Ū��Double Word���ƶq 256�H�U
' Param readData: Ū���쪺���G��J��readData�}�C���AreadData�}�C���������j��(num x 2) + 10
' Return: ��^�O�_Ū�����\
'*****************
Function ModbusTcpMaster_ReadDoubleWord(handle As UShort, inputIndex As UShort, num As UShort, ByRef readDoubleWordData() As Int32) As Boolean

    OnErr GoTo TcpErrHandle '���~�o�͸���
 
    '�ˬd�O�_�w�s�u
    If ModbusTcpMaster_Connected(handle) = False Then
        ModbusTcpMaster_ReadDoubleWord = False
        UShort portNum
        portNum = 200 + handle
        Print "prot:", portNum, " not open."
        Exit Function
    EndIf
 
    UByte sendData(11)
    '*******************************
    '�ͦ�MBAP �����Y�y�z
    '*******************************
    '�ưȤ��ѧO�Ÿ�
    sendData(0) = Int(Rnd(&HFF))
    sendData(1) = Int(Rnd(&HFF))
    '��ĳ�ѧO�Ÿ�
    sendData(2) = &H0
    sendData(3) = &H0
    '����
    sendData(4) = &H0
    sendData(5) = 6
    '����
    sendData(6) = &H1
    '*******************************
    '�ͦ��d�߫��O
    '*******************************
    '�\��X
    sendData(7) = &H03;
    '��J�f�s��
    sendData(8) = RShift(&HFF00 And inputIndex, 8)
    sendData(9) = &H00FF And inputIndex
    '����
    UShort DataLength
    DataLength = num * 2
    sendData(10) = RShift(&HFF00 And DataLength, 8)
    sendData(11) = &H00FF And DataLength
 
    '*******************************
    '�ǰe���O
    '*******************************
    UShort port
    port = 200 + handle
    WriteBin #port, sendData(), 12
 
    '*******************************
    '������^���
    '*******************************
    '���ݦ���Ʊ�����
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
 
    If numOfChars < 1 Then '�W��
        ModbusTcpMaster_ReadDoubleWord = False
        Print "prot:", port, " modbus funcion 04 error. Timeout"
        Exit Function
    EndIf
 
    UByte reciveData(1034)
    ReadBin #port, reciveData(), numOfChars
 
    '*******************************
    '�P�_�ѧO�Ÿ��O�_���T
    '*******************************
    '�ٲ�

    '*******************************
    '�ѪR���
    '���G��Ƥ� ��8��--�\��X ��9��--��ƪ��� ��10��}�l�O���
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
 
    TcpErrHandle: '�s�����~�����
    Print "Error Number# ", Err
    Print "port:", port, "ModbusTcpMaster_ReadDoubleWord read error"
    CloseNet #port
    ModbusTcpMaster_Connected(handle) = False '�s�u����
Fend

'///////////////////////////////////////////////////////////////////////////
'*****************
' Description: Ū���q����Input�q 0x02
' Param handle: �ާ@������������s��
' Param inputIndex: �_�l��J�f�s��
' Param num: Ū����J�f���ƶq
' Param readData: Ū���쪺���G��J��readData�}�C���AreadData�}�C���������j��num
' Return: ��^�O�_Ū�����\
'*****************
Function ModbusTcpMaster_ReadInput(handle As UShort, inputIndex As UShort, num As UShort, ByRef readData() As Boolean) As Boolean

    OnErr GoTo TcpErrHandle '���~�o�͸���
 
    '�ˬd�O�_�w�s�u
    If ModbusTcpMaster_Connected(handle) = False Then
        ModbusTcpMaster_ReadInput = False
    EndIf
 
    UByte sendData(12)
    UByte reciveData(1024)
    '*******************************
    '�ͦ�MBAP �����Y�y�z
    '*******************************
    '�ưȤ��ѧO�Ÿ�
    sendData(0) = Int(Rnd(&HFF))
    sendData(1) = Int(Rnd(&HFF))
    '��ĳ�ѧO�Ÿ�
    sendData(2) = &H0
    sendData(3) = &H0
    '����
    sendData(4) = &H0
    sendData(5) = 6
    '����
    sendData(6) = &H1
    '*******************************
    '�ͦ��d�߫��O
    '*******************************
    '�\��X
    sendData(7) = &H02;
    '��J�f�s��
    sendData(8) = RShift(&HFF00 And inputIndex, 8)
    sendData(9) = &H00FF And inputIndex
    '����
    sendData(10) = RShift(&HFF00 And num, 8)
    sendData(11) = &H00FF And num
 
    '*******************************
    '�ǰe���O
    '*******************************
    UShort port
    port = 200 + handle
    WriteBin #port, sendData(), 12
 
    '*******************************
    '������^���
    '*******************************
    '���ݦ���Ʊ�����
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

    If numOfChars < 1 Then '�����W��
        ModbusTcpMaster_ReadInput = False
        Print "prot:", port, " modbus funcion 02 error. Timeout"
        Exit Function
    EndIf
 
    ReadBin #port, reciveData(), numOfChars

    '*******************************
    '�ѪR���
    '���G��Ƥ� ��8��--�\��X ��9��--��ƪ��� ��10��}�l�O���
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
    Exit Function '�j��פ�禡
 
    TcpErrHandle: '�s�����~�����
    Print "Error Number# ", Err
    Print "port:", port, "ModbusTcpMaster_ReadInput read write error"
    CloseNet #port
    ModbusTcpMaster_Connected(handle) = False '�s�u����
Fend

'///////////////////////////////////////////////////////////////////////////
'*****************
' Description: Ū���q����Coil�q 0x01
' Param handle: �ާ@������������s��
' Param inputIndex: �_�lCoil�s��
' Param num: Ū��Coil���ƶq
' Param readData: Ū���쪺���G��J��readData�}�C���AreadData�}�C���������j��num
' Return: ��^�O�_Ū�����\
'*****************
Function ModbusTcpMaster_ReadCoil(handle As UShort, inputIndex As UShort, num As UShort, ByRef readData() As Boolean) As Boolean

    OnErr GoTo TcpErrHandle '���~�o�͸���
 
    '�ˬd�O�_�w�s�u
    If ModbusTcpMaster_Connected(handle) = False Then
        ModbusTcpMaster_ReadCoil = False
    EndIf
 
    UByte sendData(12)
    UByte reciveData(1024)
 
    '*******************************
    '�ͦ�MBAP �����Y�y�z
    '*******************************
    '�ưȤ��ѧO�Ÿ�
    sendData(0) = &H0
    sendData(1) = &H0
    '��ĳ�ѧO�Ÿ�
    sendData(2) = &H0
    sendData(3) = &H0
    '����
    sendData(4) = &H0
    sendData(5) = 6
    '����
    sendData(6) = &H1
    '*******************************
    '�ͦ��d�߫��O
    '*******************************
    '�\��X
    sendData(7) = &H01;
    '��J�f�s��
    sendData(8) = RShift(&HFF00 And inputIndex, 8)
    sendData(9) = &H00FF And inputIndex
    '����
    sendData(10) = RShift(&HFF00 And num, 8)
    sendData(11) = &H00FF And num
 
    '*******************************
    '�ǰe���O
    '*******************************
    UShort port
    port = 200 + handle
    WriteBin #port, sendData(), 12
 
    '*******************************
    '������^���
    '*******************************
    '���ݦ���Ʊ�����
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
 
    If numOfChars < 1 Then '�W��
        ModbusTcpMaster_ReadCoil = False
        Print "prot:", port, " modbus funcion 01 error. Timeout"
        Exit Function
    EndIf
 
    ReadBin #port, reciveData(), numOfChars

    '*******************************
    '�ѪR���
    '���G��Ƥ� ��8��--�\��X ��9��--��ƪ��� ��10��}�l�O���
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
 
    TcpErrHandle: '�s�����~�����
    Print "Error Number# ", Err
    Print "port:", port, "ModbusTcpMaster_ReadCoil read write error"
    CloseNet #port
    ModbusTcpMaster_Connected(handle) = False '�s�u����
Fend

'///////////////////////////////////////////////////////////////////////////
'*****************
' Description: ��q�g�J�q����Coil�u���� 0x0F
' Param handle: �ާ@������������s��
' Param startIndex: �_�lCoil�s��
' Param num: �g�J��XCoil���ƶq
' Param coilsData: �g�J���
' Return: ��^�O�_Ū�����\
' ��ƴV: MBAP(7byte) + �\��X(1byte) + �_�l�a�}(2byte) + Coil�ƶq(2byte) +��ƪ���(1byte) + ���(num/8)
'*****************
Function ModbusTcpMaster_WriteMultiCoil(handle As UShort, startIndex As UShort, num As UShort, ByRef coilsData() As Boolean) As Boolean
 
    OnErr GoTo TcpErrHandle '���~�o�͸���
 
    '�ˬd�O�_�w�s�u
    If ModbusTcpMaster_Connected(handle) = False Then
        ModbusTcpMaster_WriteMultiCoil = False
    EndIf
 
    UByte sendData(256)
    UByte reciveData(1024)
    Integer numOfChars, length1, length2, length3
 
    '�p��v�e��ƪ�����
    length1 = Int((num + 7) / 8)
    length2 = 7 + length1
    length3 = 13 + length1
 
    '*******************************
    '�ͦ�MBAP �����Y�y�z
    '*******************************
    '�ưȤ��ѧO�Ÿ�
    sendData(0) = Int(Rnd(&HFF))
    sendData(1) = Int(Rnd(&HFF))
    '��ĳ�ѧO�Ÿ�
    sendData(2) = &H0
    sendData(3) = &H0
    '����
    sendData(4) = RShift(&HFF00 And length2, 8)
    sendData(5) = &H00FF And length2
    '����
    sendData(6) = &H1
    '*******************************
    '�ͦ��d�߫��O
    '*******************************
    '�\��X
    sendData(7) = &H0F;
    '�_�l�a�}
    sendData(8) = RShift(&HFF00 And startIndex, 8)
    sendData(9) = &H00FF And startIndex
    '����
    sendData(10) = RShift(&HFF00 And num, 8)
    sendData(11) = &H00FF And num
    '��ƪ���
    sendData(12) = &H00FF And length1
    'coil���
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
    '�ǰe���O
    '*******************************
    UShort port
    port = 200 + handle
    WriteBin #port, sendData(), length3
 
    '*******************************
    '������^���
    '*******************************
    '���ݦ���Ʊ�����
    UShort waitCount
    waitCount = 0
    Do
        numOfChars = ChkNet(port)
        If waitCount > 10 Then '�O�ɳ]�w 0.1�� 
            Exit Do
        EndIf
 
        If numOfChars > 0 Then '������
            Wait 0.02
            Exit Do
        Else
            Wait 0.01
        EndIf
 	
        waitCount = waitCount + 1
    Loop
 
    If numOfChars < 1 Then '�ˬd��ơA�O�_�W�ɤ���F�賣�S����
        ModbusTcpMaster_WriteMultiCoil = False
        Print "prot:", port, " modbus funcion 15 error. Timeout"
        Exit Function
    EndIf
 
    ReadBin #port, reciveData(), numOfChars

    '*******************************
    '�ѪR���
    '���G��Ƥ� ��8��--�\��X
    '*******************************
    If reciveData(7) = &H0F Then
        ModbusTcpMaster_WriteMultiCoil = True
    Else
        ModbusTcpMaster_WriteMultiCoil = False
    EndIf
    Exit Function
 
    TcpErrHandle: '�s�����~�����
    Print "Error Number# ", Err
    Print "port:", port, "ModbusTcpMaster_WriteMultiCoil read write error"
    CloseNet #port
    ModbusTcpMaster_Connected(handle) = False '�s�u����

Fend

'///////////////////////////////////////////////////////////////////////////
'*****************
' Description: ��q�g�J�q����Word��� 0x10
' Param handle: �ާ@������������s��
' Param startIndex: �_�lWord�s��
' Param num: �g�JWord���ƶq 256�H�U
' Param WordData: �g�J���
' Return: ��^�O�_Ū�����\
' ��ƴV: MBAP(7byte) + �\��X(1byte) + �_�l�a�}(2byte) + word�ƶq(2byte) +��ƪ���(1byte) + ���(num * 2)
'*****************
Function ModbusTcpMaster_WriteMultiWord(handle As UShort, startIndex As UShort, num As UShort, ByRef WordData() As Int32) As Boolean
 
    OnErr GoTo TcpErrHandle '���~�o�͸���
 
    '�ˬd�O�_�w�s�u
    If ModbusTcpMaster_Connected(handle) = False Then
        ModbusTcpMaster_WriteMultiWord = False
        UShort portNum
        portNum = 200 + handle
        Print "prot:", portNum, " not open."
        Exit Function
    EndIf
 
    UByte sendData(384)
    UShort length1, length2, length3
 
    '�p��o�e��ƪ�����
    length1 = num * 2
    length2 = 7 + length1
    length3 = 13 + length1
 
    '*******************************
    '�ͦ�MBAP �����Y�y�z
    '*******************************
    '�ưȤ��ѧO�Ÿ�
    sendData(0) = Int(Rnd(&HFF))
    sendData(1) = Int(Rnd(&HFF))
    '��ĳ�ѧO�Ÿ�
    sendData(2) = &H0
    sendData(3) = &H0
    '����
    sendData(4) = RShift(&HFF00 And length2, 8)
    sendData(5) = &H00FF And length2
    '����
    sendData(6) = &H01
    '*******************************
    '�ͦ��d�߫��O
    '*******************************
    '�\��X
    sendData(7) = &H10;
    '�_�l�a�}
    sendData(8) = RShift(&HFF00 And startIndex, 8)
    sendData(9) = &H00FF And startIndex
    '����
    sendData(10) = RShift(&HFF00 And num, 8)
    sendData(11) = &H00FF And num
    '��ƪ���
    sendData(12) = &H00FF And length1
    'Word���
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
    '�ǰe���O
    '*******************************
    UShort port
    port = 200 + handle
    WriteBin #port, sendData(), length3
 
    '*******************************
    '������^���
    '*******************************
    '���ݦ���Ʊ�����
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
 
    If numOfChars < 1 Then '�ˬd��ơA�O�_�W�ɤ���F�賣�S����
        ModbusTcpMaster_WriteMultiWord = False
        Print "prot:", port, " modbus funcion 16 error. Timeout"
        Exit Function
    EndIf
 
    UByte reciveData(16)
    ReadBin #port, reciveData(), numOfChars
    '*******************************
    '�P�_�ѧO�Ÿ��O�_���T
    '*******************************
    '�ٲ�
    
    '*******************************
    '�ѪR���
    '���G��Ƥ� ��8��--�\��X
    '*******************************
    If reciveData(7) = &H10 Then
        ModbusTcpMaster_WriteMultiWord = True
    Else
        ModbusTcpMaster_WriteMultiWord = False
    EndIf
    Exit Function
 
    TcpErrHandle: '�s�����~�����
    Print "Error Number# ", Err
    Print "port:", port, "ModbusTcpMaster_WriteMultiWord read write error"
    CloseNet #port
    ModbusTcpMaster_Connected(handle) = False '�s�u����
Fend

'///////////////////////////////////////////////////////////////////////////
'*****************
' Description: ��q�g�J�q����Double Word��� 0x10
' Param handle: �ާ@������������s��
' Param startIndex: �_�l��}�s��
' Param num: �g�JDouble Word���ƶq 256�H�U
' Param WordData: �g�J���
' Return: ��^�O�_Ū�����\
' ��ƴV: MBAP(7byte) + �\��X(1byte) + �_�l�a�}(2byte) + word�ƶq(2byte) +��ƪ���(1byte) + ���(num * 2)
'*****************
Function ModbusTcpMaster_WriteMultiDoubleWord(handle As UShort, startIndex As UShort, num As UShort, ByRef DoubleWordData() As Int64) As Boolean
 
    OnErr GoTo TcpErrHandle '���~�o�͸���
 
    '�ˬd�O�_�w�s�u
    If ModbusTcpMaster_Connected(handle) = False Then
        ModbusTcpMaster_WriteMultiDoubleWord = False
        UShort portNum
        portNum = 200 + handle
        Print "prot:", portNum, " not open."
        Exit Function
    EndIf
 
    UByte sendData(384)
    UShort length1, length2, length3
 
    '�p��o�e��ƪ�����
    length1 = num * 4
    length2 = 7 + length1
    length3 = 13 + length1
 
    '*******************************
    '�ͦ�MBAP �����Y�y�z
    '*******************************
    '�ưȤ��ѧO�Ÿ�
    sendData(0) = Int(Rnd(&HFF))
    sendData(1) = Int(Rnd(&HFF))
    '��ĳ�ѧO�Ÿ�
    sendData(2) = &H0
    sendData(3) = &H0
    '����
    sendData(4) = RShift(&HFF00 And length2, 8)
    sendData(5) = &H00FF And length2
    '����
    sendData(6) = &H01
    '*******************************
    '�ͦ��d�߫��O
    '*******************************
    '�\��X
    sendData(7) = &H10;
    '�_�l�a�}
    sendData(8) = RShift(&HFF00 And startIndex, 8)
    sendData(9) = &H00FF And startIndex
    '����
    UShort DataLength
    DataLength = num * 2
    sendData(10) = RShift(&HFF00 And DataLength, 8)
    sendData(11) = &H00FF And DataLength
    '��ƪ���
    sendData(12) = &H00FF And length1
    'Word���
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
    '�ǰe���O
    '*******************************
    UShort port
    port = 200 + handle
    WriteBin #port, sendData(), length3
 
    '*******************************
    '������^���
    '*******************************
    '���ݦ���Ʊ�����
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
 
    If numOfChars < 1 Then '�ˬd��ơA�O�_�W�ɤ���F�賣�S����
        ModbusTcpMaster_WriteMultiDoubleWord = False
        Print "prot:", port, " modbus funcion 16 error. Timeout"
        Exit Function
    EndIf
 
    UByte reciveData(16)
    ReadBin #port, reciveData(), numOfChars
    '*******************************
    '�P�_�ѧO�Ÿ��O�_���T
    '*******************************
    '�ٲ�
    '*******************************
    '�ѪR���
    '���G��Ƥ� ��8��--�\��X
    '*******************************
    If reciveData(7) = &H10 Then
        ModbusTcpMaster_WriteMultiDoubleWord = True
    Else
        ModbusTcpMaster_WriteMultiDoubleWord = False
    EndIf
    Exit Function
 
    TcpErrHandle: '�s�����~�����
    Print "Error Number# ", Err
    Print "port:", port, "ModbusTcpMaster_WriteMultiDoubleWord read write error"
    CloseNet #port
    ModbusTcpMaster_Connected(handle) = False '�s�u����
Fend


