Global Integer EOUT0
Global Integer EOUT1
Global Integer EIN0
Global Integer EIN1
Function main
	Reset
	Xqt ModbusIoMap
Fend

Function ModbusIoMap()
	
	SetNet #203, "127.0.0.1", 502, CRLF, NONE, 0
	
	Boolean readData(1024)
	Boolean writeData(1024)
	Integer readWordData(1024)
	Int32 writeWordData(1024)
	Int64 writeDoubleWordData(1024)
	Int32 readDoubleWordData(1024)

	Do
		writeData(0) = EOUT0
		writeData(1) = EOUT1

		ModbusTcpMaster_CheckConnection(3) '¶}±Ò203 Port
		If ModbusTcpMaster_Connected(3) = True Then
			'ModbusTcpMaster_ReadCoil 3, 0 + 32, 32, ByRef readData() 
			'ModbusTcpMaster_ReadWord 3, 0, 3, ByRef readWordData()
			'ModbusTcpMaster_ReadDoubleWord 3, 0, 3, ByRef readDoubleWordData()
			'ModbusTcpMaster_WriteMultiWord 3, 0, 3, ByRef writeWordData()
			'ModbusTcpMaster_WriteMultiDoubleWord 3, 0, 3, ByRef writeDoubleWordData()
			'ModbusTcpMaster_WriteMultiCoil 3, 0 + 0, 32, ByRef writeData()
		EndIf
		
		EIN0 = readData(0)
		EIN1 = readData(1)
		
		Wait 0.02
	Loop
Fend
