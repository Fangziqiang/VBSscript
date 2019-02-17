'VBS的常用函数
'1、类型转换函数
'CBool(expression)
'expression是零，则返回false；否则返回True
Dim A,B,check
A=5:B=5
check = CBool(A=B)
'True
MsgBox check
A = 0
check = CBool(A)
'False
MsgBox check

'CByte()转换数据类型为整数
Dim MyDouble,MyByte
MyDouble = 125.5678
MyByte=CByte(MyDouble)
MsgBox MyByte

'Chr()
Dim MyA,MyB
MyA = Chr(65)
MyB = Chr(66)
MsgBox MyA & MyB 

