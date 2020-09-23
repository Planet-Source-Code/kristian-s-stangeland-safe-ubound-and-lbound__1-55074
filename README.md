<div align="center">

## Safe UBound and LBound


</div>

### Description

Ever wanted to use LBound and UBound to get arrays boundaries without jumping over error message when the array is empty? These functions will replace the ordinary LBound and UBound procedures so you don&#8217;t need to worry about errors. I've also included a way to get the dimensions of an array. Just paste the following code into a module, and the problem is solved.
 
### More Info
 
SafeUBound and SafeLBound: [Address to the array], [What dimension you want to obtain]

ArrayDims: [Address to the array]

You obtain the address to an array by passing it to the VarPtrArray API call. So if you want to get the boundaries of an array called aTmp, you need to call the functions like this:

lLowBound = SafeLBound(VarPtrArray(aTmp))

lHighBound = SafeUBound(VarPtrArray(aTmp))

lDimensions = ArrayDims(VarPtrArray(aTmp))

When dealing with string arrays that isn't allocated at design time, you *must* add the value 4 to the lpArray-paramenter:

lLowBound = SafeLBound(VarPtrArray(aString) + 4)

As expected from the ordinary functions, except that they will return -1 when the array is empty.

Since the return value is minus when the array is empty it's a big chance you will get problems with minus dimensioned arrays, but who use them anyway?


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Kristian S\. Stangeland](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/kristian-s-stangeland.md)
**Level**          |Advanced
**User Rating**    |5.0 (15 globes from 3 users)
**Compatibility**  |VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[VB function enhancement](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/vb-function-enhancement__1-25.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/kristian-s-stangeland-safe-ubound-and-lbound__1-55074/archive/master.zip)

### API Declarations

```
Declare Function VarPtrArray Lib "msvbvm50.dll" Alias "VarPtr" (Ptr() As Any) As Long
Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
```


### Source Code

```
Option Explicit
Declare Function VarPtrArray Lib "msvbvm50.dll" Alias "VarPtr" (Ptr() As Any) As Long
Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Public Function SafeUBound(ByVal lpArray As Long, Optional Dimension As Long = 1) As Long
  Dim lAddress&, cElements&, lLbound&, cDims%
  If Dimension < 1 Then
    SafeUBound = -1
    Exit Function
  End If
  CopyMemory lAddress, ByVal lpArray, 4
  If lAddress = 0 Then
    ' The array isn't initilized
    SafeUBound = -1
    Exit Function
  End If
  ' Calculate the dimensions
  CopyMemory cDims, ByVal lAddress, 2
  Dimension = cDims - Dimension + 1
  ' Obtain the needed data
  CopyMemory cElements, ByVal (lAddress + 16 + ((Dimension - 1) * 8)), 4
  CopyMemory lLbound, ByVal (lAddress + 20 + ((Dimension - 1) * 8)), 4
  SafeUBound = cElements + lLbound - 1
End Function
Public Function SafeLBound(ByVal lpArray As Long, Optional Dimension As Long = 1) As Long
  Dim lAddress&, cElements&, lLbound&, cDims%
  If Dimension < 1 Then
    SafeLBound = -1
    Exit Function
  End If
  CopyMemory lAddress, ByVal lpArray, 4
  If lAddress = 0 Then
    ' The array isn't initilized
    SafeLBound = -1
    Exit Function
  End If
  ' Calculate the dimensions
  CopyMemory cDims, ByVal lAddress, 2
  Dimension = cDims - Dimension + 1
  ' Obtain the needed data
  CopyMemory lLbound, ByVal (lAddress + 20 + ((Dimension - 1) * 8)), 4
  SafeLBound = lLbound
End Function
Public Function ArrayDims(ByVal lpArray As Long) As Integer
  Dim lAddress As Long
  CopyMemory lAddress, ByVal lpArray, 4
  If lAddress = 0 Then
    ' The array isn't initilized
    ArrayDims = -1
    Exit Function
  End If
  CopyMemory ArrayDims, ByVal lAddress, 2
End Function
```

