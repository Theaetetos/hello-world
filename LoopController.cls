VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LoopController"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private param_coll As New Collection 'collection of arrays of values to use in model iterations
Private name_coll As New Collection 'collection of names of params stored in param_coll
Private next_ind As Variant 'array of indices of next values to be extracted using NextParameters

Sub AddParameter(Arr As Variant, ParamName As String)

'adds 1d base 1 array of values to param_coll and name to name_coll
'1d or 2d array may be passed in; dimension of greater magnitude of 2d array is read

'-----------------------------------------

Dim dimens As Integer
dimens = are_valid_dims(Arr, 1, 2, "AddParameter Method")
If dimens = 0 Then Exit Sub

param_coll.Add to_1d(Arr, dimens)
name_coll.Add ParamName

Dim i As Integer
ReDim next_ind(1 To param_coll.Count)
For i = 1 To UBound(next_ind)
    next_ind(i) = 1
Next i

End Sub

Function NextParameters() As Variant

'returns a table (as array; base 1) of param names as headers together with the values to use for next model run
'returns a one-element 2d array of "done" when all parameters have been looped through

'---------------------------------------------------------------

Dim ret As Variant

If next_ind(1) = -1 Then
    ReDim ret(1 To 1, 1 To 1)
    ret(1, 1) = "done"
Else
    ReDim ret(1 To 2, 1 To param_coll.Count)
    Dim i As Integer
    For i = 1 To UBound(ret, 2)
        ret(1, i) = name_coll(i)
        ret(2, i) = param_coll(i)(next_ind(i))
    Next i
End If

NextParameters = ret

inc_ind

End Function

Private Sub inc_ind()

'increments requisite value in next_ind for next unique parameter combination; first element becomes -1 if all have been returned
'algorithm:
'   1) starting from bottom, find first array of param values whose value in next_ind is less than its highest value
'       if none, set next_ind(1) = 0 to indicate all params have been looped through
'   2) increment value in next_ind corresponding to that array
'   3) set value in next_ind corresponding to all param arrays below to 1

'----------------------------------------------------------------------------

Dim cur_arr As Variant
Dim cur_param_ind As Integer, cur_ubound As Integer
cur_param_ind = param_coll.Count

Do While True
    If next_ind(cur_param_ind) < UBound(param_coll(cur_param_ind)) Then
        next_ind(cur_param_ind) = mod_by_1(CInt(next_ind(cur_param_ind)), 1, UBound(param_coll(cur_param_ind)))
        For cur_param_ind = cur_param_ind + 1 To param_coll.Count
            next_ind(cur_param_ind) = 1
        Next cur_param_ind
        Exit Sub
    End If
    cur_param_ind = cur_param_ind - 1
    If cur_param_ind = 0 Then
        next_ind(1) = -1
        Exit Sub
    End If
Loop

End Sub

Private Function are_valid_dims(ByRef Arr As Variant, l_bound As Integer, u_bound As Integer, caller As String) As Integer

'checks whether passed-in array has dimensions within given bounds (inclusive)
'returns 0 if not, number of dimensions if so
'meant to be used for error handling (see message to user in case returning 0)

'------------------------------------------------------------------------------------------

'param checks:
If l_bound < 0 Then
    MsgBox "are_valid_dims procedure encountered an error:" & vbNewLine & _
           "l_bound parameter is less than 0."
    Exit Function
End If

If l_bound > u_bound Then
    MsgBox "are_valid_dims procedure encountered an error:" & vbNewLine & _
           "l_bound parameter is greater than ubound parameter."
    Exit Function
End If

Dim dimens As Integer
dimens = get_arr_dims(Arr)

If dimens >= l_bound And dimens <= u_bound Then
    are_valid_dims = dimens
Else
    MsgBox caller & " procedure encountered an error:" & vbNewLine & _
           "Dimensions of array argument not within bounds."
End If

End Function

Private Function get_arr_dims(ByRef Arr As Variant) As Integer

'returns the number of dimensions of passed-in array

'----------------------------------------

Dim i As Long, dimens As Integer
Dim done As Boolean
dimens = 1

Do Until done = True
    On Error GoTo Dims:
    i = UBound(Arr, dimens)
    On Error GoTo 0
    If done = True Then Exit Do
    dimens = dimens + 1
Loop

Exit Function

Dims:
done = True
dimens = dimens - 1
get_arr_dims = dimens
Resume Next

End Function

Private Function to_1d(ByRef Arr As Variant, dimens As Integer) As Variant

'returns Arr as 1d base 1 array
'Arr may be 1d or 2d; dest is 1d
'Arr may be base 0 or 1; dest is base 1
'reads by whichever dimension is larger (by rows if equal)

'---------------------------------------------------------------

Dim dest_arr As Variant
Dim i As Integer, corr As Integer

If dimens = 1 Then
    If LBound(Arr) = 0 Then corr = -1 'map indices to each other in case arrays have different bases
    ReDim dest_arr(1 To UBound(Arr) - LBound(Arr) + 1)
    For i = LBound(dest_arr) To UBound(dest_arr)
        dest_arr(i) = Arr(i + corr)
    Next i
ElseIf dimens = 2 Then
    Dim rd_dim As Integer
    If LBound(Arr, 1) = 0 Then corr = -1 'map indices to each other in case arrays have different bases
    'determine which dimension to read by
    Select Case UBound(Arr, 1)
        Case Is > UBound(Arr, 2)
            rd_dim = 2
        Case Is <= UBound(Arr, 2)
            rd_dim = 1
    End Select
    'read array
    Select Case rd_dim
        Case Is = 2
            ReDim dest_arr(1 To UBound(Arr, 1) - LBound(Arr, 1) + 1)
            For i = LBound(dest_arr) To UBound(dest_arr)
                dest_arr(i) = Arr(i + corr, LBound(Arr, 2))
            Next i
        Case Is = 1
            ReDim dest_arr(1 To UBound(Arr, 2) - LBound(Arr, 2) + 1)
            For i = LBound(dest_arr) To UBound(dest_arr)
                dest_arr(i) = Arr(LBound(Arr, 1), i + corr)
            Next i
    End Select
End If

to_1d = dest_arr

End Function

Private Function mod_by_1(val As Integer, l_bound As Integer, u_bound As Integer) As Integer

'modular addition by 1: adds 1 to val, then returns that number or l_bound if it is greater than u_bound
'no checking is done on relationship between val, l_bound, and u_bound

'----------------------------------------------------------------------

val = val + 1
If val > u_bound Then val = l_bound
mod_by_1 = val

End Function
