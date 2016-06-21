'Poorboy Hash Function	e-mail code snippet
'Author: David M. Lewis
'Category: Security
'Very simple method of producing hashes.  This has variable length input and output.
Public Function Poorboy(indata As String, hextype As 
Boolean) As String
'(c) David M. Lewis, 2013, Free for private, non-
commerical/govt use.  email:stmdk@hotmail.com
'Poorboy is a simple variable length hash.  Output is the 
same length as input.
'Min input 8 bytes. Max has been purposely set to 8192 bytes 
(65kbits).  It is very slow when you go that high!
'Be careful of overflow when dealing with hex output.
'HOW IT WORKS:
'Each byte value is added to the next, plus the round count, in 
a cummulative pattern.  The values are adjusted via mod.
'The process is performed for a number of rounds equal to 
the sum value of all original bytes, max 16K rounds.
'This was designed to be a password/key expander, nothing 
more.
'Slowness is meant as an anti-brute force measure.
'Collision prevention and security is not guaranteed or 
implied.  NOT peer reviewed at the time of its creation.

Dim ilen As Integer 'Length of input
Dim count As Integer 'Counter for input
Dim rounds As Integer ' Counter for rounds
Dim isum As Integer 'Sum val or all bytes.
Dim val1 As Long 'First byte
Dim val2 As Long 'Second byte
Dim val3 As Long 'New byte val
Dim hash As String 'Hex code of hash

If Len(indata) < 8 Then indata = indata & String(8 - 
Len(indata), Chr(0)) ' if too short, pad with 0's
If Len(indata) > 8192 Then indata = Left(indata, 8192) 
'Truncate input if too long.  This can be changed.

ilen = Len(indata)

'If hextype = 1 And ilen > 16000 Then Exit Function 'add 
checks here if you want.
'If hextype = 0 And ilen > 32000 Then Exit Function


For count = 1 To ilen 'Calculate the sum value of all bytes.
 val1 = Asc(Mid(indata, count, 1))
 isum = (isum + val1) Mod 16384
Next count


For rounds = 1 To isum
 For count = 1 To ilen
  val1 = Asc(Mid(indata, count, 1))
  If count < ilen Then val2 = Asc(Mid(indata, count + 1, 1))
  If count = ilen Then val2 = Asc(Mid(indata, 1, 1)) 'Wrap 
around if at the last byte.
  val3 = (val1 + val2 + rounds) Mod 256
  Mid(indata, count) = Chr(val3) 'save for string mode.
   If rounds = ilen Then 'if at last round, save hex now for 
hexmode.
    hash = hash & Hex(val3)
   End If
 Next count
Next rounds

If hextype = False Then Poorboy = indata 'altered of course.
If hextype = True Then Poorboy = hash

End Function
