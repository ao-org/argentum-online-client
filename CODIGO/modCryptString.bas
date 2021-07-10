Attribute VB_Name = "modCryptString"
'   ¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤
'   ¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤
'   ¤¤¤¤¤¤¤¤¤¤¤¤¤¤                                                 ¤¤¤¤¤¤¤¤¤¤¤¤¤¤
'   ¤¤¤¤¤¤¤¤¤¤¤¤¤¤  All Functions and Subroutines are the Complete ¤¤¤¤¤¤¤¤¤¤¤¤¤¤
'   ¤¤¤¤¤¤¤¤¤¤¤¤¤¤    and Expressed Property of Joseph Sullivan.   ¤¤¤¤¤¤¤¤¤¤¤¤¤¤
'   ¤¤¤¤¤¤¤¤¤¤¤¤¤¤                                                 ¤¤¤¤¤¤¤¤¤¤¤¤¤¤
'   ¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤
'   ¤¤¤¤¤¤¤¤¤¤¤¤¤¤                                                 ¤¤¤¤¤¤¤¤¤¤¤¤¤¤
'   ¤¤¤¤¤¤¤¤¤¤¤¤¤¤  If you have any questions or comments, please  ¤¤¤¤¤¤¤¤¤¤¤¤¤¤
'   ¤¤¤¤¤¤¤¤¤¤¤¤¤¤     contact Mr. Sullivan at bhJoeS@aol.com.     ¤¤¤¤¤¤¤¤¤¤¤¤¤¤
'   ¤¤¤¤¤¤¤¤¤¤¤¤¤¤                                                 ¤¤¤¤¤¤¤¤¤¤¤¤¤¤
'   ¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤
'   ¤¤¤¤¤¤¤¤¤¤¤¤¤¤                                                 ¤¤¤¤¤¤¤¤¤¤¤¤¤¤
'   ¤¤¤¤¤¤¤¤¤¤¤¤¤¤        Visual Basic 5.0 Generalized Code        ¤¤¤¤¤¤¤¤¤¤¤¤¤¤
'   ¤¤¤¤¤¤¤¤¤¤¤¤¤¤                                                 ¤¤¤¤¤¤¤¤¤¤¤¤¤¤
'   ¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤
'   ¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤

'   Module Name
'   ¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤
'   modSecurity

'   Last Updated
'   ¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤
'   Tuesday, August 01, 2000

'   Dependants
'   ¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤
'   [None]

'   Private Dimensions
'   ¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤
'   [None]

'   Private Constants
'   ¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤
'   [None]

'   Public Subroutines
'   ¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤
'   [None]

'   Private Subroutines
'   ¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤
'   [None]

'   Public Functions
'   ¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤
'   Encrypt
'   Decrypt

'   Private Functions
'   ¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤
'   [None]

Option Explicit

Public Function EncryptString(ByVal StringToEncrypt As String) As String

Remarks:
        '   The following function takes the parameter 'StringToEncrypt' and performs
        '   multiple mathematical transformations on it.  Every step has been
        '   documented through remarks to cut down on confusion of the process
        '   itself.  Upon any error, the error is ignored and execution of the
        '   function continues.  It is suggested that you do not attempt to encrypt
        '   more than 5k to 10k at once because the function is so memory intensive.
        '   For instance, on a 200 Mhz, with 128 MB RAM and Win98 SE, an uncompiled
        '   version of this function averaged the following times (over a period of
        '   ten trials):
        '
        '               1000 characters (1K)    -   3333 characters per second
        '               3000 characters (3K)    -   1500 characters per second
        '               5000 characters (5K)    -   1000 characters per second
        '               8000 characters (8K)    -    707 characters per second
        '
        '   At 11K, the machine locked up and an 'out of memory' error arose.  It is
        '   projected that the same machine would only do 418 characters per second
        '   on 10K, 58 characters per second on 20K, and 0.158 characters per second
        '   on 50K (all based on eighty trials).  It is strongly suggested that you
        '   encrypt 5K at a time and then concatenate the strings.  Furthermore, size
        '   needs to be taken into account.  The encrypted string will generally be
        '   between 3.9 and 4.1 times the size of the original string.  For instance,
        '   a 10k string might produce sizes between the ranges of 39K and 41K.
        '   Thus, it doesn't make sense to try to encrypt a 20MB file, unless you
        '   have the space.

OnError:
        On Error GoTo ErrHandler

Dimensions:
        Dim intMousePointer As Integer
        Dim dblCountLength As Double
        Dim intRandomNumber As Integer
        Dim strCurrentChar As String
        Dim intAscCurrentChar As Integer
        Dim intInverseAsc As Integer
        Dim intAddNinetyNine As Integer
        Dim dblMultiRandom As Double
        Dim dblWithRandom As Double
        Dim intCountPower As Integer
        Dim intPower As Integer
        Dim strConvertToBase As String

Constants:
        Const intLowerBounds = 10
        Const intUpperBounds = 28

MainCode:
        '   Start a For...Next loop that counts through the length of the parameter
        '   'StringToEncrypt'
104     For dblCountLength = 1 To Len(StringToEncrypt)

            '   Make sure random numbers do not hold any pattern
106         Randomize

            '   Choose a random integer between the constant 'intLowerBounds' and the
            '   constant 'intUpperBounds' and store it in 'intRandomNumber'
108         Let intRandomNumber = Int((intUpperBounds - intLowerBounds + 1) * Rnd + intLowerBounds)
                
            '   Select the next character in the parameter 'StringToEncrypt' based
            '   on the value of 'dblCountLength'
110         Let strCurrentChar = mid(StringToEncrypt, dblCountLength, 1)

            '   Find the ascii number associated with 'strCurrentChar'
112         Let intAscCurrentChar = Asc(strCurrentChar)

            '   Inverse the order of the numbers between 1 and 255 by subtracting the
            '   number from 256 (ie 1 turns into 255, 2 turns into 254, etc)
114         Let intInverseAsc = 256 - intAscCurrentChar

            '   Add 99 to the number
116         Let intAddNinetyNine = intInverseAsc + 99

            '   Multiply the integers 'intAddNinetyNine' and 'intRandomNumber'
            '   together
118         Let dblMultiRandom = intAddNinetyNine * intRandomNumber

            '   Insert the random number into the middle of the result of
            '   'dbsMultiRandom'
120         Let dblWithRandom = mid(dblMultiRandom, 1, 2) & intRandomNumber & mid(dblMultiRandom, 3, 2)
                
            '   Start a For...Next loop that counts through the viable powers of 93
            '   to be used to convert 'dblWithRandom' from base 10 to base 93
122         For intCountPower = 0 To 5

                '   Test to see if 'dblWithRandom' is large enough to accept the
                '   current power of 93 based on 'intCountPower'
124             If dblWithRandom / (93 ^ intCountPower) >= 1 Then
                    '   Store the power into the 'intPower' variable
126                 Let intPower = intCountPower

                '   Stop the test of 'dblWithRandom'
                End If
                
            '   Go to the next highest power of 93
128         Next intCountPower

            '   Let 'strConvertToBase' be equal to an empty string.
130         Let strConvertToBase = ""

            '   Start a For...Next loop that counts down through the viable powers
            '   of 93 based on the results of the test above
132         For intCountPower = intPower To 0 Step -1

                '   Divide 'dblWithRandom' by 93 to the power of 'intCountPower', add
                '   33, take only the integer, find the character associated with the
                '   number, and place it into the variable called 'strConvertToBase'
134             Let strConvertToBase = strConvertToBase & Chr(Int(dblWithRandom / (93 ^ intCountPower)) + 33)

                '   Let 'dblWithRandom' be equal to the remainder of the previous
                '   process
136             Let dblWithRandom = dblWithRandom Mod 93 ^ intCountPower

            '   Go to the next lowest power of 93
138         Next intCountPower

            '   Insert at the end of the function 'EncryptString' one character
            '   representing the length of 'strConvertToBase' and the value of
            '   'strConvertToBase'
140         Let EncryptString = EncryptString & Len(strConvertToBase) & strConvertToBase

        '   Go to the next character in the variable 'StringToEncrypt'
142     Next dblCountLength

        '   Stop execution of this function
        Exit Function

ErrHandler:
144     Call RegistrarError(Err.Number, Err.Description, "modCryptString.DecryptString", Erl)
146     Resume Next

End Function
