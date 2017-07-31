Attribute VB_Name = "Translit"
'Written in 2015 by Eduard E. Tikhenko <aquaried@gmail.com>
'
'To the extent possible under law, the author(s) have dedicated all copyright
'and related and neighboring rights to this software to the public domain
'worldwide. This software is distributed without any warranty.
'You should have received a copy of the CC0 Public Domain Dedication along
'with this software.
'If not, see <http://creativecommons.org/publicdomain/zero/1.0/>

Option Explicit

Function Transliteration(ByVal a As String) As String
    
    a = Replace(a, "�", "a")
    a = Replace(a, "�", "b")
    a = Replace(a, "�", "v")
    a = Replace(a, "�", "g")
    a = Replace(a, "�", "d")
    a = Replace(a, "�", "e")
    a = Replace(a, "�", "jo")
    a = Replace(a, "�", "zh")
    a = Replace(a, "�", "z")
    a = Replace(a, "�", "i")
    a = Replace(a, "�?", "j")
    a = Replace(a, "�", "k")
    a = Replace(a, "�", "l")
    a = Replace(a, "�", "m")
    a = Replace(a, "�", "n")
    a = Replace(a, "�", "o")
    a = Replace(a, "�", "p")
    a = Replace(a, "�", "r")
    a = Replace(a, "�", "s")
    a = Replace(a, "�", "t")
    a = Replace(a, "�", "u")
    a = Replace(a, "�", "f")
    a = Replace(a, "�", "h")
    a = Replace(a, "�", "c")
    a = Replace(a, "�", "ch")
    a = Replace(a, "�", "sh")
    a = Replace(a, "�", "sz")
    a = Replace(a, "�", "'")
    a = Replace(a, "�", "#")
    a = Replace(a, "�", "y")
    a = Replace(a, "�", "eh")
    a = Replace(a, "�", "ju")
    a = Replace(a, "�", "ja")
    
    a = Replace(a, "�", "A")
    a = Replace(a, "�", "B")
    a = Replace(a, "�", "V")
    a = Replace(a, "�", "G")
    a = Replace(a, "�", "D")
    a = Replace(a, "�", "E")
    a = Replace(a, "�", "Jo")
    a = Replace(a, "�", "Zh")
    a = Replace(a, "�", "Z")
    a = Replace(a, "�", "I")
    a = Replace(a, "�", "J")
    a = Replace(a, "�", "K")
    a = Replace(a, "�", "L")
    a = Replace(a, "�", "M")
    a = Replace(a, "�", "N")
    a = Replace(a, "�", "O")
    a = Replace(a, "�", "P")
    a = Replace(a, "�", "R")
    a = Replace(a, "�", "S")
    a = Replace(a, "�", "T")
    a = Replace(a, "�", "U")
    a = Replace(a, "�", "F")
    a = Replace(a, "�", "H")
    a = Replace(a, "�", "C")
    a = Replace(a, "�", "Ch")
    a = Replace(a, "�", "Sh")
    a = Replace(a, "�", "Sz")
    a = Replace(a, "�", "'")
    a = Replace(a, "�", "#")
    a = Replace(a, "�", "Y")
    a = Replace(a, "�", "Eh")
    a = Replace(a, "�", "Ju")
    a = Replace(a, "�", "Ja")
    
    Transliteration = a
End Function
