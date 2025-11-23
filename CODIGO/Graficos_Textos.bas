Attribute VB_Name = "Graficos_Textos"
'    Argentum 20 - Game Client Program
'    Copyright (C) 2022 - Noland Studios
'
'    This program is free software: you can redistribute it and/or modify
'    it under the terms of the GNU Affero General Public License as published by
'    the Free Software Foundation, either version 3 of the License, or
'    (at your option) any later version.
'
'    This program is distributed in the hope that it will be useful,
'    but WITHOUT ANY WARRANTY; without even the implied warranty of
'    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'    GNU Affero General Public License for more details.
'    You should have received a copy of the GNU Affero General Public License
'    along with this program.  If not, see <https://www.gnu.org/licenses/>.
'
'
Option Explicit

Public Type Fuente
    Tamanio As Integer
    Caracteres(0 To 11000) As Long 'indice de cada letra (ampliado para soportar caracteres extendidos)
End Type

Public font_count      As Long
Public font_last       As Long
Public font_list()     As D3DXFont

Public Const NUM_FONTS As Integer = 10

Public Fuentes(1 To 10) As Fuente

Public Sub Engine_Font_Initialize()
    On Error GoTo Engine_Font_Initialize_Err
    Dim A As Integer
    Fuentes(1).Tamanio = 9
    Fuentes(1).Caracteres(32) = 21494
    Fuentes(1).Caracteres(48) = 21452
    Fuentes(1).Caracteres(49) = 21453
    Fuentes(1).Caracteres(50) = 21454
    Fuentes(1).Caracteres(51) = 21455
    Fuentes(1).Caracteres(52) = 21456
    Fuentes(1).Caracteres(53) = 21457
    Fuentes(1).Caracteres(54) = 21458
    Fuentes(1).Caracteres(55) = 21459
    Fuentes(1).Caracteres(56) = 21460
    Fuentes(1).Caracteres(57) = 21461
    For A = 0 To 25
        Fuentes(1).Caracteres(A + 97) = 21400 + A
    Next A
    For A = 0 To 25
        Fuentes(1).Caracteres(A + 65) = 21426 + A
    Next A
    Fuentes(1).Caracteres(33) = 21462
    Fuentes(1).Caracteres(161) = 21463
    Fuentes(1).Caracteres(34) = 21464
    Fuentes(1).Caracteres(36) = 21465
    Fuentes(1).Caracteres(191) = 21466
    Fuentes(1).Caracteres(35) = 21467
    Fuentes(1).Caracteres(36) = 21468
    Fuentes(1).Caracteres(37) = 21469
    Fuentes(1).Caracteres(38) = 21470
    Fuentes(1).Caracteres(47) = 21471
    Fuentes(1).Caracteres(92) = 21472
    Fuentes(1).Caracteres(40) = 21473
    Fuentes(1).Caracteres(41) = 21474
    Fuentes(1).Caracteres(61) = 21475
    Fuentes(1).Caracteres(39) = 21476
    Fuentes(1).Caracteres(123) = 21477
    Fuentes(1).Caracteres(125) = 21478
    Fuentes(1).Caracteres(95) = 21479
    Fuentes(1).Caracteres(45) = 21480
    Fuentes(1).Caracteres(63) = 21465
    Fuentes(1).Caracteres(64) = 21481
    Fuentes(1).Caracteres(94) = 21482
    Fuentes(1).Caracteres(91) = 21483
    Fuentes(1).Caracteres(93) = 21484
    Fuentes(1).Caracteres(60) = 21485
    Fuentes(1).Caracteres(62) = 21486
    Fuentes(1).Caracteres(42) = 21487
    Fuentes(1).Caracteres(43) = 21488
    Fuentes(1).Caracteres(46) = 21489
    Fuentes(1).Caracteres(44) = 21490
    Fuentes(1).Caracteres(58) = 21491
    Fuentes(1).Caracteres(59) = 21492
    Fuentes(1).Caracteres(124) = 21493
    Fuentes(1).Caracteres(252) = 21800
    Fuentes(1).Caracteres(220) = 21801
    Fuentes(1).Caracteres(225) = 21802
    Fuentes(1).Caracteres(233) = 21803
    Fuentes(1).Caracteres(237) = 21804
    Fuentes(1).Caracteres(243) = 21805
    Fuentes(1).Caracteres(250) = 21806
    Fuentes(1).Caracteres(253) = 21807
    Fuentes(1).Caracteres(193) = 21808
    Fuentes(1).Caracteres(201) = 21809
    Fuentes(1).Caracteres(205) = 21810
    Fuentes(1).Caracteres(211) = 21811
    Fuentes(1).Caracteres(218) = 21812
    Fuentes(1).Caracteres(221) = 21813
    Fuentes(1).Caracteres(224) = 21814
    Fuentes(1).Caracteres(232) = 21815
    Fuentes(1).Caracteres(236) = 21816
    Fuentes(1).Caracteres(242) = 21817
    Fuentes(1).Caracteres(249) = 21818
    Fuentes(1).Caracteres(192) = 21819
    Fuentes(1).Caracteres(200) = 21820
    Fuentes(1).Caracteres(204) = 21821
    Fuentes(1).Caracteres(210) = 21822
    Fuentes(1).Caracteres(217) = 21823
    Fuentes(1).Caracteres(241) = 21824
    Fuentes(1).Caracteres(209) = 21825
    Fuentes(1).Caracteres(196) = 25238
    Fuentes(1).Caracteres(194) = 25239
    Fuentes(1).Caracteres(203) = 25240
    Fuentes(1).Caracteres(207) = 25241
    Fuentes(1).Caracteres(214) = 25242
    Fuentes(1).Caracteres(212) = 25243
    Fuentes(2).Tamanio = 9
    Fuentes(2).Caracteres(97) = 21936
    Fuentes(2).Caracteres(108) = 21937
    Fuentes(2).Caracteres(115) = 21938
    Fuentes(2).Caracteres(70) = 21939
    Fuentes(2).Caracteres(48) = 21940
    Fuentes(2).Caracteres(49) = 21941
    Fuentes(2).Caracteres(50) = 21942
    Fuentes(2).Caracteres(51) = 21943
    Fuentes(2).Caracteres(52) = 21944
    Fuentes(2).Caracteres(53) = 21945
    Fuentes(2).Caracteres(54) = 21946
    Fuentes(2).Caracteres(55) = 21947
    Fuentes(2).Caracteres(56) = 21948
    Fuentes(2).Caracteres(57) = 21949
    Fuentes(2).Caracteres(33) = 21950
    Fuentes(2).Caracteres(161) = 21951
    Fuentes(2).Caracteres(42) = 21952
    Fuentes(3).Tamanio = 40
    Fuentes(3).Caracteres(48) = 20428 '0
    Fuentes(3).Caracteres(49) = 20429 '1
    Fuentes(3).Caracteres(50) = 20430 '2
    Fuentes(3).Caracteres(51) = 20431 '3
    Fuentes(3).Caracteres(52) = 20432 '4
    Fuentes(3).Caracteres(53) = 20433 '5
    Fuentes(3).Caracteres(54) = 20434 '6
    Fuentes(3).Caracteres(55) = 20435 '7
    Fuentes(3).Caracteres(56) = 20436 '8
    Fuentes(3).Caracteres(57) = 20437 '9
    For A = 0 To 25
        Fuentes(3).Caracteres(A + 97) = 20477 + A 'Desde la a hasta la z (sin ñ)
    Next A
    For A = 0 To 25
        Fuentes(3).Caracteres(A + 65) = 20445 + A 'Desde la A hasta la Z (sin Ñ)
    Next A
    Fuentes(3).Caracteres(33) = 20413 '!
    Fuentes(3).Caracteres(161) = 20541 '¡
    Fuentes(3).Caracteres(34) = 20414 '"
    Fuentes(3).Caracteres(191) = 8488 '¿
    Fuentes(3).Caracteres(35) = 8332 '#
    Fuentes(3).Caracteres(36) = 20416    '$
    Fuentes(3).Caracteres(37) = 20417 '%
    Fuentes(3).Caracteres(38) = 20418 '&
    Fuentes(3).Caracteres(47) = 20427 '/
    Fuentes(3).Caracteres(92) = 8389 '\
    Fuentes(3).Caracteres(40) = 20420 '(
    Fuentes(3).Caracteres(41) = 20421 ')
    Fuentes(3).Caracteres(61) = 20441 '=
    Fuentes(3).Caracteres(39) = 24930 ''
    Fuentes(3).Caracteres(123) = 24932 ' '
    Fuentes(3).Caracteres(125) = 24931 '}
    Fuentes(3).Caracteres(95) = 20475  '_
    Fuentes(3).Caracteres(45) = 20425 '-
    Fuentes(3).Caracteres(63) = 20443 ' ?
    Fuentes(3).Caracteres(64) = 20444 '@
    Fuentes(3).Caracteres(94) = 20516 '^
    Fuentes(3).Caracteres(91) = 8388 '[
    Fuentes(3).Caracteres(93) = 8390 ']
    Fuentes(3).Caracteres(60) = 20440 '<
    Fuentes(3).Caracteres(62) = 20442 '>
    Fuentes(3).Caracteres(42) = 20422 '*
    Fuentes(3).Caracteres(43) = 20423 '+
    Fuentes(3).Caracteres(46) = 20426 '.
    Fuentes(3).Caracteres(44) = 20510 ',
    Fuentes(3).Caracteres(58) = 8355 ':
    Fuentes(3).Caracteres(59) = 8356 ';
    Fuentes(3).Caracteres(124) = 20504 '|
    Fuentes(3).Caracteres(252) = 24948 '    ü
    Fuentes(3).Caracteres(220) = 24949 'Ü
    Fuentes(3).Caracteres(225) = 8490 'á
    Fuentes(3).Caracteres(233) = 8498 'é
    Fuentes(3).Caracteres(237) = 8502 'í
    Fuentes(3).Caracteres(243) = 8508 'ó
    Fuentes(3).Caracteres(250) = 8515 'ú
    Fuentes(3).Caracteres(253) = 24955 'ý
    Fuentes(3).Caracteres(193) = 8490 'Á
    Fuentes(3).Caracteres(201) = 8498 'É
    Fuentes(3).Caracteres(205) = 8502 'Í
    Fuentes(3).Caracteres(211) = 8508 'Ó
    Fuentes(3).Caracteres(218) = 8515 'Ú
    Fuentes(3).Caracteres(221) = 24961 'Ý
    Fuentes(3).Caracteres(224) = 24962 'à
    Fuentes(3).Caracteres(232) = 24963 'è
    Fuentes(3).Caracteres(236) = 24964 'ì
    Fuentes(3).Caracteres(242) = 24965 'ò
    Fuentes(3).Caracteres(249) = 24966 'ù
    Fuentes(3).Caracteres(192) = 24967 'ü
    Fuentes(3).Caracteres(200) = 24968 '
    Fuentes(3).Caracteres(204) = 24969 '
    Fuentes(3).Caracteres(210) = 24970 '
    Fuentes(3).Caracteres(217) = 24971 '
    Fuentes(3).Caracteres(241) = 8506 'ñ
    Fuentes(3).Caracteres(209) = 24872 '
    Fuentes(3).Caracteres(196) = 24874 '
    Fuentes(3).Caracteres(194) = 24875 '
    Fuentes(3).Caracteres(203) = 24876 '
    Fuentes(3).Caracteres(207) = 24877 '
    Fuentes(3).Caracteres(214) = 24878 '
    Fuentes(3).Caracteres(212) = 24879 '
    Fuentes(3).Caracteres(172) = 20552 '¬
    Fuentes(3).Caracteres(186) = 20556 'º
    Fuentes(4).Tamanio = 3
    Fuentes(4).Caracteres(48) = 13852
    Fuentes(4).Caracteres(49) = 13853
    Fuentes(4).Caracteres(50) = 13854
    Fuentes(4).Caracteres(51) = 13855
    Fuentes(4).Caracteres(52) = 13856
    Fuentes(4).Caracteres(53) = 13857
    Fuentes(4).Caracteres(54) = 13858
    Fuentes(4).Caracteres(55) = 13859
    Fuentes(4).Caracteres(56) = 13860
    Fuentes(4).Caracteres(57) = 13861
    For A = 0 To 25
        Fuentes(4).Caracteres(A + 97) = 13800 + A
    Next A
    For A = 0 To 25
        Fuentes(4).Caracteres(A + 65) = 13826 + A
    Next A
    Fuentes(4).Caracteres(33) = 13862
    Fuentes(4).Caracteres(161) = 13863
    Fuentes(4).Caracteres(34) = 13864
    Fuentes(4).Caracteres(36) = 13865
    Fuentes(4).Caracteres(191) = 13866
    Fuentes(4).Caracteres(35) = 13867
    Fuentes(4).Caracteres(36) = 13868
    Fuentes(4).Caracteres(37) = 13869
    Fuentes(4).Caracteres(38) = 13870
    Fuentes(4).Caracteres(47) = 13871
    Fuentes(4).Caracteres(92) = 13872
    Fuentes(4).Caracteres(40) = 13873
    Fuentes(4).Caracteres(41) = 13874
    Fuentes(4).Caracteres(61) = 13875
    Fuentes(4).Caracteres(39) = 13876
    Fuentes(4).Caracteres(123) = 13877
    Fuentes(4).Caracteres(125) = 13878
    Fuentes(4).Caracteres(95) = 13879
    Fuentes(4).Caracteres(45) = 13880
    Fuentes(4).Caracteres(63) = 13865
    Fuentes(4).Caracteres(64) = 13881
    Fuentes(4).Caracteres(94) = 13882
    Fuentes(4).Caracteres(91) = 13883
    Fuentes(4).Caracteres(93) = 13884
    Fuentes(4).Caracteres(60) = 13885
    Fuentes(4).Caracteres(62) = 13886
    Fuentes(4).Caracteres(42) = 13887
    Fuentes(4).Caracteres(43) = 13888
    Fuentes(4).Caracteres(46) = 13889
    Fuentes(4).Caracteres(44) = 13890
    Fuentes(4).Caracteres(58) = 13891
    Fuentes(4).Caracteres(59) = 13892
    Fuentes(4).Caracteres(124) = 13893
    Fuentes(4).Caracteres(252) = 24948 '    ü
    Fuentes(4).Caracteres(220) = 24949 'Ü
    Fuentes(3).Caracteres(225) = 8490 'á
    Fuentes(3).Caracteres(233) = 8498 'é
    Fuentes(3).Caracteres(237) = 8502 'í
    Fuentes(3).Caracteres(243) = 8508 'ó
    Fuentes(3).Caracteres(250) = 8515 'ú
    Fuentes(3).Caracteres(253) = 24955 'ý
    Fuentes(3).Caracteres(193) = 8490 'Á
    Fuentes(3).Caracteres(201) = 8498 'É
    Fuentes(3).Caracteres(205) = 8502 'Í
    Fuentes(3).Caracteres(211) = 8508 'Ó
    Fuentes(3).Caracteres(218) = 8515 'Ú
    Fuentes(3).Caracteres(221) = 24961 'Ý
    Fuentes(3).Caracteres(224) = 24962 'à
    Fuentes(3).Caracteres(232) = 24963 'è
    Fuentes(3).Caracteres(236) = 24964 'ì
    Fuentes(3).Caracteres(242) = 24965 'ò
    Fuentes(3).Caracteres(249) = 24966 'ù
    Fuentes(3).Caracteres(192) = 24967 'ü
    Fuentes(3).Caracteres(200) = 24968 '
    Fuentes(3).Caracteres(204) = 24969 '
    Fuentes(3).Caracteres(210) = 24970 '
    Fuentes(3).Caracteres(217) = 24971 '
    Fuentes(3).Caracteres(241) = 8506 'ñ
    Fuentes(3).Caracteres(209) = 24872 '
    Fuentes(3).Caracteres(196) = 24874 '
    Fuentes(3).Caracteres(194) = 24875 '
    Fuentes(3).Caracteres(203) = 24876 '
    Fuentes(3).Caracteres(207) = 24877 '
    Fuentes(3).Caracteres(214) = 24878 '
    Fuentes(3).Caracteres(212) = 24879 '
    Fuentes(3).Caracteres(172) = 20552 '¬
    Fuentes(3).Caracteres(186) = 20556 'º
    Fuentes(1).Caracteres(196) = 25238
    Fuentes(1).Caracteres(194) = 25239
    Fuentes(1).Caracteres(203) = 25240
    Fuentes(1).Caracteres(207) = 25241
    Fuentes(1).Caracteres(214) = 25242
    Fuentes(1).Caracteres(212) = 25243
    Fuentes(5).Tamanio = 50
    Fuentes(5).Caracteres(48) = 30127
    Fuentes(5).Caracteres(49) = 30128
    Fuentes(5).Caracteres(50) = 30129
    Fuentes(5).Caracteres(51) = 30130
    Fuentes(5).Caracteres(52) = 30131
    Fuentes(5).Caracteres(53) = 30132
    Fuentes(5).Caracteres(54) = 30133
    Fuentes(5).Caracteres(55) = 30134
    Fuentes(5).Caracteres(56) = 30135
    Fuentes(5).Caracteres(57) = 30136
    For A = 0 To 25
        Fuentes(5).Caracteres(A + 97) = 30176 + A
    Next A
    For A = 0 To 25
        Fuentes(5).Caracteres(A + 65) = 30144 + A
    Next A
    Fuentes(5).Caracteres(33) = 30112 '!
    Fuentes(5).Caracteres(161) = 20541 '¡
    Fuentes(5).Caracteres(34) = 30113 '"
    Fuentes(5).Caracteres(191) = 8488 '¿
    Fuentes(5).Caracteres(35) = 8332 '#
    Fuentes(5).Caracteres(36) = 20416    '$
    Fuentes(5).Caracteres(37) = 20417 '%
    Fuentes(5).Caracteres(38) = 20418 '&
    Fuentes(5).Caracteres(47) = 20427 '/
    Fuentes(5).Caracteres(92) = 8389 '\
    Fuentes(5).Caracteres(40) = 30119 '(
    Fuentes(5).Caracteres(41) = 30120 ')
    Fuentes(5).Caracteres(61) = 30140 '=
    Fuentes(5).Caracteres(39) = 24930 ''
    Fuentes(5).Caracteres(123) = 24932 ' '
    Fuentes(5).Caracteres(125) = 24931 '}
    Fuentes(5).Caracteres(95) = 20475  '_
    Fuentes(5).Caracteres(45) = 20425 '-
    Fuentes(5).Caracteres(63) = 20443 ' ?
    Fuentes(5).Caracteres(64) = 20444 '@
    Fuentes(5).Caracteres(94) = 20516 '^
    Fuentes(5).Caracteres(91) = 8388 '[
    Fuentes(5).Caracteres(93) = 8390 ']
    Fuentes(5).Caracteres(60) = 30139 '<
    Fuentes(5).Caracteres(62) = 30141 '>
    Fuentes(5).Caracteres(42) = 20422 '*
    Fuentes(5).Caracteres(43) = 20423 '+
    Fuentes(5).Caracteres(46) = 20426 '.
    Fuentes(5).Caracteres(44) = 20510 ',
    Fuentes(5).Caracteres(58) = 8355 ':
    Fuentes(5).Caracteres(59) = 8356 ';
    Fuentes(5).Caracteres(124) = 20504 '|
    Fuentes(5).Caracteres(252) = 24948 '    ü
    Fuentes(5).Caracteres(220) = 24949 'Ü
    Fuentes(5).Caracteres(225) = 30304 'á
    Fuentes(5).Caracteres(233) = 30312 'é
    Fuentes(5).Caracteres(237) = 30316 'í
    Fuentes(5).Caracteres(243) = 30322 'ó
    Fuentes(5).Caracteres(250) = 30329 'ú
    Fuentes(5).Caracteres(253) = 24955 'ý
    Fuentes(5).Caracteres(193) = 30272 'Á
    Fuentes(5).Caracteres(201) = 30280 'É
    Fuentes(5).Caracteres(205) = 8502 'Í
    Fuentes(5).Caracteres(211) = 30290 'Ó
    Fuentes(5).Caracteres(218) = 8515 'Ú
    Fuentes(5).Caracteres(221) = 24961 'Ý
    Fuentes(5).Caracteres(224) = 24962 'à
    Fuentes(5).Caracteres(232) = 24963 'è
    Fuentes(5).Caracteres(236) = 24964 'ì
    Fuentes(5).Caracteres(242) = 24965 'ò
    Fuentes(5).Caracteres(249) = 24966 'ù
    Fuentes(5).Caracteres(192) = 24967 'ü
    Fuentes(5).Caracteres(200) = 24968 '
    Fuentes(5).Caracteres(204) = 24969 '
    Fuentes(5).Caracteres(210) = 24970 '
    Fuentes(5).Caracteres(217) = 24971 '
    Fuentes(5).Caracteres(241) = 30288 'ñ
    Fuentes(5).Caracteres(209) = 24872 '
    Fuentes(5).Caracteres(196) = 24874 '
    Fuentes(5).Caracteres(194) = 30305 'â
    Fuentes(5).Caracteres(203) = 24876 '
    Fuentes(5).Caracteres(207) = 24877 '
    Fuentes(5).Caracteres(214) = 24878 '
    Fuentes(5).Caracteres(212) = 24879 '
    Fuentes(5).Caracteres(172) = 20552 '¬
    Fuentes(5).Caracteres(186) = 20556 'º
    Fuentes(6).Tamanio = 50
    Fuentes(6).Caracteres(48) = 45866
    Fuentes(6).Caracteres(49) = 45867
    Fuentes(6).Caracteres(50) = 45868
    Fuentes(6).Caracteres(51) = 45869
    Fuentes(6).Caracteres(52) = 45870
    Fuentes(6).Caracteres(53) = 45871
    Fuentes(6).Caracteres(54) = 45872
    Fuentes(6).Caracteres(55) = 45873
    Fuentes(6).Caracteres(56) = 45874
    Fuentes(6).Caracteres(57) = 45875
    For A = 0 To 25
        Fuentes(6).Caracteres(A + 97) = 45915 + A
    Next A
    For A = 0 To 25
        Fuentes(6).Caracteres(A + 65) = 45883 + A
    Next A
    Fuentes(6).Caracteres(33) = 13862
    Fuentes(6).Caracteres(161) = 13863
    Fuentes(6).Caracteres(34) = 13864
    Fuentes(6).Caracteres(36) = 13865
    Fuentes(6).Caracteres(191) = 13866
    Fuentes(6).Caracteres(35) = 13867
    Fuentes(6).Caracteres(36) = 13868
    Fuentes(6).Caracteres(37) = 13869
    Fuentes(6).Caracteres(38) = 13870
    Fuentes(6).Caracteres(47) = 13871
    Fuentes(6).Caracteres(92) = 13872
    Fuentes(6).Caracteres(40) = 13873
    Fuentes(6).Caracteres(41) = 13874
    Fuentes(6).Caracteres(61) = 13875
    Fuentes(6).Caracteres(39) = 13876
    Fuentes(6).Caracteres(123) = 13877
    Fuentes(6).Caracteres(125) = 13878
    Fuentes(6).Caracteres(95) = 13879
    Fuentes(6).Caracteres(45) = 13880
    Fuentes(6).Caracteres(63) = 13865
    Fuentes(6).Caracteres(64) = 13881
    Fuentes(6).Caracteres(94) = 13882
    Fuentes(6).Caracteres(91) = 13883
    Fuentes(6).Caracteres(93) = 13884
    Fuentes(6).Caracteres(60) = 13885
    Fuentes(6).Caracteres(62) = 13886
    Fuentes(6).Caracteres(42) = 13887
    Fuentes(6).Caracteres(43) = 13888
    Fuentes(6).Caracteres(46) = 13889
    Fuentes(6).Caracteres(44) = 13890
    Fuentes(6).Caracteres(58) = 13891
    Fuentes(6).Caracteres(59) = 13892
    Fuentes(6).Caracteres(124) = 13893
    Fuentes(6).Caracteres(252) = 18200
    Fuentes(6).Caracteres(220) = 18201
    Fuentes(6).Caracteres(225) = 18202
    Fuentes(6).Caracteres(233) = 18203
    Fuentes(6).Caracteres(237) = 18204
    Fuentes(6).Caracteres(243) = 18205
    Fuentes(6).Caracteres(250) = 18206
    Fuentes(6).Caracteres(253) = 18207
    Fuentes(6).Caracteres(193) = 18208
    Fuentes(6).Caracteres(201) = 18209
    Fuentes(6).Caracteres(205) = 18210
    Fuentes(6).Caracteres(211) = 18211
    Fuentes(6).Caracteres(218) = 18212
    Fuentes(6).Caracteres(221) = 18213
    Fuentes(6).Caracteres(224) = 18214
    Fuentes(6).Caracteres(232) = 18215
    Fuentes(6).Caracteres(236) = 18216
    Fuentes(6).Caracteres(242) = 18217
    Fuentes(6).Caracteres(249) = 18218
    Fuentes(6).Caracteres(192) = 18219
    Fuentes(6).Caracteres(200) = 18220
    Fuentes(6).Caracteres(204) = 18221
    Fuentes(6).Caracteres(210) = 18222
    Fuentes(6).Caracteres(217) = 18223
    Fuentes(6).Caracteres(241) = 18224
    Fuentes(6).Caracteres(209) = 18225
    Fuentes(7).Tamanio = 50    ' Cardo
    For A = 0 To 25
        Fuentes(7).Caracteres(A + 65) = 26592 + A
        Fuentes(7).Caracteres(A + 97) = 26618 + A
    Next A
    For A = 0 To 9
        Fuentes(7).Caracteres(A + 48) = 26692 + A
    Next A
    Fuentes(7).Caracteres(193) = 26644
    Fuentes(7).Caracteres(201) = 26645
    Fuentes(7).Caracteres(205) = 26646
    Fuentes(7).Caracteres(211) = 26647
    Fuentes(7).Caracteres(218) = 26648
    Fuentes(7).Caracteres(225) = 26649
    Fuentes(7).Caracteres(233) = 26650
    Fuentes(7).Caracteres(237) = 26651
    Fuentes(7).Caracteres(243) = 26652
    Fuentes(7).Caracteres(250) = 26653
    Fuentes(7).Caracteres(192) = 26654
    Fuentes(7).Caracteres(200) = 26655
    Fuentes(7).Caracteres(204) = 26656
    Fuentes(7).Caracteres(210) = 26657
    Fuentes(7).Caracteres(217) = 26658
    Fuentes(7).Caracteres(224) = 26659
    Fuentes(7).Caracteres(232) = 26660
    Fuentes(7).Caracteres(236) = 26661
    Fuentes(7).Caracteres(242) = 26662
    Fuentes(7).Caracteres(249) = 26663
    Fuentes(7).Caracteres(194) = 26664
    Fuentes(7).Caracteres(202) = 26665
    Fuentes(7).Caracteres(206) = 26666
    Fuentes(7).Caracteres(212) = 26667
    Fuentes(7).Caracteres(219) = 26668
    Fuentes(7).Caracteres(226) = 26669
    Fuentes(7).Caracteres(234) = 26670
    Fuentes(7).Caracteres(238) = 26671
    Fuentes(7).Caracteres(244) = 26672
    Fuentes(7).Caracteres(251) = 26673
    Fuentes(7).Caracteres(196) = 26674
    Fuentes(7).Caracteres(203) = 26675
    Fuentes(7).Caracteres(207) = 26676
    Fuentes(7).Caracteres(214) = 26677
    Fuentes(7).Caracteres(220) = 26678
    Fuentes(7).Caracteres(228) = 26679
    Fuentes(7).Caracteres(235) = 26680
    Fuentes(7).Caracteres(239) = 26681
    Fuentes(7).Caracteres(246) = 26682
    Fuentes(7).Caracteres(252) = 26683
    Fuentes(7).Caracteres(195) = 26684
    Fuentes(7).Caracteres(213) = 26685
    Fuentes(7).Caracteres(227) = 26686
    Fuentes(7).Caracteres(245) = 26687
    Fuentes(7).Caracteres(209) = 26688
    Fuentes(7).Caracteres(241) = 26689
    Fuentes(7).Caracteres(199) = 26690
    Fuentes(7).Caracteres(231) = 26691
    Fuentes(7).Caracteres(161) = 26702
    Fuentes(7).Caracteres(33) = 26703
    Fuentes(7).Caracteres(191) = 26704
    Fuentes(7).Caracteres(63) = 26705
    Fuentes(7).Caracteres(46) = 26706
    Fuentes(7).Caracteres(44) = 26707
    Fuentes(7).Caracteres(59) = 26708
    Fuentes(7).Caracteres(58) = 26709
    Fuentes(7).Caracteres(8230) = 26710
    Fuentes(7).Caracteres(8212) = 26711
    Fuentes(7).Caracteres(45) = 26712
    Fuentes(7).Caracteres(95) = 26713
    Fuentes(7).Caracteres(40) = 26714
    Fuentes(7).Caracteres(41) = 26715
    Fuentes(7).Caracteres(91) = 26716
    Fuentes(7).Caracteres(93) = 26717
    Fuentes(7).Caracteres(123) = 26718
    Fuentes(7).Caracteres(125) = 26719
    Fuentes(7).Caracteres(171) = 26720
    Fuentes(7).Caracteres(187) = 26721
    Fuentes(7).Caracteres(34) = 26722
    Fuentes(7).Caracteres(39) = 26723
    Fuentes(7).Caracteres(8216) = 26724
    Fuentes(7).Caracteres(8217) = 26725
    Fuentes(7).Caracteres(8220) = 26726
    Fuentes(7).Caracteres(8221) = 26727
    Fuentes(7).Caracteres(47) = 26728
    Fuentes(7).Caracteres(92) = 26729
    Fuentes(7).Caracteres(124) = 26730
    Fuentes(7).Caracteres(64) = 26731
    Fuentes(7).Caracteres(35) = 26732
    Fuentes(7).Caracteres(37) = 26733
    Fuentes(7).Caracteres(38) = 26734
    Fuentes(7).Caracteres(8432) = 26735
    Fuentes(7).Caracteres(43) = 26736
    Fuentes(7).Caracteres(45) = 26737
    Fuentes(7).Caracteres(60) = 26738
    Fuentes(7).Caracteres(62) = 26739
    Fuentes(7).Caracteres(94) = 26740
    Fuentes(7).Caracteres(95) = 26741
    Fuentes(7).Caracteres(126) = 26742
    Fuentes(7).Caracteres(180) = 26743
    Fuentes(7).Caracteres(172) = 26744
    Fuentes(7).Caracteres(36) = 26745
    Fuentes(7).Caracteres(8364) = 26746
    Fuentes(7).Caracteres(163) = 26747
    Fuentes(7).Caracteres(162) = 26748
    Fuentes(7).Caracteres(165) = 26749
    Fuentes(7).Caracteres(176) = 26750
    Fuentes(7).Caracteres(167) = 26751
    Fuentes(7).Caracteres(182) = 26752
    Fuentes(7).Caracteres(177) = 26753
    Fuentes(7).Caracteres(120) = 26754
    Fuentes(7).Caracteres(247) = 26755
    Fuentes(7).Caracteres(8240) = 26756
    Fuentes(7).Caracteres(8734) = 26757
    Fuentes(7).Caracteres(8776) = 26758
    Fuentes(7).Caracteres(8800) = 26759
    Fuentes(7).Caracteres(8804) = 26760
    Fuentes(7).Caracteres(8805) = 26761
    Fuentes(7).Caracteres(8730) = 26762
    Fuentes(7).Caracteres(8721) = 26763
    Fuentes(7).Caracteres(916) = 26764
    Fuentes(7).Caracteres(181) = 26765
    Fuentes(7).Caracteres(9679) = 26766
    Fuentes(8).Tamanio = 50    ' Cardo
    For A = 0 To 25
        Fuentes(8).Caracteres(A + 65) = 27510 + A
        Fuentes(8).Caracteres(A + 97) = 27536 + A
    Next A
    For A = 0 To 9
        Fuentes(8).Caracteres(A + 48) = 27610 + A
    Next A
    Fuentes(8).Caracteres(193) = 27562
    Fuentes(8).Caracteres(201) = 27563
    Fuentes(8).Caracteres(205) = 27564
    Fuentes(8).Caracteres(211) = 27565
    Fuentes(8).Caracteres(218) = 27566
    Fuentes(8).Caracteres(225) = 27567
    Fuentes(8).Caracteres(233) = 27568
    Fuentes(8).Caracteres(237) = 27569
    Fuentes(8).Caracteres(243) = 27570
    Fuentes(8).Caracteres(250) = 27571
    Fuentes(8).Caracteres(192) = 27572
    Fuentes(8).Caracteres(200) = 27573
    Fuentes(8).Caracteres(204) = 27574
    Fuentes(8).Caracteres(210) = 27575
    Fuentes(8).Caracteres(217) = 27576
    Fuentes(8).Caracteres(224) = 27577
    Fuentes(8).Caracteres(232) = 27578
    Fuentes(8).Caracteres(236) = 27579
    Fuentes(8).Caracteres(242) = 27580
    Fuentes(8).Caracteres(249) = 27581
    Fuentes(8).Caracteres(194) = 27582
    Fuentes(8).Caracteres(202) = 27583
    Fuentes(8).Caracteres(206) = 27584
    Fuentes(8).Caracteres(212) = 27585
    Fuentes(8).Caracteres(219) = 27586
    Fuentes(8).Caracteres(226) = 27587
    Fuentes(8).Caracteres(234) = 27588
    Fuentes(8).Caracteres(238) = 27589
    Fuentes(8).Caracteres(244) = 27590
    Fuentes(8).Caracteres(251) = 27591
    Fuentes(8).Caracteres(196) = 27592
    Fuentes(8).Caracteres(203) = 27593
    Fuentes(8).Caracteres(207) = 27594
    Fuentes(8).Caracteres(214) = 27595
    Fuentes(8).Caracteres(220) = 27596
    Fuentes(8).Caracteres(228) = 27597
    Fuentes(8).Caracteres(235) = 27598
    Fuentes(8).Caracteres(239) = 27599
    Fuentes(8).Caracteres(246) = 27600
    Fuentes(8).Caracteres(252) = 27601
    Fuentes(8).Caracteres(195) = 27602
    Fuentes(8).Caracteres(213) = 27603
    Fuentes(8).Caracteres(227) = 27604
    Fuentes(8).Caracteres(245) = 27605
    Fuentes(8).Caracteres(209) = 27606
    Fuentes(8).Caracteres(241) = 27607
    Fuentes(8).Caracteres(199) = 27608
    Fuentes(8).Caracteres(231) = 27609
    Fuentes(8).Caracteres(161) = 27620
    Fuentes(8).Caracteres(33) = 27621
    Fuentes(8).Caracteres(191) = 27622
    Fuentes(8).Caracteres(63) = 27623
    Fuentes(8).Caracteres(46) = 27624
    Fuentes(8).Caracteres(44) = 27625
    Fuentes(8).Caracteres(59) = 27626
    Fuentes(8).Caracteres(58) = 27627
    Fuentes(8).Caracteres(8230) = 27628
    Fuentes(8).Caracteres(8212) = 27629
    Fuentes(8).Caracteres(45) = 27630
    Fuentes(8).Caracteres(95) = 27631
    Fuentes(8).Caracteres(40) = 27632
    Fuentes(8).Caracteres(41) = 27633
    Fuentes(8).Caracteres(91) = 27634
    Fuentes(8).Caracteres(93) = 27635
    Fuentes(8).Caracteres(123) = 27636
    Fuentes(8).Caracteres(125) = 27637
    Fuentes(8).Caracteres(171) = 27638
    Fuentes(8).Caracteres(187) = 27639
    Fuentes(8).Caracteres(34) = 27640
    Fuentes(8).Caracteres(39) = 27641
    Fuentes(8).Caracteres(8216) = 27642
    Fuentes(8).Caracteres(8217) = 27643
    Fuentes(8).Caracteres(8220) = 27644
    Fuentes(8).Caracteres(8221) = 27645
    Fuentes(8).Caracteres(47) = 27646
    Fuentes(8).Caracteres(92) = 27647
    Fuentes(8).Caracteres(124) = 27648
    Fuentes(8).Caracteres(64) = 27649
    Fuentes(8).Caracteres(35) = 27650
    Fuentes(8).Caracteres(37) = 27651
    Fuentes(8).Caracteres(38) = 27652
    Fuentes(8).Caracteres(8432) = 27653
    Fuentes(8).Caracteres(43) = 27654
    Fuentes(8).Caracteres(45) = 27655
    Fuentes(8).Caracteres(60) = 27656
    Fuentes(8).Caracteres(62) = 27657
    Fuentes(8).Caracteres(94) = 27658
    Fuentes(8).Caracteres(95) = 27659
    Fuentes(8).Caracteres(126) = 27660
    Fuentes(8).Caracteres(180) = 27661
    Fuentes(8).Caracteres(172) = 27662
    Fuentes(8).Caracteres(36) = 27663
    Fuentes(8).Caracteres(8364) = 27664
    Fuentes(8).Caracteres(163) = 27665
    Fuentes(8).Caracteres(162) = 27666
    Fuentes(8).Caracteres(165) = 27667
    Fuentes(8).Caracteres(176) = 27668
    Fuentes(8).Caracteres(167) = 27669
    Fuentes(8).Caracteres(182) = 27670
    Fuentes(8).Caracteres(177) = 27671
    Fuentes(8).Caracteres(120) = 27672
    Fuentes(8).Caracteres(247) = 27673
    Fuentes(8).Caracteres(8240) = 27674
    Fuentes(8).Caracteres(8734) = 27675
    Fuentes(8).Caracteres(8776) = 27676
    Fuentes(8).Caracteres(8800) = 27677
    Fuentes(8).Caracteres(8804) = 27678
    Fuentes(8).Caracteres(8805) = 27679
    Fuentes(8).Caracteres(8730) = 27680
    Fuentes(8).Caracteres(8721) = 27681
    Fuentes(8).Caracteres(916) = 27682
    Fuentes(8).Caracteres(181) = 27683
    Fuentes(8).Caracteres(9679) = 27684
    
    
    Fuentes(9).Caracteres(65) = 87267  ' A
    Fuentes(9).Caracteres(66) = 87268  ' B
    Fuentes(9).Caracteres(67) = 87269  ' C
    Fuentes(9).Caracteres(68) = 87270  ' D
    Fuentes(9).Caracteres(69) = 87271  ' E
    Fuentes(9).Caracteres(70) = 87272  ' F
    Fuentes(9).Caracteres(71) = 87273  ' G
    Fuentes(9).Caracteres(72) = 87274  ' H
    Fuentes(9).Caracteres(73) = 87275  ' I
    Fuentes(9).Caracteres(74) = 87276  ' J
    Fuentes(9).Caracteres(75) = 87277  ' K
    Fuentes(9).Caracteres(76) = 87278  ' L
    Fuentes(9).Caracteres(77) = 87279  ' M
    Fuentes(9).Caracteres(78) = 87280  ' N
    Fuentes(9).Caracteres(79) = 87281  ' O
    Fuentes(9).Caracteres(80) = 87282  ' P
    Fuentes(9).Caracteres(81) = 87283  ' Q
    Fuentes(9).Caracteres(82) = 87284  ' R
    Fuentes(9).Caracteres(83) = 87285  ' S
    Fuentes(9).Caracteres(84) = 87286  ' T
    Fuentes(9).Caracteres(85) = 87287  ' U
    Fuentes(9).Caracteres(86) = 87288  ' V
    Fuentes(9).Caracteres(87) = 87289  ' W
    Fuentes(9).Caracteres(88) = 87290  ' X
    Fuentes(9).Caracteres(89) = 87291  ' Y
    Fuentes(9).Caracteres(90) = 87292  ' Z
    Fuentes(9).Caracteres(97) = 87293  ' a
    Fuentes(9).Caracteres(98) = 87294  ' b
    Fuentes(9).Caracteres(99) = 87295  ' c
    Fuentes(9).Caracteres(100) = 87296  ' d
    Fuentes(9).Caracteres(101) = 87297  ' e
    Fuentes(9).Caracteres(102) = 87298  ' f
    Fuentes(9).Caracteres(103) = 87299  ' g
    Fuentes(9).Caracteres(104) = 87300  ' h
    Fuentes(9).Caracteres(105) = 87301  ' i
    Fuentes(9).Caracteres(106) = 87302  ' j
    Fuentes(9).Caracteres(107) = 87303  ' k
    Fuentes(9).Caracteres(108) = 87304  ' l
    Fuentes(9).Caracteres(109) = 87305  ' m
    Fuentes(9).Caracteres(110) = 87306  ' n
    Fuentes(9).Caracteres(111) = 87307  ' o
    Fuentes(9).Caracteres(112) = 87308  ' p
    Fuentes(9).Caracteres(113) = 87309  ' q
    Fuentes(9).Caracteres(114) = 87310  ' r
    Fuentes(9).Caracteres(115) = 87311  ' s
    Fuentes(9).Caracteres(116) = 87312  ' t
    Fuentes(9).Caracteres(117) = 87313  ' u
    Fuentes(9).Caracteres(118) = 87314  ' v
    Fuentes(9).Caracteres(119) = 87315  ' w
    Fuentes(9).Caracteres(120) = 87316  ' x
    Fuentes(9).Caracteres(121) = 87317  ' y
    Fuentes(9).Caracteres(122) = 87318  ' z
    Fuentes(9).Caracteres(193) = 87319  ' Á
    Fuentes(9).Caracteres(201) = 87320  ' É
    Fuentes(9).Caracteres(205) = 87321  ' Í
    Fuentes(9).Caracteres(211) = 87322  ' Ó
    Fuentes(9).Caracteres(218) = 87323  ' Ú
    Fuentes(9).Caracteres(225) = 87324  ' á
    Fuentes(9).Caracteres(233) = 87325  ' é
    Fuentes(9).Caracteres(237) = 87326  ' í
    Fuentes(9).Caracteres(243) = 87327  ' ó
    Fuentes(9).Caracteres(250) = 87328  ' ú
    Fuentes(9).Caracteres(192) = 87329  ' À
    Fuentes(9).Caracteres(200) = 87330  ' È
    Fuentes(9).Caracteres(204) = 87331  ' Ì
    Fuentes(9).Caracteres(210) = 87332  ' Ò
    Fuentes(9).Caracteres(217) = 87333  ' Ù
    Fuentes(9).Caracteres(224) = 87334  ' à
    Fuentes(9).Caracteres(232) = 87335  ' è
    Fuentes(9).Caracteres(236) = 87336  ' ì
    Fuentes(9).Caracteres(242) = 87337  ' ò
    Fuentes(9).Caracteres(249) = 87338  ' ù
    Fuentes(9).Caracteres(194) = 87339  ' Â
    Fuentes(9).Caracteres(202) = 87340  ' Ê
    Fuentes(9).Caracteres(206) = 87341  ' Î
    Fuentes(9).Caracteres(212) = 87342  ' Ô
    Fuentes(9).Caracteres(219) = 87343  ' Û
    Fuentes(9).Caracteres(226) = 87344  ' â
    Fuentes(9).Caracteres(234) = 87345  ' ê
    Fuentes(9).Caracteres(238) = 87346  ' î
    Fuentes(9).Caracteres(244) = 87347  ' ô
    Fuentes(9).Caracteres(251) = 87348  ' û
    Fuentes(9).Caracteres(196) = 87349  ' Ä
    Fuentes(9).Caracteres(203) = 87350  ' Ë
    Fuentes(9).Caracteres(207) = 87351  ' Ï
    Fuentes(9).Caracteres(214) = 87352  ' Ö
    Fuentes(9).Caracteres(220) = 87353  ' Ü
    Fuentes(9).Caracteres(228) = 87354  ' ä
    Fuentes(9).Caracteres(235) = 87355  ' ë
    Fuentes(9).Caracteres(239) = 87356  ' ï
    Fuentes(9).Caracteres(246) = 87357  ' ö
    Fuentes(9).Caracteres(252) = 87358  ' ü
    Fuentes(9).Caracteres(195) = 87359  ' Ã
    Fuentes(9).Caracteres(213) = 87360  ' Õ
    Fuentes(9).Caracteres(227) = 87361  ' ã
    Fuentes(9).Caracteres(245) = 87362  ' õ
    Fuentes(9).Caracteres(209) = 87363  ' Ñ
    Fuentes(9).Caracteres(241) = 87364  ' ñ
    Fuentes(9).Caracteres(199) = 87365  ' Ç
    Fuentes(9).Caracteres(231) = 87366  ' ç
    Fuentes(9).Caracteres(48) = 87367  ' 0
    Fuentes(9).Caracteres(49) = 87368  ' 1
    Fuentes(9).Caracteres(50) = 87369  ' 2
    Fuentes(9).Caracteres(51) = 87370  ' 3
    Fuentes(9).Caracteres(52) = 87371  ' 4
    Fuentes(9).Caracteres(53) = 87372  ' 5
    Fuentes(9).Caracteres(54) = 87373  ' 6
    Fuentes(9).Caracteres(55) = 87374  ' 7
    Fuentes(9).Caracteres(56) = 87375  ' 8
    Fuentes(9).Caracteres(57) = 87376  ' 9
    Fuentes(9).Caracteres(161) = 87377  ' ¡
    Fuentes(9).Caracteres(33) = 87378  ' !
    Fuentes(9).Caracteres(191) = 87379  ' ¿
    Fuentes(9).Caracteres(63) = 87380  ' ?
    Fuentes(9).Caracteres(46) = 87381  ' .
    Fuentes(9).Caracteres(44) = 87382  ' ,
    Fuentes(9).Caracteres(59) = 87383  ' ;
    Fuentes(9).Caracteres(58) = 87384  ' :
    Fuentes(9).Caracteres(8230) = 87385  ' …
    Fuentes(9).Caracteres(8212) = 87386  ' —
    Fuentes(9).Caracteres(45) = 87387  ' -
    Fuentes(9).Caracteres(95) = 87388  ' _
    Fuentes(9).Caracteres(40) = 87389  ' (
    Fuentes(9).Caracteres(41) = 87390  ' )
    Fuentes(9).Caracteres(91) = 87391  ' [
    Fuentes(9).Caracteres(93) = 87392  ' ]
    Fuentes(9).Caracteres(123) = 87393  ' {
    Fuentes(9).Caracteres(125) = 87394  ' }
    Fuentes(9).Caracteres(171) = 87395  ' «
    Fuentes(9).Caracteres(187) = 87396  ' »
    Fuentes(9).Caracteres(34) = 87397  ' "
    Fuentes(9).Caracteres(39) = 87398  ' ''
    Fuentes(9).Caracteres(8216) = 87399  ' ‘
    Fuentes(9).Caracteres(8217) = 87400  ' ’
    Fuentes(9).Caracteres(8220) = 87401  ' “
    Fuentes(9).Caracteres(8221) = 87402  ' ”
    Fuentes(9).Caracteres(47) = 87403  ' /
    Fuentes(9).Caracteres(92) = 87404  ' \
    Fuentes(9).Caracteres(124) = 87405  ' |
    Fuentes(9).Caracteres(64) = 87406  ' @
    Fuentes(9).Caracteres(35) = 87407  ' #
    Fuentes(9).Caracteres(37) = 87408  ' %
    Fuentes(9).Caracteres(38) = 87409  ' &
    Fuentes(9).Caracteres(8432) = 87410  ' ?
    Fuentes(9).Caracteres(43) = 87411  ' +
    Fuentes(9).Caracteres(45) = 87412  ' -
    Fuentes(9).Caracteres(60) = 87413  ' <
    Fuentes(9).Caracteres(62) = 87414  ' >
    Fuentes(9).Caracteres(94) = 87415  ' ^
    Fuentes(9).Caracteres(95) = 87416  ' _
    Fuentes(9).Caracteres(126) = 87417  ' ~
    Fuentes(9).Caracteres(180) = 87418  ' ´
    Fuentes(9).Caracteres(172) = 87419  ' ¬
    Fuentes(9).Caracteres(36) = 87420  ' $
    Fuentes(9).Caracteres(8364) = 87421  ' €
    Fuentes(9).Caracteres(163) = 87422  ' £
    Fuentes(9).Caracteres(162) = 87423  ' ¢
    Fuentes(9).Caracteres(165) = 87424  ' ¥
    Fuentes(9).Caracteres(176) = 87425  ' °
    Fuentes(9).Caracteres(167) = 87426  ' §
    Fuentes(9).Caracteres(182) = 87427  ' ¶
    Fuentes(9).Caracteres(177) = 87428  ' ±
    Fuentes(9).Caracteres(120) = 87429  ' x
    Fuentes(9).Caracteres(247) = 87430  ' ÷
    Fuentes(9).Caracteres(8240) = 87431  ' ‰
    Fuentes(9).Caracteres(8734) = 87432  ' ?
    Fuentes(9).Caracteres(8776) = 87433  ' ?
    Fuentes(9).Caracteres(8800) = 87434  ' ?
    Fuentes(9).Caracteres(8804) = 87435  ' ?
    Fuentes(9).Caracteres(8805) = 87436  ' ?
    Fuentes(9).Caracteres(8730) = 87437  ' ?
    Fuentes(9).Caracteres(8721) = 87438  ' ?
    Fuentes(9).Caracteres(916) = 87439  ' ?
    Fuentes(9).Caracteres(181) = 87440  ' µ
    Fuentes(9).Caracteres(9679) = 87441  ' ?
    
    Fuentes(10).Caracteres(65) = 87442  ' A
    Fuentes(10).Caracteres(66) = 87443  ' B
    Fuentes(10).Caracteres(67) = 87444  ' C
    Fuentes(10).Caracteres(68) = 87445  ' D
    Fuentes(10).Caracteres(69) = 87446  ' E
    Fuentes(10).Caracteres(70) = 87447  ' F
    Fuentes(10).Caracteres(71) = 87448  ' G
    Fuentes(10).Caracteres(72) = 87449  ' H
    Fuentes(10).Caracteres(73) = 87450  ' I
    Fuentes(10).Caracteres(74) = 87451  ' J
    Fuentes(10).Caracteres(75) = 87452  ' K
    Fuentes(10).Caracteres(76) = 87453  ' L
    Fuentes(10).Caracteres(77) = 87454  ' M
    Fuentes(10).Caracteres(78) = 87455  ' N
    Fuentes(10).Caracteres(79) = 87456  ' O
    Fuentes(10).Caracteres(80) = 87457  ' P
    Fuentes(10).Caracteres(81) = 87458  ' Q
    Fuentes(10).Caracteres(82) = 87459  ' R
    Fuentes(10).Caracteres(83) = 87460  ' S
    Fuentes(10).Caracteres(84) = 87461  ' T
    Fuentes(10).Caracteres(85) = 87462  ' U
    Fuentes(10).Caracteres(86) = 87463  ' V
    Fuentes(10).Caracteres(87) = 87464  ' W
    Fuentes(10).Caracteres(88) = 87465  ' X
    Fuentes(10).Caracteres(89) = 87466  ' Y
    Fuentes(10).Caracteres(90) = 87467  ' Z
    Fuentes(10).Caracteres(97) = 87468  ' a
    Fuentes(10).Caracteres(98) = 87469  ' b
    Fuentes(10).Caracteres(99) = 87470  ' c
    Fuentes(10).Caracteres(100) = 87471  ' d
    Fuentes(10).Caracteres(101) = 87472  ' e
    Fuentes(10).Caracteres(102) = 87473  ' f
    Fuentes(10).Caracteres(103) = 87474  ' g
    Fuentes(10).Caracteres(104) = 87475  ' h
    Fuentes(10).Caracteres(105) = 87476  ' i
    Fuentes(10).Caracteres(106) = 87477  ' j
    Fuentes(10).Caracteres(107) = 87478  ' k
    Fuentes(10).Caracteres(108) = 87479  ' l
    Fuentes(10).Caracteres(109) = 87480  ' m
    Fuentes(10).Caracteres(110) = 87481  ' n
    Fuentes(10).Caracteres(111) = 87482  ' o
    Fuentes(10).Caracteres(112) = 87483  ' p
    Fuentes(10).Caracteres(113) = 87484  ' q
    Fuentes(10).Caracteres(114) = 87485  ' r
    Fuentes(10).Caracteres(115) = 87486  ' s
    Fuentes(10).Caracteres(116) = 87487  ' t
    Fuentes(10).Caracteres(117) = 87488  ' u
    Fuentes(10).Caracteres(118) = 87489  ' v
    Fuentes(10).Caracteres(119) = 87490  ' w
    Fuentes(10).Caracteres(120) = 87491  ' x
    Fuentes(10).Caracteres(121) = 87492  ' y
    Fuentes(10).Caracteres(122) = 87493  ' z
    Fuentes(10).Caracteres(193) = 87494  ' Á
    Fuentes(10).Caracteres(201) = 87495  ' É
    Fuentes(10).Caracteres(205) = 87496  ' Í
    Fuentes(10).Caracteres(211) = 87497  ' Ó
    Fuentes(10).Caracteres(218) = 87498  ' Ú
    Fuentes(10).Caracteres(225) = 87499  ' á
    Fuentes(10).Caracteres(233) = 87500  ' é
    Fuentes(10).Caracteres(237) = 87501  ' í
    Fuentes(10).Caracteres(243) = 87502  ' ó
    Fuentes(10).Caracteres(250) = 87503  ' ú
    Fuentes(10).Caracteres(192) = 87504  ' À
    Fuentes(10).Caracteres(200) = 87505  ' È
    Fuentes(10).Caracteres(204) = 87506  ' Ì
    Fuentes(10).Caracteres(210) = 87507  ' Ò
    Fuentes(10).Caracteres(217) = 87508  ' Ù
    Fuentes(10).Caracteres(224) = 87509  ' à
    Fuentes(10).Caracteres(232) = 87510  ' è
    Fuentes(10).Caracteres(236) = 87511  ' ì
    Fuentes(10).Caracteres(242) = 87512  ' ò
    Fuentes(10).Caracteres(249) = 87513  ' ù
    Fuentes(10).Caracteres(194) = 87514  ' Â
    Fuentes(10).Caracteres(202) = 87515  ' Ê
    Fuentes(10).Caracteres(206) = 87516  ' Î
    Fuentes(10).Caracteres(212) = 87517  ' Ô
    Fuentes(10).Caracteres(219) = 87518  ' Û
    Fuentes(10).Caracteres(226) = 87519  ' â
    Fuentes(10).Caracteres(234) = 87520  ' ê
    Fuentes(10).Caracteres(238) = 87521  ' î
    Fuentes(10).Caracteres(244) = 87522  ' ô
    Fuentes(10).Caracteres(251) = 87523  ' û
    Fuentes(10).Caracteres(196) = 87524  ' Ä
    Fuentes(10).Caracteres(203) = 87525  ' Ë
    Fuentes(10).Caracteres(207) = 87526  ' Ï
    Fuentes(10).Caracteres(214) = 87527  ' Ö
    Fuentes(10).Caracteres(220) = 87528  ' Ü
    Fuentes(10).Caracteres(228) = 87529  ' ä
    Fuentes(10).Caracteres(235) = 87530  ' ë
    Fuentes(10).Caracteres(239) = 87531  ' ï
    Fuentes(10).Caracteres(246) = 87532  ' ö
    Fuentes(10).Caracteres(252) = 87533  ' ü
    Fuentes(10).Caracteres(195) = 87534  ' Ã
    Fuentes(10).Caracteres(213) = 87535  ' Õ
    Fuentes(10).Caracteres(227) = 87536  ' ã
    Fuentes(10).Caracteres(245) = 87537  ' õ
    Fuentes(10).Caracteres(209) = 87538  ' Ñ
    Fuentes(10).Caracteres(241) = 87539  ' ñ
    Fuentes(10).Caracteres(199) = 87540  ' Ç
    Fuentes(10).Caracteres(231) = 87541  ' ç
    Fuentes(10).Caracteres(48) = 87542  ' 0
    Fuentes(10).Caracteres(49) = 87543  ' 1
    Fuentes(10).Caracteres(50) = 87544  ' 2
    Fuentes(10).Caracteres(51) = 87545  ' 3
    Fuentes(10).Caracteres(52) = 87546  ' 4
    Fuentes(10).Caracteres(53) = 87547  ' 5
    Fuentes(10).Caracteres(54) = 87548  ' 6
    Fuentes(10).Caracteres(55) = 87549  ' 7
    Fuentes(10).Caracteres(56) = 87550  ' 8
    Fuentes(10).Caracteres(57) = 87551  ' 9
    Fuentes(10).Caracteres(161) = 87552  ' ¡
    Fuentes(10).Caracteres(33) = 87553  ' !
    Fuentes(10).Caracteres(191) = 87554  ' ¿
    Fuentes(10).Caracteres(63) = 87555  ' ?
    Fuentes(10).Caracteres(46) = 87556  ' .
    Fuentes(10).Caracteres(44) = 87557  ' ,
    Fuentes(10).Caracteres(59) = 87558  ' ;
    Fuentes(10).Caracteres(58) = 87559  ' :
    Fuentes(10).Caracteres(8230) = 87560  ' …
    Fuentes(10).Caracteres(8212) = 87561  ' —
    Fuentes(10).Caracteres(45) = 87562  ' -
    Fuentes(10).Caracteres(95) = 87563  ' _
    Fuentes(10).Caracteres(40) = 87564  ' (
    Fuentes(10).Caracteres(41) = 87565  ' )
    Fuentes(10).Caracteres(91) = 87566  ' [
    Fuentes(10).Caracteres(93) = 87567  ' ]
    Fuentes(10).Caracteres(123) = 87568  ' {
    Fuentes(10).Caracteres(125) = 87569  ' }
    Fuentes(10).Caracteres(171) = 87570  ' «
    Fuentes(10).Caracteres(187) = 87571  ' »
    Fuentes(10).Caracteres(34) = 87572  ' "
    Fuentes(10).Caracteres(39) = 87573  ' ''
    Fuentes(10).Caracteres(8216) = 87574  ' ‘
    Fuentes(10).Caracteres(8217) = 87575  ' ’
    Fuentes(10).Caracteres(8220) = 87576  ' “
    Fuentes(10).Caracteres(8221) = 87577  ' ”
    Fuentes(10).Caracteres(47) = 87578  ' /
    Fuentes(10).Caracteres(92) = 87579  ' \
    Fuentes(10).Caracteres(124) = 87580  ' |
    Fuentes(10).Caracteres(64) = 87581  ' @
    Fuentes(10).Caracteres(35) = 87582  ' #
    Fuentes(10).Caracteres(37) = 87583  ' %
    Fuentes(10).Caracteres(38) = 87584  ' &
    Fuentes(10).Caracteres(8432) = 87585  ' ?
    Fuentes(10).Caracteres(43) = 87586  ' +
    Fuentes(10).Caracteres(45) = 87587  ' -
    Fuentes(10).Caracteres(60) = 87588  ' <
    Fuentes(10).Caracteres(62) = 87589  ' >
    Fuentes(10).Caracteres(94) = 87590  ' ^
    Fuentes(10).Caracteres(95) = 87591  ' _
    Fuentes(10).Caracteres(126) = 87592  ' ~
    Fuentes(10).Caracteres(180) = 87593  ' ´
    Fuentes(10).Caracteres(172) = 87594  ' ¬
    Fuentes(10).Caracteres(36) = 87595  ' $
    Fuentes(10).Caracteres(8364) = 87596  ' €
    Fuentes(10).Caracteres(163) = 87597  ' £
    Fuentes(10).Caracteres(162) = 87598  ' ¢
    Fuentes(10).Caracteres(165) = 87599  ' ¥
    Fuentes(10).Caracteres(176) = 87600  ' °
    Fuentes(10).Caracteres(167) = 87601  ' §
    Fuentes(10).Caracteres(182) = 87602  ' ¶
    Fuentes(10).Caracteres(177) = 87603  ' ±
    Fuentes(10).Caracteres(120) = 87604  ' x
    Fuentes(10).Caracteres(247) = 87605  ' ÷
    Fuentes(10).Caracteres(8240) = 87606  ' ‰
    Fuentes(10).Caracteres(8734) = 87607  ' ?
    Fuentes(10).Caracteres(8776) = 87608  ' ?
    Fuentes(10).Caracteres(8800) = 87609  ' ?
    Fuentes(10).Caracteres(8804) = 87610  ' ?
    Fuentes(10).Caracteres(8805) = 87611  ' ?
    Fuentes(10).Caracteres(8730) = 87612  ' ?
    Fuentes(10).Caracteres(8721) = 87613  ' ?
    Fuentes(10).Caracteres(916) = 87614  ' ?
    Fuentes(10).Caracteres(181) = 87615  ' µ
    Fuentes(10).Caracteres(9679) = 87616  ' ?
    
    
    Exit Sub
Engine_Font_Initialize_Err:
    Call RegistrarError(Err.Number, Err.Description, "Graficos_Textos.Engine_Font_Initialize", Erl)
    Resume Next
End Sub

Public Function Engine_Text_Height(Texto As String, Optional multi As Boolean = False, Optional font As Byte = 1) As Integer
    On Error GoTo Engine_Text_Height_Err
    Dim A, B, c, d, e, f As Integer
    Dim graf As Grh
    If multi = False Then
        Engine_Text_Height = 0
    Else
        e = 0
        f = 0
        If font = 1 Then
            For A = 1 To Len(Texto)
                B = Asc(mid(Texto, A, 1))
                graf.GrhIndex = Fuentes(1).Caracteres(B)
                If B = 32 Or B = 13 Then
                    If e >= 20 Then 'reemplazar por lo que os plazca
                        f = f + 1
                        e = 0
                        d = 0
                    Else
                        If B = 32 Then
                            d = d + 4
                        End If
                    End If
                    'Else
                    'If graf.GrhIndex > 12 Then
                End If
                e = e + 1
            Next A
        Else
            For A = 1 To Len(Texto)
                B = Asc(mid(Texto, A, 1))
                graf.GrhIndex = Fuentes(font).Caracteres(B)
                If B = 32 Or B = 13 Then
                    If e >= 20 Then 'reemplazar por lo que os plazca
                        f = f + 1
                        e = 0
                        d = 0
                    Else
                        If B = 32 Then
                            d = d + 4
                        End If
                    End If
                    'Else
                    'If graf.GrhIndex > 12 Then
                End If
                e = e + 1
            Next A
        End If
        Engine_Text_Height = f * 14
    End If
    Exit Function
Engine_Text_Height_Err:
    Call RegistrarError(Err.Number, Err.Description, "Graficos_Textos.Engine_Text_Height", Erl)
    Resume Next
End Function

Sub Engine_Text_Render_LetraGrande(Texto As String, _
                                   x As Integer, _
                                   y As Integer, _
                                   ByRef text_color() As RGBA, _
                                   Optional ByVal font_index As Integer = 1, _
                                   Optional multi_line As Boolean = False, _
                                   Optional charindex As Integer = 0, _
                                   Optional ByVal alpha As Byte = 255)
    On Error GoTo Engine_Text_Render_LetraGrande_Err
    Dim A, B, c, d, e, f As Integer
    Dim graf          As Grh
    Dim temp_array(3) As RGBA 'Si le queres dar color a la letra pasa este parametro dsp xD
    temp_array(0) = text_color(0)
    If charindex = 0 Then
        A = 255
    Else
        A = charlist(charindex).AlphaText
    End If
    If alpha <> 255 Then
        A = alpha
    End If
    Dim i              As Long
    Dim removedDialogs As Long
    For i = 0 To dialogCount - 1
        'Decrease index to prevent jumping over a dialog
        'Crappy VB will cache the limit of the For loop, so even if it changed, it won't matter
        With dialogs(i - removedDialogs)
            If FrameTime - .startTime >= .lifeTime Then
                Call Char_Dialog_Remove(.charindex, charindex)
                If charlist(charindex).AlphaText = 0 Then
                    removedDialogs = removedDialogs + 1
                End If
            Else
            End If
        End With
    Next i
    If (Len(Texto) = 0) Then Exit Sub
    d = 0
    If multi_line = False Then
        e = 0
        f = 0
        For A = 1 To Len(Texto)
            B = Asc(mid(Texto, A, 1))
            graf.GrhIndex = Fuentes(font_index).Caracteres(B)
            If B = 32 Or B = 13 Then
                If e >= 35 Then 'reemplazar por lo que os plazca
                    f = f + 1
                    e = 0
                    d = 0
                Else
                    If B = 32 Then d = d + 30
                End If
            Else
                If graf.GrhIndex > 12 Then
                    'mega sombra O-matica
                    graf.GrhIndex = Fuentes(font_index).Caracteres(B)
                    If font_index <> 3 Then
                        'Call Draw_GrhColor(graf.GrhIndex, (x + d), y + f * 14, Sombra())
                    End If
                    Call Draw_GrhFont(graf.GrhIndex, (x + d) + 1, y + 1 + f * 14, temp_array())
                    ' graf.grhindex = Fuentes(font_index).Caracteres(b)
                    ' Grh_Render graf, (X + d), Y + f * 14, temp_array, False, False, False '14 es el height de esta fuente dsp lo hacemos dinamico
                    d = d + GrhData(GrhData(graf.GrhIndex).Frames(1)).pixelWidth - 70
                End If
            End If
            e = e + 1
        Next A
    Else
        e = 0
        f = 0
        For A = 1 To Len(Texto)
            B = Asc(mid(Texto, A, 1))
            graf.GrhIndex = Fuentes(font_index).Caracteres(B)
            If B = 32 Or B = 13 Then
                If e >= 33 Then 'reemplazar por lo que os plazca
                    f = f + 1
                    e = 0
                    d = 0
                Else
                    If B = 32 Then d = d + 2
                End If
            Else
                If graf.GrhIndex > 12 Then
                    'mega sombra O-matica
                    graf.GrhIndex = Fuentes(font_index).Caracteres(B)
                    ' Call Draw_GrhColor(graf.GrhIndex, (x + d) + 1, y + 1 + f * 14, Sombra())
                    Call Draw_GrhFont(graf.GrhIndex, (x + d), y + f * 14, temp_array())
                    ' graf.grhindex = Fuentes(font_index).Caracteres(b)
                    'Grh_Render graf, (x + d), y + f * 14, temp_array, False, False, False '14 es el height de esta fuente dsp lo hacemos dinamico
                    If font_index = 5 Then
                        d = d + GrhData(GrhData(graf.GrhIndex).Frames(1)).pixelWidth - 50
                    Else
                        d = d + GrhData(GrhData(graf.GrhIndex).Frames(1)).pixelWidth
                    End If
                End If
            End If
            e = e + 1
        Next A
    End If
    Exit Sub
Engine_Text_Render_LetraGrande_Err:
    Call RegistrarError(Err.Number, Err.Description, "Graficos_Textos.Engine_Text_Render_LetraGrande", Erl)
    Resume Next
End Sub

Public Sub RenderText(ByVal Texto As String, _
                      ByVal x As Integer, _
                      ByVal y As Integer, _
                      ByRef text_color() As RGBA, _
                      Optional ByVal font_index As Integer = 1, _
                      Optional multi_line As Boolean = False, _
                      Optional charindex As Integer = 0, _
                      Optional ByVal alpha As Byte = 255)
    On Error GoTo RenderText_Err
    If (Len(Texto) = 0) Then Exit Sub
    Dim A, B, c, d, e, f As Integer
    Dim graf As Grh
    If charindex = 0 Then
        A = 255
    Else
        A = charlist(charindex).AlphaText
    End If
    If alpha <> 255 Then
        A = alpha
    End If
    Dim i              As Long
    Dim removedDialogs As Long
    For i = 0 To dialogCount - 1
        'Decrease index to prevent jumping over a dialog
        'Crappy VB will cache the limit of the For loop, so even if it changed, it won't matter
        With dialogs(i - removedDialogs)
            If FrameTime - .startTime >= .lifeTime Then
                Call Char_Dialog_Remove(.charindex, charindex)
                If A <= 0 Then
                    removedDialogs = removedDialogs + 1
                End If
            End If
        End With
    Next i
    d = 0
    If multi_line = False Then
        e = 0
        f = 0
        For A = 1 To Len(Texto)
            B = Asc(mid(Texto, A, 1))
            graf.GrhIndex = Fuentes(font_index).Caracteres(B)
            If B = 32 Or B = 13 Then
                If e >= 30 Then 'reemplazar por lo que os plazca
                    f = f + 1
                    e = 0
                    d = 0
                Else
                    If B = 32 Then d = d + 2
                End If
            Else
                If graf.GrhIndex > 12 Then
                    graf.GrhIndex = Fuentes(font_index).Caracteres(B)
                    Call Draw_GrhFont(graf.GrhIndex, (x + d) + 1, y + 1 + f * 14, text_color())
                    d = d + GrhData(GrhData(graf.GrhIndex).Frames(1)).pixelWidth
                End If
            End If
            e = e + 1
        Next A
    Else
        e = 0
        f = 0
        For A = 1 To Len(Texto)
            B = Asc(mid(Texto, A, 1))
            graf.GrhIndex = Fuentes(font_index).Caracteres(B)
            If B = 32 Or B = 13 Then
                If e >= 33 Then 'reemplazar por lo que os plazca
                    f = f + 1
                    e = 0
                    d = 0
                Else
                    If B = 32 Then d = d + 2
                End If
            Else
                If graf.GrhIndex > 12 Then
                    graf.GrhIndex = Fuentes(font_index).Caracteres(B)
                    Call Draw_GrhFont(graf.GrhIndex, (x + d), y + f * 14, text_color())
                    If font_index = 4 Then
                        d = d + GrhData(GrhData(graf.GrhIndex).Frames(1)).pixelWidth - 1
                    Else
                        d = d + GrhData(GrhData(graf.GrhIndex).Frames(1)).pixelWidth
                    End If
                End If
            End If
            e = e + 1
        Next A
    End If
    Exit Sub
RenderText_Err:
    Call RegistrarError(Err.Number, Err.Description, "Graficos_Textos.RenderText", Erl)
    Resume Next
End Sub

Public Sub Engine_Text_Render(Texto As String, _
                              ByVal x As Integer, _
                              ByVal y As Integer, _
                              ByRef text_color() As RGBA, _
                              Optional ByVal font_index As Integer = 1, _
                              Optional multi_line As Boolean = False, _
                              Optional charindex As Integer = 0, _
                              Optional ByVal alpha As Byte = 255)
    On Error GoTo Engine_Text_Render_Err
    Dim A, B, c, d, e, f As Integer
    Dim graf          As Grh
    Dim temp_array(3) As RGBA
    If charindex = 0 Then
        A = 255
    Else
        A = Clamp(charlist(charindex).AlphaText, 0, 255)
    End If
    If alpha <> 255 Then
        A = alpha
    End If
    Call RGBAList(temp_array, text_color(0).R, text_color(0).G, text_color(0).B, A)
    Dim i              As Long
    Dim removedDialogs As Long
    For i = 0 To dialogCount - 1
        'Decrease index to prevent jumping over a dialog
        'Crappy VB will cache the limit of the For loop, so even if it changed, it won't matter
        With dialogs(i - removedDialogs)
            If FrameTime - .startTime >= .lifeTime Then
                Call Char_Dialog_Remove(.charindex, charindex)
                If A <= 0 Then
                    removedDialogs = removedDialogs + 1
                End If
            Else
            End If
        End With
    Next i
    Dim Sombra(3) As RGBA 'Sombra
    Call RGBAList(Sombra, text_color(0).R / 6, text_color(0).G / 6, text_color(0).B / 6, 0.8 * A)
    If (Len(Texto) = 0) Then Exit Sub
    d = 0
    If multi_line = False Then
        e = 0
        f = 0
        For A = 1 To Len(Texto)
            B = Asc(mid(Texto, A, 1))
            graf.GrhIndex = Fuentes(font_index).Caracteres(B)
            If B = 32 Or B = 13 Then
                If e >= 32 Then 'reemplazar por lo que os plazca
                    f = f + 1
                    e = 0
                    d = 0
                Else
                    If B = 32 Then d = d + 4
                End If
            Else
                If graf.GrhIndex > 12 Then
                    'mega sombra O-matica
                    graf.GrhIndex = Fuentes(font_index).Caracteres(B)
                    If font_index <> 3 Then
                        Call Draw_GrhFont(graf.GrhIndex, (x + d) + 1, y + 1 + f * 14, Sombra())
                    End If
                    Call Draw_GrhFont(graf.GrhIndex, (x + d), y + f * 14, temp_array())
                    ' graf.grhindex = Fuentes(font_index).Caracteres(b)
                    ' Grh_Render graf, (X + d), Y + f * 14, temp_array, False, False, False '14 es el height de esta fuente dsp lo hacemos dinamico
                    d = d + GrhData(GrhData(graf.GrhIndex).Frames(1)).pixelWidth
                End If
            End If
            e = e + 1
        Next A
    Else
        e = 0
        f = 0
        For A = 1 To Len(Texto)
            B = Asc(mid(Texto, A, 1))
            graf.GrhIndex = Fuentes(font_index).Caracteres(B)
            If B = 32 Or B = 13 Then
                If e >= 20 Then 'reemplazar por lo que os plazca
                    f = f + 1
                    e = 0
                    d = 0
                Else
                    If B = 32 Then d = d + 4
                End If
            Else
                If graf.GrhIndex > 12 Then
                    'mega sombra O-matica
                    graf.GrhIndex = Fuentes(font_index).Caracteres(B)
                    Call Draw_GrhFont(graf.GrhIndex, (x + d) + 1, y + 1 + f * 14, Sombra())
                    Call Draw_GrhFont(graf.GrhIndex, (x + d), y + f * 14, temp_array())
                    ' graf.grhindex = Fuentes(font_index).Caracteres(b)
                    'Grh_Render graf, (x + d), y + f * 14, temp_array, False, False, False '14 es el height de esta fuente dsp lo hacemos dinamico
                    If font_index = 4 Then
                        d = d + GrhData(GrhData(graf.GrhIndex).Frames(1)).pixelWidth - 1
                    Else
                        d = d + GrhData(GrhData(graf.GrhIndex).Frames(1)).pixelWidth
                    End If
                End If
            End If
            e = e + 1
        Next A
    End If
    Exit Sub
Engine_Text_Render_Err:
    Call RegistrarError(Err.Number, Err.Description, "Graficos_Textos.Engine_Text_Render", Erl)
    Resume Next
End Sub

Public Sub simple_text_render(Texto As String, _
                              ByVal x As Integer, _
                              ByVal y As Integer, _
                              ByRef text_color() As RGBA, _
                              Optional ByVal font_index As Integer = 1, _
                              Optional multi_line As Boolean = False, _
                              Optional charindex As Integer = 0, _
                              Optional ByVal alpha As Byte = 255)
    On Error GoTo Engine_Text_Render_Err
    Dim A, B, c, d, e, f As Integer
    Dim graf          As Grh
    Dim temp_array(3) As RGBA
    If charindex = 0 Then
        A = 255
    Else
        A = Clamp(charlist(charindex).AlphaText, 0, 255)
    End If
    If alpha <> 255 Then
        A = alpha
    End If
    Call RGBAList(temp_array, text_color(0).R, text_color(0).G, text_color(0).B, A)
    Dim i         As Long
    Dim Sombra(3) As RGBA 'Sombra
    Call RGBAList(Sombra, text_color(0).R / 6, text_color(0).G / 6, text_color(0).B / 6, 0.8 * A)
    If (Len(Texto) = 0) Then Exit Sub
    d = 0
    f = 0
    For A = 1 To Len(Texto)
        B = Asc(mid(Texto, A, 1))
        graf.GrhIndex = Fuentes(font_index).Caracteres(B)
        If graf.GrhIndex > 12 Then
            'mega sombra O-matica
            graf.GrhIndex = Fuentes(font_index).Caracteres(B)
            Call Draw_GrhFont(graf.GrhIndex, (x + d) + 1, y + 1 + f * 14, Sombra())
            Call Draw_GrhFont(graf.GrhIndex, (x + d), y + f * 14, temp_array())
            ' graf.grhindex = Fuentes(font_index).Caracteres(b)
            If font_index = 4 Then
                d = d + GrhData(GrhData(graf.GrhIndex).Frames(1)).pixelWidth - 1
            Else
                d = d + GrhData(GrhData(graf.GrhIndex).Frames(1)).pixelWidth
            End If
        End If
    Next A
    Exit Sub
Engine_Text_Render_Err:
    Call RegistrarError(Err.Number, Err.Description, "Graficos_Textos.Engine_Text_Render", Erl)
    Resume Next
End Sub

Public Sub Engine_Text_Render_No_Ladder(Texto As String, _
                                        ByVal x As Integer, _
                                        ByVal y As Integer, _
                                        ByRef text_color() As RGBA, _
                                        ByVal status As Byte, _
                                        Optional ByVal font_index As Integer = 1, _
                                        Optional multi_line As Boolean = False, _
                                        Optional charindex As Integer = 0, _
                                        Optional ByVal alpha As Byte = 255)
    On Error GoTo Engine_Text_Render_Err
    Dim A         As Integer, B As Integer, c As Integer, d As Integer
    Dim graf      As Grh
    Dim color1(3) As RGBA
    Dim color2(3) As RGBA
    If charindex = 0 Then
        A = 255
    Else
        A = Clamp(charlist(charindex).AlphaText, 0, 255)
    End If
    If alpha <> 255 Then
        A = alpha
    End If
    Select Case status
        Case 0 'criminal
            Call RGBAList(color1, 225, 0, 0, A)
            Call RGBAList(color2, 255, 255, 255, A)
        Case 1 'ciuda
            Call RGBAList(color1, 0, 128, 255, A)
            Call RGBAList(color2, 255, 255, 255, A)
        Case 2 'legión oscura
            Call RGBAList(color1, 155, 0, 0, A)
            Call RGBAList(color2, 255, 255, 255, A)
        Case 3 'armada real
            Call RGBAList(color1, 0, 175, 255, A)
            Call RGBAList(color2, 255, 255, 255, A)
        Case 4 'Legión
            Call RGBAList(color1, 155, 0, 0, A)
            Call RGBAList(color2, 255, 255, 255, A)
        Case 5 'Consejo
            Call RGBAList(color1, 22, 239, 253, A)
            Call RGBAList(color2, 255, 255, 255, A)
        Case 7 'aviso solicitud
            Call RGBAList(color2, 255, 255, 0, A)
        Case 8 'aviso desconectado
            Call RGBAList(color2, 255, 0, 0, A)
        Case 9 'aviso conectado
            Call RGBAList(color2, 10, 182, 70, A)
        Case 10 'lider
            Call RGBAList(color1, 222, 194, 112, A)
            Call RGBAList(color2, 255, 255, 255, A)
    End Select
    'Call RGBAList(color2, 255, 255, 255, A)
    Dim i              As Long
    Dim removedDialogs As Long
    For i = 0 To dialogCount - 1
        'Decrease index to prevent jumping over a dialog
        'Crappy VB will cache the limit of the For loop, so even if it changed, it won't matter
        With dialogs(i - removedDialogs)
            If FrameTime - .startTime >= .lifeTime Then
                Call Char_Dialog_Remove(.charindex, charindex)
                If A <= 0 Then
                    removedDialogs = removedDialogs + 1
                End If
            Else
            End If
        End With
    Next i
    Dim Sombra(3) As RGBA 'Sombra
    Call RGBAList(Sombra, text_color(0).R / 6, text_color(0).G / 6, text_color(0).B / 6, 0.8 * A)
    If (Len(Texto) = 0) Then Exit Sub
    Dim row As Integer, charPos As Integer
    d = 0
    row = 0
    charPos = 0
    Dim separador As Boolean
    For A = 1 To Len(Texto)
        B = Asc(mid(Texto, A, 1))
        graf.GrhIndex = Fuentes(font_index).Caracteres(B)
        If B = 1 Then separador = Not separador
        If graf.GrhIndex > 12 Then
            'mega sombra O-matica
            graf.GrhIndex = Fuentes(font_index).Caracteres(B)
            If font_index <> 3 Then
                Call Draw_GrhFont(graf.GrhIndex, (x + d) + 1, y + 1 + 10, Sombra())
            End If
            If status >= 0 And status <= 5 Or status = 10 Then
                If separador Then
                    Call Draw_GrhFont(graf.GrhIndex, (x + d), y + 10, color1)
                Else
                    Call Draw_GrhFont(graf.GrhIndex, (x + d), y + 10, color2)
                End If
            Else
                Call Draw_GrhFont(graf.GrhIndex, (x + d), y + 10, color2)
            End If
            d = d + GrhData(GrhData(graf.GrhIndex).Frames(1)).pixelWidth
        End If
        charPos = charPos + 1
    Next A
    Exit Sub
Engine_Text_Render_Err:
    Call RegistrarError(Err.Number, Err.Description, "Graficos_Textos.Engine_Text_Render", Erl)
    Resume Next
End Sub

Public Sub Engine_Text_RenderGrande(Texto As String, _
                                    x As Integer, _
                                    y As Integer, _
                                    ByRef text_color() As RGBA, _
                                    Optional ByVal font_index As Integer = 1, _
                                    Optional multi_line As Boolean = False, _
                                    Optional charindex As Integer = 0, _
                                    Optional ByVal alpha As Byte = 255)
    On Error GoTo Engine_Text_RenderGrande_Err
    Dim A, B, c, d, e, f As Integer
    Dim graf          As Grh
    Dim temp_array(3) As RGBA
    If charindex = 0 Then
        A = 255
    Else
        A = charlist(charindex).AlphaText
    End If
    If alpha <> 255 Then
        A = alpha
    End If
    Call RGBAList(temp_array, text_color(0).R, text_color(0).G, text_color(0).B, A)
    Dim i              As Long
    Dim removedDialogs As Long
    For i = 0 To dialogCount - 1
        'Decrease index to prevent jumping over a dialog
        'Crappy VB will cache the limit of the For loop, so even if it changed, it won't matter
        With dialogs(i - removedDialogs)
            If FrameTime - .startTime >= .lifeTime Then
                Call Char_Dialog_Remove(.charindex, charindex)
                If A <= 0 Then
                    removedDialogs = removedDialogs + 1
                End If
            Else
            End If
        End With
    Next i
    Dim Sombra(3) As RGBA 'Sombra
    Call RGBAList(Sombra, text_color(0).R / 6, text_color(0).G / 6, text_color(0).B / 6, 0.8 * alpha)
    If (Len(Texto) = 0) Then Exit Sub
    d = 0
    If multi_line = False Then
        e = 0
        f = 0
        For A = 1 To Len(Texto)
            B = Asc(mid(Texto, A, 1))
            graf.GrhIndex = Fuentes(font_index).Caracteres(B)
            If B = 32 Or B = 13 Then
                If e >= 35 Then 'reemplazar por lo que os plazca
                    f = f + 1
                    e = 0
                    d = 0
                Else
                    If B = 32 Then d = d + 12
                End If
            Else
                If graf.GrhIndex > 12 Then
                    'mega sombra O-matica
                    graf.GrhIndex = Fuentes(font_index).Caracteres(B)
                    If font_index <> 3 Then
                        Call Draw_GrhFont(graf.GrhIndex, (x + d), y + f * 14, Sombra())
                    End If
                    Call Draw_GrhFont(graf.GrhIndex, (x + d) + 1, y + 1 + f * 14, temp_array())
                    ' graf.grhindex = Fuentes(font_index).Caracteres(b)
                    ' Grh_Render graf, (X + d), Y + f * 14, temp_array, False, False, False '14 es el height de esta fuente dsp lo hacemos dinamico
                    d = d + GrhData(GrhData(graf.GrhIndex).Frames(1)).pixelWidth
                End If
            End If
            e = e + 1
        Next A
    Else
        e = 0
        f = 0
        For A = 1 To Len(Texto)
            B = Asc(mid(Texto, A, 1))
            graf.GrhIndex = Fuentes(font_index).Caracteres(B)
            If B = 32 Or B = 13 Then
                If e >= 10 Then 'reemplazar por lo que os plazca
                    f = f + 3
                    e = 0
                    d = 0
                Else
                    If B = 32 Then d = d + 12
                End If
            Else
                If graf.GrhIndex > 12 Then
                    'mega sombra O-matica
                    graf.GrhIndex = Fuentes(font_index).Caracteres(B)
                    'Call Draw_GrhColor(graf.GrhIndex, (x + d) + 1, y + 1 + f * 14, Sombra())
                    Call Draw_GrhFont(graf.GrhIndex, (x + d), y + f * 14, temp_array())
                    ' graf.grhindex = Fuentes(font_index).Caracteres(b)
                    'Grh_Render graf, (x + d), y + f * 14, temp_array, False, False, False '14 es el height de esta fuente dsp lo hacemos dinamico
                    If font_index = 4 Then
                        d = d + GrhData(GrhData(graf.GrhIndex).Frames(1)).pixelWidth
                    Else
                        d = d + GrhData(GrhData(graf.GrhIndex).Frames(1)).pixelWidth
                    End If
                End If
            End If
            e = e + 1
        Next A
    End If
    Exit Sub
Engine_Text_RenderGrande_Err:
    Call RegistrarError(Err.Number, Err.Description, "Graficos_Textos.Engine_Text_RenderGrande", Erl)
    Resume Next
End Sub

Public Sub Engine_Text_Render2(Texto As String, _
                               x As Integer, _
                               y As Integer, _
                               ByRef text_color As RGBA, _
                               Optional ByVal font_index As Integer = 1, _
                               Optional multi_line As Boolean = False, _
                               Optional charindex As Long = 0, _
                               Optional ByVal alpha As Boolean = False)
    On Error GoTo Engine_Text_Render2_Err
    Dim A, B, c, d, e, f As Integer
    Dim graf          As Grh
    Dim temp_array(3) As RGBA
    Call RGBAList(temp_array, text_color.R, text_color.G, text_color.B, text_color.A)
    Dim Sombra(3) As RGBA 'Sombra
    Call RGBAList(Sombra, text_color.R / 6, text_color.G / 6, text_color.B / 6, 0.8 * text_color.A)
    If (Len(Texto) = 0) Then Exit Sub
    d = 0
    If multi_line = False Then
        e = 0
        f = 0
        For A = 1 To Len(Texto)
            B = Asc(mid(Texto, A, 1))
            graf.GrhIndex = Fuentes(font_index).Caracteres(B)
            If B = 32 Or B = 13 Then
                If e >= 35 Then 'reemplazar por lo que os plazca
                    f = f + 1
                    e = 0
                    d = 0
                Else
                    If B = 32 Then d = d + 4
                End If
            Else
                If graf.GrhIndex > 12 Then
                    'mega sombra O-matica
                    graf.GrhIndex = Fuentes(font_index).Caracteres(B)
                    If font_index <> 3 Then
                        Call Draw_GrhFont(graf.GrhIndex, (x + d) + 1, y + 1 + f * 14, Sombra())
                    End If
                    Call Draw_GrhFont(graf.GrhIndex, (x + d), y + f * 14, temp_array())
                    ' graf.grhindex = Fuentes(font_index).Caracteres(b)
                    ' Grh_Render graf, (X + d), Y + f * 14, temp_array, False, False, False '14 es el height de esta fuente dsp lo hacemos dinamico
                    d = d + GrhData(GrhData(graf.GrhIndex).Frames(1)).pixelWidth
                End If
            End If
            e = e + 1
        Next A
    Else
        e = 0
        f = 0
        For A = 1 To Len(Texto)
            B = Asc(mid(Texto, A, 1))
            graf.GrhIndex = Fuentes(font_index).Caracteres(B)
            If B = 32 Or B = 13 Then
                If e >= 20 Then 'reemplazar por lo que os plazca
                    f = f + 1
                    e = 0
                    d = 0
                Else
                    If B = 32 Then d = d + 4
                End If
            Else
                If graf.GrhIndex > 12 Then
                    'mega sombra O-matica
                    graf.GrhIndex = Fuentes(font_index).Caracteres(B)
                    Call Draw_GrhFont(graf.GrhIndex, (x + d) + 1, y + 1 + f * 14, Sombra())
                    Call Draw_GrhFont(graf.GrhIndex, (x + d), y + f * 14, temp_array())
                    ' graf.grhindex = Fuentes(font_index).Caracteres(b)
                    'Grh_Render graf, (x + d), y + f * 14, temp_array, False, False, False '14 es el height de esta fuente dsp lo hacemos dinamico
                    If font_index <> 3 Then
                        d = d + GrhData(GrhData(graf.GrhIndex).Frames(1)).pixelWidth
                    Else
                        d = d + GrhData(GrhData(graf.GrhIndex).Frames(1)).pixelWidth
                    End If
                End If
            End If
            e = e + 1
        Next A
    End If
    Exit Sub
Engine_Text_Render2_Err:
    Call RegistrarError(Err.Number, Err.Description, "Graficos_Textos.Engine_Text_Render2", Erl)
    Resume Next
End Sub

Public Sub Engine_Text_Render_Efect(charindex As Integer, _
                                    Texto As String, _
                                    x As Integer, _
                                    y As Integer, _
                                    ByRef text_color() As RGBA, _
                                    Optional ByVal font_index As Integer = 1, _
                                    Optional multi_line As Boolean = False)
    On Error GoTo Engine_Text_Render_Efect_Err
    Dim A, B, c, d, e, f As Integer
    Dim graf As Grh
    If (Len(Texto) = 0) Then Exit Sub
    d = 0
    e = 0
    f = 0
    Dim Sombra(3) As RGBA 'Sombra
    Call RGBAList(Sombra, text_color(0).R / 6, text_color(0).G / 6, text_color(0).B / 6, 0.8 * text_color(0).A)
    For A = 1 To Len(Texto)
        B = Asc(mid(Texto, A, 1))
        graf.GrhIndex = Fuentes(font_index).Caracteres(B)
        If B = 32 Or B = 13 Then
            If e >= 20 Then 'reemplazar por lo que os plazca
                f = f + 1
                e = 0
                d = 0
            Else
                If B = 32 Then d = d + 4
            End If
        Else
            If graf.GrhIndex > 12 Then
                'mega sombra O-matica
                graf.GrhIndex = Fuentes(font_index).Caracteres(B)
                Call Draw_GrhFont(graf.GrhIndex, (x + d) + 1, y + 1 + f * 14, Sombra())
                Call Draw_GrhFont(graf.GrhIndex, (x + d), y + f * 14, text_color())
                ' graf.grhindex = Fuentes(font_index).Caracteres(b)
                'Grh_Render graf, (x + d), y + f * 14, temp_array, False, False, False '14 es el height de esta fuente dsp lo hacemos dinamico
                d = d + GrhData(GrhData(graf.GrhIndex).Frames(1)).pixelWidth
            End If
        End If
        e = e + 1
    Next A
    Exit Sub
Engine_Text_Render_Efect_Err:
    Call RegistrarError(Err.Number, Err.Description, "Graficos_Textos.Engine_Text_Render_Efect", Erl)
    Resume Next
End Sub

Public Function Engine_Text_Width(Texto As String, Optional multi As Boolean = False, Optional Fon As Byte = 1) As Integer
    On Error GoTo Engine_Text_Width_Err
    Dim A, B, d, e, f As Integer
    Dim graf As Grh
    Select Case Fon
        Case 1
            If multi = False Then
                For A = 1 To Len(Texto)
                    B = Asc(mid(Texto, A, 1))
                    graf.GrhIndex = Fuentes(1).Caracteres(B)
                    If graf.GrhIndex = 0 Then graf.GrhIndex = 1
                    If B <> 32 Then
                        Engine_Text_Width = Engine_Text_Width + GrhData(GrhData(graf.GrhIndex).Frames(1)).pixelWidth '+ 1
                    Else
                        Engine_Text_Width = Engine_Text_Width + 4
                    End If
                Next A
            Else
                e = 0
                f = 0
                For A = 1 To Len(Texto)
                    B = Asc(mid(Texto, A, 1))
                    graf.GrhIndex = Fuentes(1).Caracteres(B)
                    If B = 32 Or B = 13 Then
                        If e >= 20 Then 'reemplazar por lo que os plazca
                            f = f + 1
                            e = 0
                            d = 0
                        Else
                            If B = 32 Then d = d + 4
                        End If
                    Else
                        If graf.GrhIndex > 12 Then
                            d = d + GrhData(GrhData(graf.GrhIndex).Frames(1)).pixelWidth '+ 1
                            If d > Engine_Text_Width Then Engine_Text_Width = d
                        End If
                    End If
                    e = e + 1
                Next A
            End If
        Case 4
            If multi = False Then
                For A = 1 To Len(Texto)
                    B = Asc(mid(Texto, A, 1))
                    graf.GrhIndex = Fuentes(Fon).Caracteres(B)
                    If graf.GrhIndex = 0 Then graf.GrhIndex = 1
                    If B <> 20 Then
                        Engine_Text_Width = Engine_Text_Width + GrhData(GrhData(graf.GrhIndex + 1).Frames(1)).pixelWidth + 10
                    Else
                        Engine_Text_Width = Engine_Text_Width - 15
                    End If
                Next A
            Else
                e = 0
                f = 0
                For A = 1 To Len(Texto)
                    B = Asc(mid(Texto, A, 1))
                    graf.GrhIndex = Fuentes(Fon).Caracteres(B)
                    If B = 32 Or B = 13 Then
                        If e >= 20 Then 'reemplazar por lo que os plazca
                            f = f + 1
                            e = 0
                            d = 0
                        Else
                            If B = 32 Then d = d + 4
                        End If
                    Else
                        If graf.GrhIndex > 12 Then
                            d = d + GrhData(GrhData(graf.GrhIndex).Frames(1)).pixelWidth '+ 1
                            If d > Engine_Text_Width Then Engine_Text_Width = d
                        End If
                    End If
                    e = e + 1
                Next A
            End If
        Case Else
            If multi = False Then
                For A = 1 To Len(Texto)
                    B = Asc(mid(Texto, A, 1))
                    graf.GrhIndex = Fuentes(Fon).Caracteres(B)
                    If graf.GrhIndex = 0 Then graf.GrhIndex = 1
                    If B <> 32 Then
                        Engine_Text_Width = Engine_Text_Width + GrhData(GrhData(graf.GrhIndex).Frames(1)).pixelWidth '+ 1
                    Else
                        Engine_Text_Width = Engine_Text_Width + 4
                    End If
                Next A
            Else
                e = 0
                f = 0
                For A = 1 To Len(Texto)
                    B = Asc(mid(Texto, A, 1))
                    graf.GrhIndex = Fuentes(Fon).Caracteres(B)
                    If B = 32 Or B = 13 Then
                        If e >= 20 Then 'reemplazar por lo que os plazca
                            f = f + 1
                            e = 0
                            d = 0
                        Else
                            If B = 32 Then d = d + 4
                        End If
                    Else
                        If graf.GrhIndex > 12 Then
                            d = d + GrhData(GrhData(graf.GrhIndex).Frames(1)).pixelWidth '+ 1
                            If d > Engine_Text_Width Then Engine_Text_Width = d
                        End If
                    End If
                    e = e + 1
                Next A
            End If
    End Select
    Exit Function
Engine_Text_Width_Err:
    Call RegistrarError(Err.Number, Err.Description, "Graficos_Textos.Engine_Text_Width", Erl)
    Resume Next
End Function

Public Function Engine_Text_WidthCentrado(Texto As String, Optional multi As Boolean = False, Optional Fon As Byte = 1) As Integer
    On Error GoTo Engine_Text_WidthCentrado_Err
    Dim A, B, d, e, f As Integer
    Dim graf As Grh
    Select Case Fon
        Case 1
            '
            If multi = False Then
                For A = 1 To Len(Texto)
                    B = Asc(mid(Texto, A, 1))
                    graf.GrhIndex = Fuentes(1).Caracteres(B)
                    If graf.GrhIndex = 0 Then graf.GrhIndex = 1
                    If B <> 32 Then
                        Engine_Text_WidthCentrado = Engine_Text_WidthCentrado + GrhData(GrhData(graf.GrhIndex + 1).Frames(1)).pixelWidth '+ 1
                    Else
                        Engine_Text_WidthCentrado = Engine_Text_WidthCentrado + 4
                    End If
                Next A
            Else
                e = 0
                f = 0
                For A = 1 To Len(Texto)
                    B = Asc(mid(Texto, A, 1))
                    graf.GrhIndex = Fuentes(1).Caracteres(B)
                    If B = 32 Or B = 13 Then
                        If e >= 20 Then 'reemplazar por lo que os plazca
                            f = f + 1
                            e = 0
                            d = 0
                        Else
                            If B = 32 Then d = d + 4
                        End If
                    Else
                        If graf.GrhIndex > 12 Then
                            d = d + GrhData(GrhData(graf.GrhIndex).Frames(1)).pixelWidth '+ 1
                            If d > Engine_Text_WidthCentrado Then Engine_Text_WidthCentrado = d
                        End If
                    End If
                    e = e + 1
                Next A
            End If
        Case 4
            If multi = False Then
                For A = 1 To Len(Texto)
                    B = Asc(mid(Texto, A, 1))
                    graf.GrhIndex = Fuentes(Fon).Caracteres(B)
                    If graf.GrhIndex = 0 Then graf.GrhIndex = 1
                    If B <> 20 Then
                        Engine_Text_WidthCentrado = Engine_Text_WidthCentrado + GrhData(GrhData(graf.GrhIndex + 1).Frames(1)).pixelWidth + 10
                    Else
                        Engine_Text_WidthCentrado = Engine_Text_WidthCentrado - 15
                    End If
                Next A
            Else
                e = 0
                f = 0
                For A = 1 To Len(Texto)
                    B = Asc(mid(Texto, A, 1))
                    graf.GrhIndex = Fuentes(Fon).Caracteres(B)
                    If B = 32 Or B = 13 Then
                        If e >= 20 Then 'reemplazar por lo que os plazca
                            f = f + 1
                            e = 0
                            d = 0
                        Else
                            If B = 32 Then d = d + 4
                        End If
                    Else
                        If graf.GrhIndex > 12 Then
                            d = d + GrhData(GrhData(graf.GrhIndex).Frames(1)).pixelWidth '+ 1
                            If d > Engine_Text_WidthCentrado Then Engine_Text_WidthCentrado = d
                        End If
                    End If
                    e = e + 1
                Next A
            End If
        Case Else
            If multi = False Then
                For A = 1 To Len(Texto)
                    B = Asc(mid(Texto, A, 1))
                    graf.GrhIndex = Fuentes(Fon).Caracteres(B)
                    If graf.GrhIndex = 0 Then graf.GrhIndex = 1
                    If B <> 32 Then
                        Engine_Text_WidthCentrado = Engine_Text_WidthCentrado + GrhData(GrhData(graf.GrhIndex).Frames(1)).pixelWidth '+ 1
                    Else
                        Engine_Text_WidthCentrado = Engine_Text_WidthCentrado + 4
                    End If
                Next A
            Else
                e = 0
                f = 0
                For A = 1 To Len(Texto)
                    B = Asc(mid(Texto, A, 1))
                    graf.GrhIndex = Fuentes(Fon).Caracteres(B)
                    If B = 32 Or B = 13 Then
                        If e >= 20 Then 'reemplazar por lo que os plazca
                            f = f + 1
                            e = 0
                            d = 0
                        Else
                            If B = 32 Then d = d + 4
                        End If
                    Else
                        If graf.GrhIndex > 12 Then
                            d = d + GrhData(GrhData(graf.GrhIndex).Frames(1)).pixelWidth '+ 1
                            If d > Engine_Text_WidthCentrado Then Engine_Text_WidthCentrado = d
                        End If
                    End If
                    e = e + 1
                Next A
            End If
    End Select
    Exit Function
Engine_Text_WidthCentrado_Err:
    Call RegistrarError(Err.Number, Err.Description, "Graficos_Textos.Engine_Text_WidthCentrado", Erl)
    Resume Next
End Function

Public Sub Text_Render(ByVal font As D3DXFont, _
                       text As String, _
                       ByVal Top As Long, _
                       ByVal Left As Long, _
                       ByVal Width As Long, _
                       ByVal Height As Long, _
                       ByVal color As Long, _
                       ByVal format As Long, _
                       Optional ByVal Shadow As Boolean = False)
    On Error GoTo Text_Render_Err
    Dim TextRect   As Rect
    Dim ShadowRect As Rect
    TextRect.Top = Top
    TextRect.Left = Left
    TextRect.Bottom = Top + Height
    TextRect.Right = Left + Width
    If Shadow Then
        ShadowRect.Top = Top - 1
        ShadowRect.Left = Left - 2
        ShadowRect.Bottom = (Top + Height) - 1
        ShadowRect.Right = (Left + Width) - 2
        DirectD3D8.drawText font, &HFF000000, text, ShadowRect, format
    End If
    DirectD3D8.drawText font, color, text, TextRect, format
    Exit Sub
Text_Render_Err:
    Call RegistrarError(Err.Number, Err.Description, "Graficos_Textos.Text_Render", Erl)
    Resume Next
End Sub

Public Sub Text_Render_ext(text As String, _
                           ByVal Top As Long, _
                           ByVal Left As Long, _
                           ByVal Width As Long, _
                           ByVal Height As Long, _
                           ByVal color As Long, _
                           Optional ByVal Shadow As Boolean = False, _
                           Optional ByVal center As Boolean = False, _
                           Optional ByVal font As Long = 0)
    On Error GoTo Text_Render_ext_Err
    If center = True Then
        Call Text_Render(font_list(font), text, Top, Left, Width, Height, color, DT_VCENTER & DT_CENTER, Shadow)
    Else
        Call Text_Render(font_list(font), text, Top, Left, Width, Height, color, DT_TOP Or DT_LEFT, Shadow)
    End If
    Exit Sub
Text_Render_ext_Err:
    Call RegistrarError(Err.Number, Err.Description, "Graficos_Textos.Text_Render_ext", Erl)
    Resume Next
End Sub

Private Sub Font_Make(ByVal font_index As Long, ByVal Style As String, ByVal bold As Boolean, ByVal italic As Boolean, ByVal size As Long)
    On Error GoTo Font_Make_Err
    If font_index > font_last Then
        font_last = font_index
        ReDim Preserve font_list(1 To font_last)
    End If
    font_count = font_count + 1
    Dim font_desc As IFont
    Dim fnt       As New StdFont
    fnt.Name = Style
    fnt.size = size
    fnt.bold = bold
    fnt.italic = italic
    Set font_desc = fnt
    Set font_list(font_index) = DirectD3D8.CreateFont(DirectDevice, font_desc.hFont)
    Exit Sub
Font_Make_Err:
    Call RegistrarError(Err.Number, Err.Description, "Graficos_Textos.Font_Make", Erl)
    Resume Next
End Sub

Public Function Font_Create(ByVal Style As String, ByVal size As Long, ByVal bold As Boolean, ByVal italic As Boolean) As Long
    On Error GoTo ErrorHandler:
    Font_Create = Font_Next_Open
    Font_Make Font_Create, Style, bold, italic, size
ErrorHandler:
    Font_Create = 0
End Function

Public Function Font_Next_Open() As Long
    On Error GoTo Font_Next_Open_Err
    Font_Next_Open = font_last + 1
    Exit Function
Font_Next_Open_Err:
    Call RegistrarError(Err.Number, Err.Description, "Graficos_Textos.Font_Next_Open", Erl)
    Resume Next
End Function

Public Function Font_Check(ByVal font_index As Long) As Boolean
    On Error GoTo Font_Check_Err
    If font_index > 0 And font_index <= font_last Then
        Font_Check = True
    End If
    Exit Function
Font_Check_Err:
    Call RegistrarError(Err.Number, Err.Description, "Graficos_Textos.Font_Check", Erl)
    Resume Next
End Function

Public Function Prepare_Multiline_Text(text As String, ByVal MaxWidth As Integer, Optional ByVal FontIndex As Integer = 1) As String()
    On Error GoTo Handler
    Dim Lines() As String
    If LenB(text) = 0 Then
        ReDim Lines(0)
        Prepare_Multiline_Text = Lines
        Exit Function
    End If
    Dim LetterIndex As Long, CurLetter As Integer, LastBreak As Long, CanBreak As Long, CurWidth As Integer, CurLine As Integer, CanBreakWidth As Integer
    With Fuentes(FontIndex)
        LastBreak = 1
        For LetterIndex = 1 To Len(text)
            CurLetter = Asc(mid$(text, LetterIndex, 1))
            If CurLetter = vbKeyReturn Then
                ReDim Preserve Lines(CurLine)
                If LetterIndex - LastBreak > 0 Then
                    Lines(CurLine) = mid$(text, LastBreak, LetterIndex - LastBreak)
                End If
                LastBreak = LetterIndex + 2
                CanBreak = LastBreak
                CurLine = CurLine + 1
                CurWidth = 0
            Else
                If .Caracteres(CurLetter) <> 0 Then CurWidth = CurWidth + GrhData(.Caracteres(CurLetter)).pixelWidth
                If CurLetter = vbKeySpace Or CurLetter = vbKeyTab Then
                    CanBreak = LetterIndex
                    CanBreakWidth = CurWidth
                End If
                If CurWidth > MaxWidth And MaxWidth > 0 Then
                    ReDim Preserve Lines(CurLine)
                    If CanBreak - LastBreak > 0 Then
                        Lines(CurLine) = mid$(text, LastBreak, CanBreak - LastBreak)
                        CurWidth = CurWidth - CanBreakWidth
                        LastBreak = CanBreak + 1
                    Else
                        Lines(CurLine) = mid$(text, LastBreak, LetterIndex - LastBreak)
                        CurWidth = GrhData(.Caracteres(CurLetter)).pixelWidth
                        LastBreak = LetterIndex
                    End If
                    CanBreak = LastBreak
                    CurLine = CurLine + 1
                End If
            End If
        Next
        If LetterIndex - LastBreak > 0 Then
            ReDim Preserve Lines(CurLine)
            Lines(CurLine) = mid$(text, LastBreak, LetterIndex - LastBreak)
        End If
    End With
    Prepare_Multiline_Text = Lines
    Exit Function
Handler:
    Call RegistrarError(Err.Number, Err.Description, "clsDX8Engine.Prepare_Multiline_Text", Erl)
    ReDim Lines(0)
    Prepare_Multiline_Text = Lines
End Function

Public Function Text_Width(text As String, Optional ByVal FontIndex As Byte = 1) As Integer
    On Error GoTo Handler
    Dim LetterIndex As Long, CurLetter As Integer
    With Fuentes(FontIndex)
        For LetterIndex = 1 To Len(text)
            CurLetter = Asc(mid$(text, LetterIndex, 1))
            Text_Width = Text_Width + GrhData(.Caracteres(CurLetter)).pixelWidth
        Next
    End With
    Exit Function
Handler:
    Call RegistrarError(Err.Number, Err.Description, "clsDX8Engine.Text_Width", Erl)
End Function
