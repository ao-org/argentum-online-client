Attribute VB_Name = "basCryptoSys"
' $Id: basCryptoSys.bas $

'/**
' The VBA/VB6 interface to CryptoSys API.
'
' @author dai
' @version 6.20.0
'**/

' Last updated:
' * $Date: 2021-09-25 10:01:00 $

' Updated [v6.20.0]
' Combined all of basCryptoSys.bas, basCryptoSys64.bas, basCryptoSys64_32.bas and basCryptoSysWrappers.bas
' into this one file.

'********************* COPYRIGHT NOTICE*********************
' Copyright (c) 2001-21 DI Management Services Pty Limited.
' <www.di-mgt.com.au> <www.cryptosys.net>
' All rights reserved.
' The latest version of CryptoSys(tm) API and a licence
' may be obtained from <https://www.cryptosys.net/>.
' Refer to licence for conditions of use.
' This copyright notice must always be left intact.
'****************** END OF COPYRIGHT NOTICE*****************

Option Explicit
Option Base 0

' CONSTANTS
Public Const ENCRYPT As Boolean = True
Public Const DECRYPT As Boolean = False
' Maximum number of bytes in hash digest byte array
Public Const API_MAX_HASH_BYTES As Long = 64
Public Const API_SHA1_BYTES     As Long = 20
Public Const API_SHA224_BYTES   As Long = 28
Public Const API_SHA256_BYTES   As Long = 32
Public Const API_SHA384_BYTES   As Long = 48
Public Const API_SHA512_BYTES   As Long = 64
Public Const API_MD5_BYTES      As Long = 16
Public Const API_MD2_BYTES      As Long = 16
Public Const API_RMD160_BYTES   As Long = 20
' Maximum number of hex characters in hash digest
Public Const API_MAX_HASH_CHARS As Long = (2 * API_MAX_HASH_BYTES)
Public Const API_SHA1_CHARS     As Long = (2 * API_SHA1_BYTES)
Public Const API_SHA224_CHARS   As Long = (2 * API_SHA224_BYTES)
Public Const API_SHA256_CHARS   As Long = (2 * API_SHA256_BYTES)
Public Const API_SHA384_CHARS   As Long = (2 * API_SHA384_BYTES)
Public Const API_SHA512_CHARS   As Long = (2 * API_SHA512_BYTES)
Public Const API_MD5_CHARS      As Long = (2 * API_MD5_BYTES)
Public Const API_MD2_CHARS      As Long = (2 * API_MD2_BYTES)
Public Const API_RMD160_CHARS   As Long = (2 * API_RMD160_BYTES)
' Maximum lengths of MAC tags
Public Const API_MAX_MAC_BYTES  As Long = 64
Public Const API_MAX_HMAC_BYTES As Long = 64
Public Const API_MAX_CMAC_BYTES As Long = 16
Public Const API_MAX_GMAC_BYTES As Long = 16
Public Const API_POLY1305_BYTES As Long = 16
Public Const API_AEAD_TAG_MAX_BYTES As Long = 16
Public Const API_MAX_MAC_CHARS  As Long = (2 * API_MAX_MAC_BYTES)
Public Const API_MAX_HMAC_CHARS As Long = (2 * API_MAX_HMAC_BYTES)
Public Const API_MAX_CMAC_CHARS As Long = (2 * API_MAX_CMAC_BYTES)
Public Const API_MAX_GMAC_CHARS As Long = (2 * API_MAX_GMAC_BYTES)
Public Const API_POLY1305_CHARS As Long = (2 * API_POLY1305_BYTES)
' Synonyms retained for backwards compatibility
Public Const API_MAX_SHA1_BYTES As Long = 20
Public Const API_MAX_SHA2_BYTES As Long = 32  ' (This was for SHA-256)
Public Const API_MAX_MD5_BYTES  As Long = 16
Public Const API_MAX_SHA1_CHARS As Long = (2 * API_MAX_SHA1_BYTES)
Public Const API_MAX_SHA2_CHARS As Long = (2 * API_MAX_SHA2_BYTES)
Public Const API_MAX_MD5_CHARS  As Long = (2 * API_MAX_MD5_BYTES)
' Encryption block sizes in bytes
Public Const API_BLK_DES_BYTES  As Long = 8
Public Const API_BLK_TDEA_BYTES As Long = 8
Public Const API_BLK_BLF_BYTES  As Long = 8
Public Const API_BLK_AES_BYTES  As Long = 16
' Key size in bytes
Public Const API_KEYSIZE_TDEA_BYTES As Long = 24
' Required size for RNG seed file
Public Const API_RNG_SEED_BYTES As Long = 64
' Maximum number of characters in an error lookup message
Public Const API_MAX_ERRORLOOKUP_CHARS = 127

' Options for HASH functions
Public Const API_HASH_SHA1   As Long = 0
Public Const API_HASH_MD5    As Long = 1
Public Const API_HASH_MD2    As Long = 2
Public Const API_HASH_SHA256 As Long = 3
Public Const API_HASH_SHA384 As Long = 4
Public Const API_HASH_SHA512 As Long = 5
Public Const API_HASH_SHA224 As Long = 6
Public Const API_HASH_RMD160 As Long = 7
' SHA-3 added back [v5.3]
Public Const API_HASH_SHA3_224 As Long = &HA&
Public Const API_HASH_SHA3_256 As Long = &HB&
Public Const API_HASH_SHA3_384 As Long = &HC&
Public Const API_HASH_SHA3_512 As Long = &HD&

Public Const API_HASH_MODE_TEXT  As Long = &H10000

' HMAC algorithms
Public Const API_HMAC_SHA1     As Long = 0
Public Const API_HMAC_SHA224   As Long = 6
Public Const API_HMAC_SHA256   As Long = 3
Public Const API_HMAC_SHA384   As Long = 4
Public Const API_HMAC_SHA512   As Long = 5
Public Const API_HMAC_SHA3_224 As Long = &HA&
Public Const API_HMAC_SHA3_256 As Long = &HB&
Public Const API_HMAC_SHA3_384 As Long = &HC&
Public Const API_HMAC_SHA3_512 As Long = &HD&

' Options for MAC functions
Public Const API_CMAC_TDEA    As Long = &H100  ' ) synonyms
Public Const API_CMAC_DESEDE  As Long = &H100  ' ) synonyms
Public Const API_CMAC_AES128  As Long = &H101
Public Const API_CMAC_AES192  As Long = &H102
Public Const API_CMAC_AES256  As Long = &H103
Public Const API_MAC_POLY1305 As Long = &H200
Public Const API_KMAC_128     As Long = &H201
Public Const API_KMAC_256     As Long = &H202
Public Const API_XOF_SHAKE128 As Long = &H203
Public Const API_XOF_SHAKE256 As Long = &H204

' Options for RNG functions
Public Const API_RNG_STRENGTH_112 As Long = &H0
Public Const API_RNG_STRENGTH_128 As Long = &H1

' Block cipher (BC) algorithm options
Public Const API_BC_TDEA    As Long = &H10  ' )
Public Const API_BC_DESEDE3 As Long = &H10  ' ) equiv. synonyms for Triple DES
Public Const API_BC_3DES    As Long = &H10  ' )
Public Const API_BC_AES128  As Long = &H20
Public Const API_BC_AES192  As Long = &H30
Public Const API_BC_AES256  As Long = &H40

' Block cipher mode options
Public Const API_MODE_ECB As Long = &H0
Public Const API_MODE_CBC As Long = &H100
Public Const API_MODE_OFB As Long = &H200
Public Const API_MODE_CFB As Long = &H300
Public Const API_MODE_CTR As Long = &H400

' Block cipher option flags
Public Const API_IV_PREFIX  As Long = &H1000
Public Const API_PAD_LEAVE As Long = &H2000

' Block cipher padding options
Public Const API_PAD_DEFAULT As Long = &H0
Public Const API_PAD_NOPAD  As Long = &H10000
Public Const API_PAD_PKCS5  As Long = &H20000
Public Const API_PAD_1ZERO  As Long = &H30000
Public Const API_PAD_AX923  As Long = &H40000
Public Const API_PAD_W3C    As Long = &H50000

' Stream cipher (SC) algorithm options (NB no zero default)
Public Const API_SC_ARCFOUR  As Long = 1
Public Const API_SC_SALSA20  As Long = 2
Public Const API_SC_CHACHA20 As Long = 3

' AEAD algorithm options
Public Const API_AEAD_AES_128_GCM  As Long = 1
Public Const API_AEAD_AES_256_GCM  As Long = 2
Public Const API_AEAD_CHACHA20_POLY1305 As Long = 29

' Wipefile options
Public Const API_WIPEFILE_DOD7 As Long = &H0    ' Default
Public Const API_WIPEFILE_SIMPLE As Long = &H1

' Compression algorithm options - added [v6.20]
Public Const API_COMPR_ZLIB As Long = &H0    ' Default
Public Const API_COMPR_ZSTD As Long = &H1

' General
Public Const API_GEN_PLATFORM As Long = &H40

' *********************
' FUNCTION DECLARATIONS
' *********************

#If VBA7 Then
' Declarations for 64-bit Office
' (In VB6 these will appear red. Turn off "Auto Syntax Check" in Tools > Options to avoid annoying warnings)

' GENERAL FUNCTIONS
Public Declare PtrSafe Function API_Version Lib "diCryptoSys.dll" () As Long
Public Declare PtrSafe Function API_LicenceType Lib "diCryptoSys.dll" (ByVal nReserved As Long) As Long
Public Declare PtrSafe Function API_CompileTime Lib "diCryptoSys.dll" (ByVal strCompiledOn As String, ByVal nMaxChars As Long) As Long
Public Declare PtrSafe Function API_ModuleName Lib "diCryptoSys.dll" (ByVal strModuleName As String, ByVal nMaxChars As Long, ByVal nReserved As Long) As Long
Public Declare PtrSafe Function API_ErrorCode Lib "diCryptoSys.dll" () As Long    ' Added [v4.2] (only for certain fns)
Public Declare PtrSafe Function API_ErrorLookup Lib "diCryptoSys.dll" (ByVal strErrMsg As String, ByVal nMaxChars As Long, ByVal nErrCode As Long) As Long
Public Declare PtrSafe Function API_PowerUpTests Lib "diCryptoSys.dll" (ByVal nReserved As Long) As Long

' ADVANCED ENCRYPTION STANDARD (AES) BLOCK CIPHER WITH 128-BIT KEY
Public Declare PtrSafe Function AES128_Bytes Lib "diCryptoSys.dll" (ByRef lpOutput As Byte, ByRef lpData As Byte, ByVal nBytes As Long, ByRef lpKey As Byte, ByVal fEncrypt As Long) As Long
Public Declare PtrSafe Function AES128_BytesMode Lib "diCryptoSys.dll" (ByRef lpOutput As Byte, ByRef lpData As Byte, ByVal nBytes As Long, ByRef lpKey As Byte, ByVal fEncrypt As Long, ByVal strMode As String, ByRef lpInitV As Byte) As Long
Public Declare PtrSafe Function AES128_Hex Lib "diCryptoSys.dll" (ByVal strOutput As String, ByVal strInput As String, ByVal strHexKey As String, ByVal fEncrypt As Long) As Long
Public Declare PtrSafe Function AES128_HexMode Lib "diCryptoSys.dll" (ByVal strOutput As String, ByVal strInput As String, ByVal strHexKey As String, ByVal fEncrypt As Long, ByVal strMode As String, ByVal strHexIV As String) As Long
Public Declare PtrSafe Function AES128_B64Mode Lib "diCryptoSys.dll" (ByVal strOutput As String, ByVal strInput As String, ByVal strB64Key As String, ByVal fEncrypt As Long, ByVal strMode As String, ByVal strB64IV As String) As Long
Public Declare PtrSafe Function AES128_File Lib "diCryptoSys.dll" (ByVal strFileOut As String, ByVal strFileIn As String, ByRef lpKey As Byte, ByVal fEncrypt As Long, ByVal strMode As String, ByRef lpInitV As Byte) As Long
Public Declare PtrSafe Function AES128_FileExt Lib "diCryptoSys.dll" (ByVal strFileOut As String, ByVal strFileIn As String, ByRef lpKey As Byte, ByVal fEncrypt As Long, ByVal strMode As String, ByRef lpInitV As Byte, ByVal nOptions As Long) As Long
Public Declare PtrSafe Function AES128_FileHex Lib "diCryptoSys.dll" (ByVal strFileOut As String, ByVal strFileIn As String, ByVal strHexKey As String, ByVal fEncrypt As Long, ByVal strMode As String, ByVal strHexIV As String) As Long
Public Declare PtrSafe Function AES128_Init Lib "diCryptoSys.dll" (ByRef lpKey As Byte, ByVal fEncrypt As Long, ByVal strMode As String, ByRef lpInitV As Byte) As Long
Public Declare PtrSafe Function AES128_InitHex Lib "diCryptoSys.dll" (ByVal strHexKey As String, ByVal fEncrypt As Long, ByVal strMode As String, ByVal strHexIV As String) As Long
Public Declare PtrSafe Function AES128_Update Lib "diCryptoSys.dll" (ByVal hContext As Long, ByRef lpData As Byte, ByVal nBytes As Long) As Long
Public Declare PtrSafe Function AES128_UpdateHex Lib "diCryptoSys.dll" (ByVal hContext As Long, ByVal strHexData As String) As Long
Public Declare PtrSafe Function AES128_Final Lib "diCryptoSys.dll" (ByVal hContext As Long) As Long
Public Declare PtrSafe Function AES128_InitError Lib "diCryptoSys.dll" () As Long

' ADVANCED ENCRYPTION STANDARD (AES) BLOCK CIPHER WITH 192-BIT KEY
Public Declare PtrSafe Function AES192_Bytes Lib "diCryptoSys.dll" (ByRef lpOutput As Byte, ByRef lpData As Byte, ByVal nBytes As Long, ByRef lpKey As Byte, ByVal fEncrypt As Long) As Long
Public Declare PtrSafe Function AES192_BytesMode Lib "diCryptoSys.dll" (ByRef lpOutput As Byte, ByRef lpData As Byte, ByVal nBytes As Long, ByRef lpKey As Byte, ByVal fEncrypt As Long, ByVal strMode As String, ByRef lpInitV As Byte) As Long
Public Declare PtrSafe Function AES192_Hex Lib "diCryptoSys.dll" (ByVal strOutput As String, ByVal strInput As String, ByVal strHexKey As String, ByVal fEncrypt As Long) As Long
Public Declare PtrSafe Function AES192_HexMode Lib "diCryptoSys.dll" (ByVal strOutput As String, ByVal strInput As String, ByVal strHexKey As String, ByVal fEncrypt As Long, ByVal strMode As String, ByVal strHexIV As String) As Long
Public Declare PtrSafe Function AES192_B64Mode Lib "diCryptoSys.dll" (ByVal strOutput As String, ByVal strInput As String, ByVal strB64Key As String, ByVal fEncrypt As Long, ByVal strMode As String, ByVal strB64IV As String) As Long
Public Declare PtrSafe Function AES192_File Lib "diCryptoSys.dll" (ByVal strFileOut As String, ByVal strFileIn As String, ByRef lpKey As Byte, ByVal fEncrypt As Long, ByVal strMode As String, ByRef lpInitV As Byte) As Long
Public Declare PtrSafe Function AES192_FileExt Lib "diCryptoSys.dll" (ByVal strFileOut As String, ByVal strFileIn As String, ByRef lpKey As Byte, ByVal fEncrypt As Long, ByVal strMode As String, ByRef lpInitV As Byte, ByVal nOptions As Long) As Long
Public Declare PtrSafe Function AES192_FileHex Lib "diCryptoSys.dll" (ByVal strFileOut As String, ByVal strFileIn As String, ByVal strHexKey As String, ByVal fEncrypt As Long, ByVal strMode As String, ByVal strHexIV As String) As Long
Public Declare PtrSafe Function AES192_Init Lib "diCryptoSys.dll" (ByRef lpKey As Byte, ByVal fEncrypt As Long, ByVal strMode As String, ByRef lpInitV As Byte) As Long
Public Declare PtrSafe Function AES192_InitHex Lib "diCryptoSys.dll" (ByVal strHexKey As String, ByVal fEncrypt As Long, ByVal strMode As String, ByVal strHexIV As String) As Long
Public Declare PtrSafe Function AES192_Update Lib "diCryptoSys.dll" (ByVal hContext As Long, ByRef lpData As Byte, ByVal nBytes As Long) As Long
Public Declare PtrSafe Function AES192_UpdateHex Lib "diCryptoSys.dll" (ByVal hContext As Long, ByVal strHexData As String) As Long
Public Declare PtrSafe Function AES192_Final Lib "diCryptoSys.dll" (ByVal hContext As Long) As Long
Public Declare PtrSafe Function AES192_InitError Lib "diCryptoSys.dll" () As Long

' ADVANCED ENCRYPTION STANDARD (AES) BLOCK CIPHER WITH 256-BIT KEY
Public Declare PtrSafe Function AES256_Bytes Lib "diCryptoSys.dll" (ByRef lpOutput As Byte, ByRef lpData As Byte, ByVal nBytes As Long, ByRef lpKey As Byte, ByVal fEncrypt As Long) As Long
Public Declare PtrSafe Function AES256_BytesMode Lib "diCryptoSys.dll" (ByRef lpOutput As Byte, ByRef lpData As Byte, ByVal nBytes As Long, ByRef lpKey As Byte, ByVal fEncrypt As Long, ByVal strMode As String, ByRef lpInitV As Byte) As Long
Public Declare PtrSafe Function AES256_Hex Lib "diCryptoSys.dll" (ByVal strOutput As String, ByVal strInput As String, ByVal strHexKey As String, ByVal fEncrypt As Long) As Long
Public Declare PtrSafe Function AES256_HexMode Lib "diCryptoSys.dll" (ByVal strOutput As String, ByVal strInput As String, ByVal strHexKey As String, ByVal fEncrypt As Long, ByVal strMode As String, ByVal strHexIV As String) As Long
Public Declare PtrSafe Function AES256_B64Mode Lib "diCryptoSys.dll" (ByVal strOutput As String, ByVal strInput As String, ByVal strB64Key As String, ByVal fEncrypt As Long, ByVal strMode As String, ByVal strB64IV As String) As Long
Public Declare PtrSafe Function AES256_File Lib "diCryptoSys.dll" (ByVal strFileOut As String, ByVal strFileIn As String, ByRef lpKey As Byte, ByVal fEncrypt As Long, ByVal strMode As String, ByRef lpInitV As Byte) As Long
Public Declare PtrSafe Function AES256_FileExt Lib "diCryptoSys.dll" (ByVal strFileOut As String, ByVal strFileIn As String, ByRef lpKey As Byte, ByVal fEncrypt As Long, ByVal strMode As String, ByRef lpInitV As Byte, ByVal nOptions As Long) As Long
Public Declare PtrSafe Function AES256_FileHex Lib "diCryptoSys.dll" (ByVal strFileOut As String, ByVal strFileIn As String, ByVal strHexKey As String, ByVal fEncrypt As Long, ByVal strMode As String, ByVal strHexIV As String) As Long
Public Declare PtrSafe Function AES256_Init Lib "diCryptoSys.dll" (ByRef lpKey As Byte, ByVal fEncrypt As Long, ByVal strMode As String, ByRef lpInitV As Byte) As Long
Public Declare PtrSafe Function AES256_InitHex Lib "diCryptoSys.dll" (ByVal strHexKey As String, ByVal fEncrypt As Long, ByVal strMode As String, ByVal strHexIV As String) As Long
Public Declare PtrSafe Function AES256_Update Lib "diCryptoSys.dll" (ByVal hContext As Long, ByRef lpData As Byte, ByVal nBytes As Long) As Long
Public Declare PtrSafe Function AES256_UpdateHex Lib "diCryptoSys.dll" (ByVal hContext As Long, ByVal strHexData As String) As Long
Public Declare PtrSafe Function AES256_Final Lib "diCryptoSys.dll" (ByVal hContext As Long) As Long
Public Declare PtrSafe Function AES256_InitError Lib "diCryptoSys.dll" () As Long

' BLOWFISH BLOCK CIPHER FUNCTIONS
Public Declare PtrSafe Function BLF_Bytes Lib "diCryptoSys.dll" (ByRef lpOutput As Byte, ByRef lpData As Byte, ByVal nBytes As Long, ByRef lpKey As Byte, ByVal nKeyBytes As Long, ByVal fEncrypt As Long) As Long
Public Declare PtrSafe Function BLF_BytesMode Lib "diCryptoSys.dll" (ByRef lpOutput As Byte, ByRef lpData As Byte, ByVal nBytes As Long, ByRef lpKey As Byte, ByVal nKeyBytes As Long, ByVal fEncrypt As Long, ByVal strMode As String, ByRef lpInitV As Byte) As Long
Public Declare PtrSafe Function BLF_Hex Lib "diCryptoSys.dll" (ByVal strOutput As String, ByVal strInput As String, ByVal strHexKey As String, ByVal fEncrypt As Long) As Long
Public Declare PtrSafe Function BLF_HexMode Lib "diCryptoSys.dll" (ByVal strOutput As String, ByVal strInput As String, ByVal strHexKey As String, ByVal fEncrypt As Long, ByVal strMode As String, ByVal strHexIV As String) As Long
Public Declare PtrSafe Function BLF_B64Mode Lib "diCryptoSys.dll" (ByVal strOutput As String, ByVal strInput As String, ByVal strB64Key As String, ByVal fEncrypt As Long, ByVal strMode As String, ByVal strB64IV As String) As Long
Public Declare PtrSafe Function BLF_File Lib "diCryptoSys.dll" (ByVal strFileOut As String, ByVal strFileIn As String, ByRef lpKey As Byte, ByVal nKeyBytes As Long, ByVal fEncrypt As Long, ByVal strMode As String, ByRef lpInitV As Byte) As Long
Public Declare PtrSafe Function BLF_FileExt Lib "diCryptoSys.dll" (ByVal strFileOut As String, ByVal strFileIn As String, ByRef lpKey As Byte, ByVal nKeyBytes As Long, ByVal fEncrypt As Long, ByVal strMode As String, ByRef lpInitV As Byte, ByVal nOptions As Long) As Long
Public Declare PtrSafe Function BLF_FileHex Lib "diCryptoSys.dll" (ByVal strFileOut As String, ByVal strFileIn As String, ByVal strHexKey As String, ByVal fEncrypt As Long, ByVal strMode As String, ByVal strHexIV As String) As Long
Public Declare PtrSafe Function BLF_InitHex Lib "diCryptoSys.dll" (ByVal strHexKey As String, ByVal fEncrypt As Long, ByVal strMode As String, ByVal strHexIV As String) As Long
Public Declare PtrSafe Function BLF_Init Lib "diCryptoSys.dll" (ByRef lpKey As Byte, ByVal nKeyBytes As Long, ByVal fEncrypt As Long, ByVal strMode As String, ByRef lpInitV As Byte) As Long
Public Declare PtrSafe Function BLF_UpdateHex Lib "diCryptoSys.dll" (ByVal hContext As Long, ByVal strHexData As String) As Long
Public Declare PtrSafe Function BLF_Update Lib "diCryptoSys.dll" (ByVal hContext As Long, ByRef lpData As Byte, ByVal nBytes As Long) As Long
Public Declare PtrSafe Function BLF_Final Lib "diCryptoSys.dll" (ByVal hContext As Long) As Long
Public Declare PtrSafe Function BLF_InitError Lib "diCryptoSys.dll" () As Long
    
' DATA ENCRYPTION STANDARD (DES) BLOCK CIPHER
Public Declare PtrSafe Function DES_Bytes Lib "diCryptoSys.dll" (ByRef lpOutput As Byte, ByRef lpData As Byte, ByVal nBytes As Long, ByRef lpKey As Byte, ByVal fEncrypt As Long) As Long
Public Declare PtrSafe Function DES_BytesMode Lib "diCryptoSys.dll" (ByRef lpOutput As Byte, ByRef lpData As Byte, ByVal nBytes As Long, ByRef lpKey As Byte, ByVal fEncrypt As Long, ByVal strMode As String, ByRef lpInitV As Byte) As Long
Public Declare PtrSafe Function DES_Hex Lib "diCryptoSys.dll" (ByVal strOutput As String, ByVal strInput As String, ByVal strHexKey As String, ByVal fEncrypt As Long) As Long
Public Declare PtrSafe Function DES_HexMode Lib "diCryptoSys.dll" (ByVal strOutput As String, ByVal strInput As String, ByVal strHexKey As String, ByVal fEncrypt As Long, ByVal strMode As String, ByVal strHexIV As String) As Long
Public Declare PtrSafe Function DES_B64Mode Lib "diCryptoSys.dll" (ByVal strOutput As String, ByVal strInput As String, ByVal strB64Key As String, ByVal fEncrypt As Long, ByVal strMode As String, ByVal strB64IV As String) As Long
Public Declare PtrSafe Function DES_File Lib "diCryptoSys.dll" (ByVal strFileOut As String, ByVal strFileIn As String, ByRef lpKey As Byte, ByVal fEncrypt As Long, ByVal strMode As String, ByRef lpInitV As Byte) As Long
Public Declare PtrSafe Function DES_FileExt Lib "diCryptoSys.dll" (ByVal strFileOut As String, ByVal strFileIn As String, ByRef lpKey As Byte, ByVal fEncrypt As Long, ByVal strMode As String, ByRef lpInitV As Byte, ByVal nOptions As Long) As Long
Public Declare PtrSafe Function DES_FileHex Lib "diCryptoSys.dll" (ByVal strFileOut As String, ByVal strFileIn As String, ByVal strHexKey As String, ByVal fEncrypt As Long, ByVal strMode As String, ByVal strHexIV As String) As Long
Public Declare PtrSafe Function DES_Init Lib "diCryptoSys.dll" (ByRef lpKey As Byte, ByVal fEncrypt As Long, ByVal strMode As String, ByRef lpInitV As Byte) As Long
Public Declare PtrSafe Function DES_InitHex Lib "diCryptoSys.dll" (ByVal strHexKey As String, ByVal fEncrypt As Long, ByVal strMode As String, ByVal strHexIV As String) As Long
Public Declare PtrSafe Function DES_Update Lib "diCryptoSys.dll" (ByVal hContext As Long, ByRef lpData As Byte, ByVal nBytes As Long) As Long
Public Declare PtrSafe Function DES_UpdateHex Lib "diCryptoSys.dll" (ByVal hContext As Long, ByVal strHexData As String) As Long
Public Declare PtrSafe Function DES_Final Lib "diCryptoSys.dll" (ByVal hContext As Long) As Long
Public Declare PtrSafe Function DES_InitError Lib "diCryptoSys.dll" () As Long
    
' Checks for weak or invalid-length DES or TDEA keys
Public Declare PtrSafe Function DES_CheckKey Lib "diCryptoSys.dll" (ByRef lpKey As Byte, ByVal nKeyBytes As Long) As Long
Public Declare PtrSafe Function DES_CheckKeyHex Lib "diCryptoSys.dll" (ByVal strHexKey As String) As Long

' TRIPLE DATA ENCRYPTION ALGORITHM (TDEA) BLOCK CIPHER
Public Declare PtrSafe Function TDEA_Bytes Lib "diCryptoSys.dll" (ByRef lpOutput As Byte, ByRef lpData As Byte, ByVal nBytes As Long, ByRef lpKey As Byte, ByVal fEncrypt As Long) As Long
Public Declare PtrSafe Function TDEA_BytesMode Lib "diCryptoSys.dll" (ByRef lpOutput As Byte, ByRef lpData As Byte, ByVal nBytes As Long, ByRef lpKey As Byte, ByVal fEncrypt As Long, ByVal strMode As String, ByRef lpInitV As Byte) As Long
Public Declare PtrSafe Function TDEA_Hex Lib "diCryptoSys.dll" (ByVal strOutput As String, ByVal strInput As String, ByVal strHexKey As String, ByVal fEncrypt As Long) As Long
Public Declare PtrSafe Function TDEA_HexMode Lib "diCryptoSys.dll" (ByVal strOutput As String, ByVal strInput As String, ByVal strHexKey As String, ByVal fEncrypt As Long, ByVal strMode As String, ByVal strHexIV As String) As Long
Public Declare PtrSafe Function TDEA_B64Mode Lib "diCryptoSys.dll" (ByVal strOutput As String, ByVal strInput As String, ByVal strB64Key As String, ByVal fEncrypt As Long, ByVal strMode As String, ByVal strB64IV As String) As Long
Public Declare PtrSafe Function TDEA_File Lib "diCryptoSys.dll" (ByVal strFileOut As String, ByVal strFileIn As String, ByRef lpKey As Byte, ByVal fEncrypt As Long, ByVal strMode As String, ByRef lpInitV As Byte) As Long
Public Declare PtrSafe Function TDEA_FileExt Lib "diCryptoSys.dll" (ByVal strFileOut As String, ByVal strFileIn As String, ByRef lpKey As Byte, ByVal fEncrypt As Long, ByVal strMode As String, ByRef lpInitV As Byte, ByVal nOptions As Long) As Long
Public Declare PtrSafe Function TDEA_FileHex Lib "diCryptoSys.dll" (ByVal strFileOut As String, ByVal strFileIn As String, ByVal strHexKey As String, ByVal fEncrypt As Long, ByVal strMode As String, ByVal strHexIV As String) As Long
Public Declare PtrSafe Function TDEA_Init Lib "diCryptoSys.dll" (ByRef lpKey As Byte, ByVal fEncrypt As Long, ByVal strMode As String, ByRef lpInitV As Byte) As Long
Public Declare PtrSafe Function TDEA_InitHex Lib "diCryptoSys.dll" (ByVal strHexKey As String, ByVal fEncrypt As Long, ByVal strMode As String, ByVal strHexIV As String) As Long
Public Declare PtrSafe Function TDEA_Update Lib "diCryptoSys.dll" (ByVal hContext As Long, ByRef lpData As Byte, ByVal nBytes As Long) As Long
Public Declare PtrSafe Function TDEA_UpdateHex Lib "diCryptoSys.dll" (ByVal hContext As Long, ByVal strHexData As String) As Long
Public Declare PtrSafe Function TDEA_Final Lib "diCryptoSys.dll" (ByVal hContext As Long) As Long
Public Declare PtrSafe Function TDEA_InitError Lib "diCryptoSys.dll" () As Long

' GENERIC BLOCK CIPHER FUNCTIONS
' Added in [v6.20] (to get rid of that annoying 2)
Public Declare PtrSafe Function CIPHER_EncryptBytes Lib "diCryptoSys.dll" (ByRef lpOutput As Byte, ByVal nOutBytes As Long, ByRef lpInput As Byte, ByVal nInputLen As Long, ByRef lpKey As Byte, ByVal nKeyLen As Long, ByRef lpIV As Byte, ByVal nIvLen As Long, ByVal strAlgModePad As String, ByVal nOptions As Long) As Long
Public Declare PtrSafe Function CIPHER_DecryptBytes Lib "diCryptoSys.dll" (ByRef lpOutput As Byte, ByVal nOutBytes As Long, ByRef lpInput As Byte, ByVal nInputLen As Long, ByRef lpKey As Byte, ByVal nKeyLen As Long, ByRef lpIV As Byte, ByVal nIvLen As Long, ByVal strAlgModePad As String, ByVal nOptions As Long) As Long
' Aliases for backwards compatibility
Public Declare PtrSafe Function CIPHER_EncryptBytes2 Lib "diCryptoSys.dll" Alias "CIPHER_EncryptBytes" (ByRef lpOutput As Byte, ByVal nOutBytes As Long, ByRef lpInput As Byte, ByVal nInputLen As Long, ByRef lpKey As Byte, ByVal nKeyLen As Long, ByRef lpIV As Byte, ByVal nIvLen As Long, ByVal strAlgModePad As String, ByVal nOptions As Long) As Long
Public Declare PtrSafe Function CIPHER_DecryptBytes2 Lib "diCryptoSys.dll" Alias "CIPHER_DecryptBytes" (ByRef lpOutput As Byte, ByVal nOutBytes As Long, ByRef lpInput As Byte, ByVal nInputLen As Long, ByRef lpKey As Byte, ByVal nKeyLen As Long, ByRef lpIV As Byte, ByVal nIvLen As Long, ByVal strAlgModePad As String, ByVal nOptions As Long) As Long

Public Declare PtrSafe Function CIPHER_FileEncrypt Lib "diCryptoSys.dll" (ByVal strFileOut As String, ByVal strFileIn As String, ByRef lpKey As Byte, ByVal nKeyLen As Long, ByRef lpIV As Byte, ByVal nIvLen As Long, ByVal strAlgModePad As String, ByVal nOptions As Long) As Long
Public Declare PtrSafe Function CIPHER_FileDecrypt Lib "diCryptoSys.dll" (ByVal strFileOut As String, ByVal strFileIn As String, ByRef lpKey As Byte, ByVal nKeyLen As Long, ByRef lpIV As Byte, ByVal nIvLen As Long, ByVal strAlgModePad As String, ByVal nOptions As Long) As Long
' New in [v6.0]
Public Declare PtrSafe Function CIPHER_EncryptHex Lib "diCryptoSys.dll" (ByVal strOutput As String, ByVal nOutChars As Long, ByVal strInputHex As String, ByVal strKeyHex As String, ByVal strIvHex As String, ByVal strAlgModePad As String, ByVal nOptions As Long) As Long
Public Declare PtrSafe Function CIPHER_DecryptHex Lib "diCryptoSys.dll" (ByVal strOutput As String, ByVal nOutChars As Long, ByVal strInputHex As String, ByVal strKeyHex As String, ByVal strIvHex As String, ByVal strAlgModePad As String, ByVal nOptions As Long) As Long
' Stateful CIPHER functions added in [v6.0]
Public Declare PtrSafe Function CIPHER_Init Lib "diCryptoSys.dll" (ByVal fEncrypt As Integer, ByVal strAlgAndMode As String, ByRef lpKey As Byte, ByVal nKeyLen As Long, ByRef lpIV As Byte, ByVal nIvLen As Long, ByVal nOptions As Long) As Long
Public Declare PtrSafe Function CIPHER_InitHex Lib "diCryptoSys.dll" (ByVal fEncrypt As Integer, ByVal strAlgAndMode As String, ByVal strKeyHex As String, ByVal strIvHex As String, ByVal nOptions As Long) As Long
Public Declare PtrSafe Function CIPHER_Update Lib "diCryptoSys.dll" (ByVal hContext As Long, ByRef lpOutput As Byte, ByVal nOutBytes As Long, ByRef lpData As Byte, ByVal nDataLen As Long) As Long
Public Declare PtrSafe Function CIPHER_UpdateHex Lib "diCryptoSys.dll" (ByVal hContext As Long, ByVal strOutput As String, ByVal nOutChars As Long, ByVal strDataHex As String) As Long
Public Declare PtrSafe Function CIPHER_Final Lib "diCryptoSys.dll" (ByVal hContext As Long) As Long

' KEY WRAP FUNCTIONS
Public Declare PtrSafe Function CIPHER_KeyWrap Lib "diCryptoSys.dll" (ByRef lpOutput As Byte, ByVal nOutBytes As Long, ByRef lpData As Byte, ByVal nDataLen As Long, ByRef lpKek As Byte, ByVal nKekLen As Long, ByVal nOptions As Long) As Long
Public Declare PtrSafe Function CIPHER_KeyUnwrap Lib "diCryptoSys.dll" (ByRef lpOutput As Byte, ByVal nOutBytes As Long, ByRef lpData As Byte, ByVal nDataLen As Long, ByRef lpKek As Byte, ByVal nKekLen As Long, ByVal nOptions As Long) As Long

' STREAM CIPHER FUNCTIONS
Public Declare PtrSafe Function CIPHER_StreamBytes Lib "diCryptoSys.dll" (ByRef lpOutput As Byte, ByRef lpData As Byte, ByVal nDataLen As Long, ByRef lpKey As Byte, ByVal nKeyLen As Long, ByRef lpIV As Byte, ByVal nIvLen As Long, ByVal nCounter As Long, ByVal nOptions As Long) As Long
Public Declare PtrSafe Function CIPHER_StreamHex Lib "diCryptoSys.dll" (ByVal strOutput As String, ByVal nOutChars As Long, ByVal strInputHex As String, ByVal strKeyHex As String, ByVal strIvHex As String, ByVal nCounter As Long, ByVal nOptions As Long) As Long
Public Declare PtrSafe Function CIPHER_StreamFile Lib "diCryptoSys.dll" (ByVal strFileOut As String, ByVal strFileIn As String, ByRef lpKey As Byte, ByVal nKeyLen As Long, ByRef lpIV As Byte, ByVal nIvLen As Long, ByVal nCounter As Long, ByVal nOptions As Long) As Long
Public Declare PtrSafe Function CIPHER_StreamInit Lib "diCryptoSys.dll" (ByRef lpKey As Byte, ByVal nKeyLen As Long, ByRef lpIV As Byte, ByVal nIvLen As Long, ByVal nCounter As Long, ByVal nOptions As Long) As Long
Public Declare PtrSafe Function CIPHER_StreamUpdate Lib "diCryptoSys.dll" (ByVal hContext As Long, ByRef lpOutput As Byte, ByRef lpData As Byte, ByVal nDataLen As Long) As Long
Public Declare PtrSafe Function CIPHER_StreamFinal Lib "diCryptoSys.dll" (ByVal hContext As Long) As Long

' AEAD FUNCTIONS
Public Declare PtrSafe Function AEAD_Encrypt Lib "diCryptoSys.dll" (ByRef lpOutput As Byte, ByVal nOutLen As Long, ByRef lpTagOut As Byte, ByVal nTagLen As Long, ByRef lpData As Byte, ByVal nDataLen As Long, ByRef lpKey As Byte, ByVal nKeyLen As Long, ByRef lpNonce As Byte, ByVal nNonceLen As Long, ByRef lpAAD As Byte, ByVal nAadLen As Long, ByVal nOptions As Long) As Long
Public Declare PtrSafe Function AEAD_Decrypt Lib "diCryptoSys.dll" (ByRef lpOutput As Byte, ByVal nOutLen As Long, ByRef lpData As Byte, ByVal nDataLen As Long, ByRef lpKey As Byte, ByVal nKeyLen As Long, ByRef lpNonce As Byte, ByVal nNonceLen As Long, ByRef lpAAD As Byte, ByVal nAadLen As Long, ByRef lpTag As Byte, ByVal nTagLen As Long, ByVal nOptions As Long) As Long
Public Declare PtrSafe Function AEAD_InitKey Lib "diCryptoSys.dll" (ByRef lpKey As Byte, ByVal nKeyLen As Long, ByVal nOptions As Long) As Long
Public Declare PtrSafe Function AEAD_SetNonce Lib "diCryptoSys.dll" (ByVal hContext As Long, ByRef lpNonce As Byte, ByVal nNonceLen As Long) As Long
Public Declare PtrSafe Function AEAD_AddAAD Lib "diCryptoSys.dll" (ByVal hContext As Long, ByRef lpAAD As Byte, ByVal nAadLen As Long) As Long
Public Declare PtrSafe Function AEAD_StartEncrypt Lib "diCryptoSys.dll" (ByVal hContext As Long) As Long
Public Declare PtrSafe Function AEAD_StartDecrypt Lib "diCryptoSys.dll" (ByVal hContext As Long, ByRef lpTagToCheck As Byte, ByVal nTagLen As Long) As Long
Public Declare PtrSafe Function AEAD_Update Lib "diCryptoSys.dll" (ByVal hContext As Long, ByRef lpOutput As Byte, ByVal nOutLen As Long, ByRef lpData As Byte, ByVal nDataLen As Long) As Long
Public Declare PtrSafe Function AEAD_FinishEncrypt Lib "diCryptoSys.dll" (ByVal hContext As Long, ByRef lpTagOut As Byte, ByVal nTagLen As Long) As Long
Public Declare PtrSafe Function AEAD_FinishDecrypt Lib "diCryptoSys.dll" (ByVal hContext As Long) As Long
Public Declare PtrSafe Function AEAD_Destroy Lib "diCryptoSys.dll" (ByVal hContext As Long) As Long
' Added in [v5.4]
Public Declare PtrSafe Function AEAD_EncryptWithTag Lib "diCryptoSys.dll" (ByRef lpOutput As Byte, ByVal nOutLen As Long, ByRef lpData As Byte, ByVal nDataLen As Long, ByRef lpKey As Byte, ByVal nKeyLen As Long, ByRef lpNonce As Byte, ByVal nNonceLen As Long, ByRef lpAAD As Byte, ByVal nAadLen As Long, ByVal nOptions As Long) As Long
Public Declare PtrSafe Function AEAD_DecryptWithTag Lib "diCryptoSys.dll" (ByRef lpOutput As Byte, ByVal nOutLen As Long, ByRef lpData As Byte, ByVal nDataLen As Long, ByRef lpKey As Byte, ByVal nKeyLen As Long, ByRef lpNonce As Byte, ByVal nNonceLen As Long, ByRef lpAAD As Byte, ByVal nAadLen As Long, ByVal nOptions As Long) As Long

' GCM AUTHENTICATED EN/DECRYPTION FUNCTIONS
' Partly superseded by AEAD functions in [v5.1]
Public Declare PtrSafe Function GCM_Encrypt Lib "diCryptoSys.dll" (ByRef lpOutput As Byte, ByVal nOutLen As Long, ByRef lpTagOut As Byte, ByVal nTagLen As Long, ByRef lpData As Byte, ByVal nDataLen As Long, ByRef lpKey As Byte, ByVal nKeyLen As Long, ByRef lpIV As Byte, ByVal nIvLen As Long, ByRef lpAAD As Byte, ByVal nAadLen As Long, ByVal nOptions As Long) As Long
Public Declare PtrSafe Function GCM_Decrypt Lib "diCryptoSys.dll" (ByRef lpOutput As Byte, ByVal nOutLen As Long, ByRef lpData As Byte, ByVal nDataLen As Long, ByRef lpKey As Byte, ByVal nKeyLen As Long, ByRef lpIV As Byte, ByVal nIvLen As Long, ByRef lpAAD As Byte, ByVal nAadLen As Long, ByRef lpTag As Byte, ByVal nTagLen As Long, ByVal nOptions As Long) As Long
Public Declare PtrSafe Function GCM_InitKey Lib "diCryptoSys.dll" (ByRef lpKey As Byte, ByVal nKeyLen As Long, ByVal nOptions As Long) As Long
Public Declare PtrSafe Function GCM_NextEncrypt Lib "diCryptoSys.dll" (ByVal hContext As Long, ByRef lpOutput As Byte, ByVal nOutLen As Long, ByRef lpTagOut As Byte, ByVal nTagLen As Long, ByRef lpData As Byte, ByVal nDataLen As Long, ByRef lpIV As Byte, ByVal nIvLen As Long, ByRef lpAAD As Byte, ByVal nAadLen As Long) As Long
Public Declare PtrSafe Function GCM_NextDecrypt Lib "diCryptoSys.dll" (ByVal hContext As Long, ByRef lpOutput As Byte, ByVal nOutLen As Long, ByRef lpData As Byte, ByVal nDataLen As Long, ByRef lpIV As Byte, ByVal nIvLen As Long, ByRef lpAAD As Byte, ByVal nAadLen As Long, ByRef lpTag As Byte, ByVal nTagLen As Long) As Long
Public Declare PtrSafe Function GCM_FinishKey Lib "diCryptoSys.dll" (ByVal hContext As Long) As Long

' GENERIC MESSAGE DIGEST HASH FUNCTIONS
Public Declare PtrSafe Function HASH_Bytes Lib "diCryptoSys.dll" (ByRef lpOutput As Byte, ByVal nOutLen As Long, ByRef lpMessage As Byte, ByVal nMsgLen As Long, ByVal nOptions As Long) As Long
Public Declare PtrSafe Function HASH_File Lib "diCryptoSys.dll" (ByRef lpOutput As Byte, ByVal nOutLen As Long, ByVal strFileName As String, ByVal nOptions As Long) As Long
Public Declare PtrSafe Function HASH_HexFromBytes Lib "diCryptoSys.dll" (ByVal strOutput As String, ByVal nMaxChars As Long, ByRef lpMessage As Byte, ByVal nMsgLen As Long, ByVal nOptions As Long) As Long
Public Declare PtrSafe Function HASH_HexFromFile Lib "diCryptoSys.dll" (ByVal strOutput As String, ByVal nMaxChars As Long, ByVal strFileName As String, ByVal nOptions As Long) As Long
Public Declare PtrSafe Function HASH_HexFromHex Lib "diCryptoSys.dll" (ByVal strOutput As String, ByVal nMaxChars As Long, ByVal strMsgHex As String, ByVal nOptions As Long) As Long
Public Declare PtrSafe Function HASH_HexFromBits Lib "diCryptoSys.dll" (ByVal strOutput As String, ByVal nMaxChars As Long, ByRef lpData As Byte, ByVal nDataBitLen As Long, ByVal nOptions As Long) As Long
' Alias for VB6 strings
Public Declare PtrSafe Function HASH_HexFromString Lib "diCryptoSys.dll" Alias "HASH_HexFromBytes" (ByVal strOutput As String, ByVal nMaxChars As Long, ByVal strMessage As String, ByVal nStrLen As Long, ByVal nOptions As Long) As Long
' Stateful HASH functions added in [v6.0]
Public Declare PtrSafe Function HASH_Init Lib "diCryptoSys.dll" (ByVal nAlg As Long) As Long
Public Declare PtrSafe Function HASH_AddBytes Lib "diCryptoSys.dll" (ByVal hContext As Long, ByRef lpData As Byte, ByVal nDataLen As Long) As Long
Public Declare PtrSafe Function HASH_Final Lib "diCryptoSys.dll" (ByRef lpOutput As Byte, ByVal nOutLen As Long, ByVal hContext As Long) As Long
Public Declare PtrSafe Function HASH_DigestLength Lib "diCryptoSys.dll" (ByVal hContext As Long) As Long
Public Declare PtrSafe Function HASH_Reset Lib "diCryptoSys.dll" (ByVal hContext As Long) As Long

' GENERIC MAC FUNCTIONS (HMAC, CMAC, Poly1305 [v5.0], KMAC [v5.3])
Public Declare PtrSafe Function MAC_Bytes Lib "diCryptoSys.dll" (ByRef lpOutput As Byte, ByVal nOutLen As Long, ByRef lpMessage As Byte, ByVal nMsgLen As Long, ByRef lpKey As Byte, ByVal nKeyLen As Long, ByVal nOptions As Long) As Long
Public Declare PtrSafe Function MAC_HexFromBytes Lib "diCryptoSys.dll" (ByVal strOutput As String, ByVal nMaxChars As Long, ByRef lpMessage As Byte, ByVal nMsgLen As Long, ByRef lpKey As Byte, ByVal nKeyLen As Long, ByVal nOptions As Long) As Long
Public Declare PtrSafe Function MAC_HexFromHex Lib "diCryptoSys.dll" (ByVal strOutput As String, ByVal nMaxChars As Long, ByVal strMsgHex As String, ByVal strKeyHex As String, ByVal nOptions As Long) As Long
' Stateful MAC functions added in [v6.0] (HMAC only)
Public Declare PtrSafe Function MAC_Init Lib "diCryptoSys.dll" (ByRef lpKey As Byte, ByVal nKeyLen As Long, ByVal nAlg As Long) As Long
Public Declare PtrSafe Function MAC_AddBytes Lib "diCryptoSys.dll" (ByVal hContext As Long, ByRef lpMessage As Byte, ByVal nMsgLen As Long) As Long
Public Declare PtrSafe Function MAC_Final Lib "diCryptoSys.dll" (ByRef lpOutput As Byte, ByVal nOutLen As Long, ByVal hContext As Long) As Long
Public Declare PtrSafe Function MAC_CodeLength Lib "diCryptoSys.dll" (ByVal hContext As Long) As Long
Public Declare PtrSafe Function MAC_Reset Lib "diCryptoSys.dll" (ByVal hContext As Long) As Long

' SECURE HASH ALGORITHM 1 (SHA-1)
Public Declare PtrSafe Function SHA1_StringHexHash Lib "diCryptoSys.dll" (ByVal strDigest As String, ByVal strData As String) As Long
Public Declare PtrSafe Function SHA1_FileHexHash Lib "diCryptoSys.dll" (ByVal strDigest As String, ByVal strFileName As String, ByVal strMode As String) As Long
Public Declare PtrSafe Function SHA1_BytesHexHash Lib "diCryptoSys.dll" (ByVal strDigest As String, ByRef lpData As Byte, ByVal nBytes As Long) As Long
Public Declare PtrSafe Function SHA1_BytesHash Lib "diCryptoSys.dll" (ByRef lpDigest As Byte, ByRef lpData As Byte, ByVal nBytes As Long) As Long
Public Declare PtrSafe Function SHA1_Init Lib "diCryptoSys.dll" () As Long
Public Declare PtrSafe Function SHA1_AddString Lib "diCryptoSys.dll" (ByVal hContext As Long, ByVal strData As String) As Long
Public Declare PtrSafe Function SHA1_AddBytes Lib "diCryptoSys.dll" (ByVal hContext As Long, ByRef lpData As Byte, ByVal nBytes As Long) As Long
Public Declare PtrSafe Function SHA1_HexDigest Lib "diCryptoSys.dll" (ByVal strDigest As String, ByVal hContext As Long) As Long
Public Declare PtrSafe Function SHA1_Reset Lib "diCryptoSys.dll" (ByVal hContext As Long) As Long
Public Declare PtrSafe Function SHA1_Hmac Lib "diCryptoSys.dll" (ByVal strDigest As String, ByRef lpData As Byte, ByVal nBytes As Long, ByRef lpKey As Byte, ByVal nKeyBytes As Long) As Long
Public Declare PtrSafe Function SHA1_HmacHex Lib "diCryptoSys.dll" (ByVal strDigest As String, ByVal strHexData As String, ByVal strHexKey As String) As Long
    
' SECURE HASH ALGORITHM (SHA-256)
Public Declare PtrSafe Function SHA2_StringHexHash Lib "diCryptoSys.dll" (ByVal strDigest As String, ByVal strData As String) As Long
Public Declare PtrSafe Function SHA2_FileHexHash Lib "diCryptoSys.dll" (ByVal strDigest As String, ByVal strFileName As String, ByVal strMode As String) As Long
Public Declare PtrSafe Function SHA2_BytesHexHash Lib "diCryptoSys.dll" (ByVal strDigest As String, ByRef lpData As Byte, ByVal nBytes As Long) As Long
Public Declare PtrSafe Function SHA2_BytesHash Lib "diCryptoSys.dll" (ByRef lpDigest As Byte, ByRef lpData As Byte, ByVal nBytes As Long) As Long
Public Declare PtrSafe Function SHA2_Init Lib "diCryptoSys.dll" () As Long
Public Declare PtrSafe Function SHA2_AddString Lib "diCryptoSys.dll" (ByVal hContext As Long, ByVal strData As String) As Long
Public Declare PtrSafe Function SHA2_AddBytes Lib "diCryptoSys.dll" (ByVal hContext As Long, ByRef lpData As Byte, ByVal nBytes As Long) As Long
Public Declare PtrSafe Function SHA2_HexDigest Lib "diCryptoSys.dll" (ByVal strDigest As String, ByVal hContext As Long) As Long
Public Declare PtrSafe Function SHA2_Reset Lib "diCryptoSys.dll" (ByVal hContext As Long) As Long
Public Declare PtrSafe Function SHA2_Hmac Lib "diCryptoSys.dll" (ByVal strDigest As String, ByRef lpData As Byte, ByVal nBytes As Long, ByRef lpKey As Byte, ByVal nKeyBytes As Long) As Long
Public Declare PtrSafe Function SHA2_HmacHex Lib "diCryptoSys.dll" (ByVal strDigest As String, ByVal strHexData As String, ByVal strHexKey As String) As Long

' SECURE HASH ALGORITHM (SHA-3)
' New in [v5.3]
Public Declare PtrSafe Function SHA3_Init Lib "diCryptoSys.dll" (ByVal nHashBitLen As Long) As Long
Public Declare PtrSafe Function SHA3_AddString Lib "diCryptoSys.dll" (ByVal hContext As Long, ByVal strMessage As String) As Long
Public Declare PtrSafe Function SHA3_AddBytes Lib "diCryptoSys.dll" (ByVal hContext As Long, ByRef lpData As Byte, ByVal nDataLen As Long) As Long
Public Declare PtrSafe Function SHA3_HexDigest Lib "diCryptoSys.dll" (ByVal strOutput As String, ByVal nMaxChars As Long, ByVal hContext As Long) As Long
Public Declare PtrSafe Function SHA3_LengthInBytes Lib "diCryptoSys.dll" (ByVal hContext As Long) As Long
Public Declare PtrSafe Function SHA3_Reset Lib "diCryptoSys.dll" (ByVal hContext As Long) As Long

' RSA DATA SECURITY, INC. MD5 MESSAGE-DIGEST ALGORITHM
Public Declare PtrSafe Function MD5_StringHexHash Lib "diCryptoSys.dll" (ByVal strDigest As String, ByVal strData As String) As Long
Public Declare PtrSafe Function MD5_FileHexHash Lib "diCryptoSys.dll" (ByVal strDigest As String, ByVal strFileName As String, ByVal strMode As String) As Long
Public Declare PtrSafe Function MD5_BytesHexHash Lib "diCryptoSys.dll" (ByVal strDigest As String, ByRef lpData As Byte, ByVal nBytes As Long) As Long
Public Declare PtrSafe Function MD5_BytesHash Lib "diCryptoSys.dll" (ByRef lpDigest As Byte, ByRef lpData As Byte, ByVal nBytes As Long) As Long
Public Declare PtrSafe Function MD5_Init Lib "diCryptoSys.dll" () As Long
Public Declare PtrSafe Function MD5_AddString Lib "diCryptoSys.dll" (ByVal hContext As Long, ByVal strData As String) As Long
Public Declare PtrSafe Function MD5_AddBytes Lib "diCryptoSys.dll" (ByVal hContext As Long, ByRef lpData As Byte, ByVal nBytes As Long) As Long
Public Declare PtrSafe Function MD5_HexDigest Lib "diCryptoSys.dll" (ByVal strDigest As String, ByVal hContext As Long) As Long
Public Declare PtrSafe Function MD5_Reset Lib "diCryptoSys.dll" (ByVal hContext As Long) As Long
Public Declare PtrSafe Function MD5_Hmac Lib "diCryptoSys.dll" (ByVal strDigest As String, ByRef lpData As Byte, ByVal nBytes As Long, ByRef lpKey As Byte, ByVal nKeyBytes As Long) As Long
Public Declare PtrSafe Function MD5_HmacHex Lib "diCryptoSys.dll" (ByVal strDigest As String, ByVal strHexData As String, ByVal strHexKey As String) As Long
    
' RC4-COMPATIBLE PC1 FUNCTIONS (Superseded by CIPHER_Stream functions in [v5.0])
Public Declare PtrSafe Function PC1_Bytes Lib "diCryptoSys.dll" (ByRef lpOutput As Byte, ByRef lpData As Byte, ByVal nDataLen As Long, ByRef lpKey As Byte, ByVal nKeyLen As Long) As Long
Public Declare PtrSafe Function PC1_Hex Lib "diCryptoSys.dll" (ByVal strOutput As String, ByVal nMaxChars As Long, ByVal strInputHex As String, ByVal strKeyHex As String) As Long
Public Declare PtrSafe Function PC1_File Lib "diCryptoSys.dll" (ByVal strFileOut As String, ByVal strFileIn As String, ByRef lpKey As Byte, ByVal nKeyLen As Long) As Long

' RANDOM NUMBER GENERATOR (RNG) FUNCTIONS
Public Declare PtrSafe Function RNG_KeyBytes Lib "diCryptoSys.dll" (ByRef lpOutput As Byte, ByVal nBytes As Long, ByVal strSeed As String, ByVal nSeedLen As Long) As Long
Public Declare PtrSafe Function RNG_KeyHex Lib "diCryptoSys.dll" (ByVal strOutput As String, ByVal nMaxChars As Long, ByVal nBytes As Long, ByVal strSeed As String, ByVal nSeedLen As Long) As Long
Public Declare PtrSafe Function RNG_NonceData Lib "diCryptoSys.dll" (ByRef lpOutput As Byte, ByVal nBytes As Long) As Long
Public Declare PtrSafe Function RNG_NonceDataHex Lib "diCryptoSys.dll" (ByVal strOutput As String, ByVal nMaxChars As Long, ByVal nBytes As Long) As Long
Public Declare PtrSafe Function RNG_Test Lib "diCryptoSys.dll" (ByVal strFileName As String) As Long
Public Declare PtrSafe Function RNG_Number Lib "diCryptoSys.dll" (ByVal nLower As Long, ByVal nUpper As Long) As Long
Public Declare PtrSafe Function RNG_BytesWithPrompt Lib "diCryptoSys.dll" (ByRef lpOutput As Byte, ByVal nOutputLen As Long, ByVal strPrompt As String, ByVal nOptions As Long) As Long
Public Declare PtrSafe Function RNG_HexWithPrompt Lib "diCryptoSys.dll" (ByVal strOutput As String, ByVal nMaxChars As Long, ByVal nBytes As Long, ByVal strPrompt As String, ByVal nOptions As Long) As Long
Public Declare PtrSafe Function RNG_Initialize Lib "diCryptoSys.dll" (ByVal strSeedFile As String, ByVal nOptions As Long) As Long
Public Declare PtrSafe Function RNG_MakeSeedFile Lib "diCryptoSys.dll" (ByVal strSeedFile As String, ByVal strPrompt As String, ByVal nOptions As Long) As Long
Public Declare PtrSafe Function RNG_UpdateSeedFile Lib "diCryptoSys.dll" (ByVal strSeedFile As String, ByVal nOptions As Long) As Long
Public Declare PtrSafe Function RNG_TestDRBGVS Lib "diCryptoSys.dll" (ByVal strOutput As String, ByVal nMaxChars As Long, ByVal nReturnedBitsLen As Long, ByVal strEntropyInput As String, ByVal strNonce As String, ByVal strPersonalizationString As String, ByVal strAdditionalInput1 As String, ByVal strEntropyReseed As String, ByVal strAdditionalInputReseed As String, ByVal strAdditionalInput2 As String, ByVal nOptions As Long) As Long
    
' ZLIB COMPRESSION FUNCTIONS
Public Declare PtrSafe Function ZLIB_Deflate Lib "diCryptoSys.dll" (ByRef lpOutput As Byte, ByVal nOutputLen As Long, ByRef lpInput As Byte, ByVal nInputLen As Long) As Long
Public Declare PtrSafe Function ZLIB_Inflate Lib "diCryptoSys.dll" (ByRef lpOutput As Byte, ByVal nOutputLen As Long, ByRef lpInput As Byte, ByVal nInputLen As Long) As Long

' GENERIC COMPRESSION FUNCTIONS
' New in [v6.20]
Public Declare PtrSafe Function COMPR_Compress Lib "diCryptoSys.dll" (ByRef lpOutput As Byte, ByVal nOutBytes As Long, ByRef lpInput As Byte, ByVal nInputLen As Long, ByVal nOptions As Long) As Long
Public Declare PtrSafe Function COMPR_Uncompress Lib "diCryptoSys.dll" (ByRef lpOutput As Byte, ByVal nOutBytes As Long, ByRef lpInput As Byte, ByVal nInputLen As Long, ByVal nOptions As Long) As Long

' PASSWORD-BASED KEY DERIVATION FUNCTIONS
Public Declare PtrSafe Function PBE_Kdf2 Lib "diCryptoSys.dll" (ByRef lpDerivedKey As Byte, ByVal nKeyLen As Long, ByRef lpPwd As Byte, ByVal nPwdlen As Long, ByRef lpSalt As Byte, ByVal nSaltLen As Long, ByVal nCount As Long, ByVal nOptions As Long) As Long
Public Declare PtrSafe Function PBE_Kdf2Hex Lib "diCryptoSys.dll" (ByVal strOutput As String, ByVal nMaxChars As Long, ByVal nKeyBytes As Long, ByVal strPwd As String, ByVal strSaltHex As String, ByVal nCount As Long, ByVal nOptions As Long) As Long
' New in [v5.2]
Public Declare PtrSafe Function PBE_Scrypt Lib "diCryptoSys.dll" (ByRef lpDerivedKey As Byte, ByVal nKeyLen As Long, ByRef lpPwd As Byte, ByVal nPwdlen As Long, ByRef lpSalt As Byte, ByVal nSaltLen As Long, ByVal nParamN As Long, ByVal nParamR As Long, ByVal nParamP As Long, ByVal nOptions As Long) As Long
Public Declare PtrSafe Function PBE_ScryptHex Lib "diCryptoSys.dll" (ByVal strOutput As String, ByVal nMaxChars As Long, ByVal dkBytes As Long, ByVal strPwd As String, ByVal strSaltHex As String, ByVal nParamN As Long, ByVal nParamR As Long, ByVal nParamP As Long, ByVal nOptions As Long) As Long

' HEX ENCODING CONVERSION FUNCTIONS
' See cnvHexStrFromBytes, cnvBytesFromHexStr, cnvHexFilter below
Public Declare PtrSafe Function CNV_HexStrFromBytes Lib "diCryptoSys.dll" (ByVal strOutput As String, ByVal nMaxChars As Long, ByRef lpData As Byte, ByVal nBytes As Long) As Long
Public Declare PtrSafe Function CNV_BytesFromHexStr Lib "diCryptoSys.dll" (ByRef lpOutput As Byte, ByVal nOutLen As Long, ByVal strInput As String) As Long
Public Declare PtrSafe Function CNV_HexFilter Lib "diCryptoSys.dll" (ByVal strOutput As String, ByVal strInput As String, ByVal nStrLen As Long) As Long

' BASE64 ENCODING CONVERSION FUNCTIONS
' See cnvB64StrFromBytes, cnvBytesFromHexB64, cnvB64Filter below
Public Declare PtrSafe Function CNV_B64StrFromBytes Lib "diCryptoSys.dll" (ByVal strOutput As String, ByVal nMaxChars As Long, ByRef lpData As Byte, ByVal nBytes As Long) As Long
Public Declare PtrSafe Function CNV_BytesFromB64Str Lib "diCryptoSys.dll" (ByRef lpOutput As Byte, ByVal nOutLen As Long, ByVal strInput As String) As Long
Public Declare PtrSafe Function CNV_B64Filter Lib "diCryptoSys.dll" (ByVal strOutput As String, ByVal strInput As String, ByVal nStrLen As Long) As Long

' CRC FUNCTIONS
Public Declare PtrSafe Function CRC_Bytes Lib "diCryptoSys.dll" (ByRef lpInput As Byte, ByVal nBytes As Long, ByVal nOptions As Long) As Long
Public Declare PtrSafe Function CRC_String Lib "diCryptoSys.dll" (ByVal strInput As String, ByVal nOptions As Long) As Long
Public Declare PtrSafe Function CRC_File Lib "diCryptoSys.dll" (ByVal strFileName As String, ByVal nOptions As Long) As Long

' FUNCTIONS TO WIPE DATA
Public Declare PtrSafe Function WIPE_File Lib "diCryptoSys.dll" (ByVal strFileName As String, ByVal nOptions As Long) As Long
Public Declare PtrSafe Function WIPE_Data Lib "diCryptoSys.dll" (ByRef lpData As Byte, ByVal nBytes As Long) As Long
' Alternative Aliases to cope with Byte and String types explicitly...
Public Declare PtrSafe Function WIPE_Bytes Lib "diCryptoSys.dll" Alias "WIPE_Data" (ByRef lpData As Byte, ByVal nBytes As Long) As Long
Public Declare PtrSafe Function WIPE_String Lib "diCryptoSys.dll" Alias "WIPE_Data" (ByVal strData As String, ByVal nStrLen As Long) As Long

' PADDING FUNCTIONS
Public Declare PtrSafe Function PAD_BytesBlock Lib "diCryptoSys.dll" (ByRef lpOutput As Byte, ByVal nOutputLen As Long, ByRef lpInput As Byte, ByVal nInputLen As Long, ByVal nBlockLen As Long, ByVal nOptions As Long) As Long
Public Declare PtrSafe Function PAD_UnpadBytes Lib "diCryptoSys.dll" (ByRef lpOutput As Byte, ByVal nOutputLen As Long, ByRef lpInput As Byte, ByVal nInputLen As Long, ByVal nBlockLen As Long, ByVal nOptions As Long) As Long
Public Declare PtrSafe Function PAD_HexBlock Lib "diCryptoSys.dll" (ByVal strOutput As String, ByVal nMaxChars As Long, ByVal strInputHex As String, ByVal nBlockLen As Long, ByVal nOptions As Long) As Long
Public Declare PtrSafe Function PAD_UnpadHex Lib "diCryptoSys.dll" (ByVal strOutput As String, ByVal nMaxChars As Long, ByVal strInputHex As String, ByVal nBlockLen As Long, ByVal nOptions As Long) As Long

' XOF/PRF PROTOTYPES
Public Declare PtrSafe Function XOF_Bytes Lib "diCryptoSys.dll" (ByRef lpOutput As Byte, ByVal nOutBytes As Long, ByRef lpMessage As Byte, ByVal nMsgLen As Long, ByVal nOptions As Long) As Long
Public Declare PtrSafe Function PRF_Bytes Lib "diCryptoSys.dll" (ByRef lpOutput As Byte, ByVal nOutBytes As Long, ByRef lpMessage As Byte, ByVal nMsgLen As Long, ByRef lpKey As Byte, ByVal nKeyLen As Long, ByVal strCustom As String, ByVal nOptions As Long) As Long

#Else
' Declarations for VB6 and 32-bit Office

' GENERAL FUNCTIONS
Public Declare Function API_Version Lib "diCryptoSys.dll" () As Long
Public Declare Function API_LicenceType Lib "diCryptoSys.dll" (ByVal nReserved As Long) As Long
Public Declare Function API_CompileTime Lib "diCryptoSys.dll" (ByVal strCompiledOn As String, ByVal nMaxChars As Long) As Long
Public Declare Function API_ModuleName Lib "diCryptoSys.dll" (ByVal strModuleName As String, ByVal nMaxChars As Long, ByVal nReserved As Long) As Long
Public Declare Function API_ErrorCode Lib "diCryptoSys.dll" () As Long    ' Added [v4.2] (only for certain fns)
Public Declare Function API_ErrorLookup Lib "diCryptoSys.dll" (ByVal strErrMsg As String, ByVal nMaxChars As Long, ByVal nErrCode As Long) As Long
Public Declare Function API_PowerUpTests Lib "diCryptoSys.dll" (ByVal nReserved As Long) As Long

' ADVANCED ENCRYPTION STANDARD (AES) BLOCK CIPHER WITH 128-BIT KEY
Public Declare Function AES128_Bytes Lib "diCryptoSys.dll" (ByRef lpOutput As Byte, ByRef lpData As Byte, ByVal nBytes As Long, ByRef lpKey As Byte, ByVal fEncrypt As Long) As Long
Public Declare Function AES128_BytesMode Lib "diCryptoSys.dll" (ByRef lpOutput As Byte, ByRef lpData As Byte, ByVal nBytes As Long, ByRef lpKey As Byte, ByVal fEncrypt As Long, ByVal strMode As String, ByRef lpInitV As Byte) As Long
Public Declare Function AES128_Hex Lib "diCryptoSys.dll" (ByVal strOutput As String, ByVal strInput As String, ByVal strHexKey As String, ByVal fEncrypt As Long) As Long
Public Declare Function AES128_HexMode Lib "diCryptoSys.dll" (ByVal strOutput As String, ByVal strInput As String, ByVal strHexKey As String, ByVal fEncrypt As Long, ByVal strMode As String, ByVal strHexIV As String) As Long
Public Declare Function AES128_B64Mode Lib "diCryptoSys.dll" (ByVal strOutput As String, ByVal strInput As String, ByVal strB64Key As String, ByVal fEncrypt As Long, ByVal strMode As String, ByVal strB64IV As String) As Long
Public Declare Function AES128_File Lib "diCryptoSys.dll" (ByVal strFileOut As String, ByVal strFileIn As String, ByRef lpKey As Byte, ByVal fEncrypt As Long, ByVal strMode As String, ByRef lpInitV As Byte) As Long
Public Declare Function AES128_FileExt Lib "diCryptoSys.dll" (ByVal strFileOut As String, ByVal strFileIn As String, ByRef lpKey As Byte, ByVal fEncrypt As Long, ByVal strMode As String, ByRef lpInitV As Byte, ByVal nOptions As Long) As Long
Public Declare Function AES128_FileHex Lib "diCryptoSys.dll" (ByVal strFileOut As String, ByVal strFileIn As String, ByVal strHexKey As String, ByVal fEncrypt As Long, ByVal strMode As String, ByVal strHexIV As String) As Long
Public Declare Function AES128_Init Lib "diCryptoSys.dll" (ByRef lpKey As Byte, ByVal fEncrypt As Long, ByVal strMode As String, ByRef lpInitV As Byte) As Long
Public Declare Function AES128_InitHex Lib "diCryptoSys.dll" (ByVal strHexKey As String, ByVal fEncrypt As Long, ByVal strMode As String, ByVal strHexIV As String) As Long
Public Declare Function AES128_Update Lib "diCryptoSys.dll" (ByVal hContext As Long, ByRef lpData As Byte, ByVal nBytes As Long) As Long
Public Declare Function AES128_UpdateHex Lib "diCryptoSys.dll" (ByVal hContext As Long, ByVal strHexData As String) As Long
Public Declare Function AES128_Final Lib "diCryptoSys.dll" (ByVal hContext As Long) As Long
Public Declare Function AES128_InitError Lib "diCryptoSys.dll" () As Long

' ADVANCED ENCRYPTION STANDARD (AES) BLOCK CIPHER WITH 192-BIT KEY
Public Declare Function AES192_Bytes Lib "diCryptoSys.dll" (ByRef lpOutput As Byte, ByRef lpData As Byte, ByVal nBytes As Long, ByRef lpKey As Byte, ByVal fEncrypt As Long) As Long
Public Declare Function AES192_BytesMode Lib "diCryptoSys.dll" (ByRef lpOutput As Byte, ByRef lpData As Byte, ByVal nBytes As Long, ByRef lpKey As Byte, ByVal fEncrypt As Long, ByVal strMode As String, ByRef lpInitV As Byte) As Long
Public Declare Function AES192_Hex Lib "diCryptoSys.dll" (ByVal strOutput As String, ByVal strInput As String, ByVal strHexKey As String, ByVal fEncrypt As Long) As Long
Public Declare Function AES192_HexMode Lib "diCryptoSys.dll" (ByVal strOutput As String, ByVal strInput As String, ByVal strHexKey As String, ByVal fEncrypt As Long, ByVal strMode As String, ByVal strHexIV As String) As Long
Public Declare Function AES192_B64Mode Lib "diCryptoSys.dll" (ByVal strOutput As String, ByVal strInput As String, ByVal strB64Key As String, ByVal fEncrypt As Long, ByVal strMode As String, ByVal strB64IV As String) As Long
Public Declare Function AES192_File Lib "diCryptoSys.dll" (ByVal strFileOut As String, ByVal strFileIn As String, ByRef lpKey As Byte, ByVal fEncrypt As Long, ByVal strMode As String, ByRef lpInitV As Byte) As Long
Public Declare Function AES192_FileExt Lib "diCryptoSys.dll" (ByVal strFileOut As String, ByVal strFileIn As String, ByRef lpKey As Byte, ByVal fEncrypt As Long, ByVal strMode As String, ByRef lpInitV As Byte, ByVal nOptions As Long) As Long
Public Declare Function AES192_FileHex Lib "diCryptoSys.dll" (ByVal strFileOut As String, ByVal strFileIn As String, ByVal strHexKey As String, ByVal fEncrypt As Long, ByVal strMode As String, ByVal strHexIV As String) As Long
Public Declare Function AES192_Init Lib "diCryptoSys.dll" (ByRef lpKey As Byte, ByVal fEncrypt As Long, ByVal strMode As String, ByRef lpInitV As Byte) As Long
Public Declare Function AES192_InitHex Lib "diCryptoSys.dll" (ByVal strHexKey As String, ByVal fEncrypt As Long, ByVal strMode As String, ByVal strHexIV As String) As Long
Public Declare Function AES192_Update Lib "diCryptoSys.dll" (ByVal hContext As Long, ByRef lpData As Byte, ByVal nBytes As Long) As Long
Public Declare Function AES192_UpdateHex Lib "diCryptoSys.dll" (ByVal hContext As Long, ByVal strHexData As String) As Long
Public Declare Function AES192_Final Lib "diCryptoSys.dll" (ByVal hContext As Long) As Long
Public Declare Function AES192_InitError Lib "diCryptoSys.dll" () As Long

' ADVANCED ENCRYPTION STANDARD (AES) BLOCK CIPHER WITH 256-BIT KEY
Public Declare Function AES256_Bytes Lib "diCryptoSys.dll" (ByRef lpOutput As Byte, ByRef lpData As Byte, ByVal nBytes As Long, ByRef lpKey As Byte, ByVal fEncrypt As Long) As Long
Public Declare Function AES256_BytesMode Lib "diCryptoSys.dll" (ByRef lpOutput As Byte, ByRef lpData As Byte, ByVal nBytes As Long, ByRef lpKey As Byte, ByVal fEncrypt As Long, ByVal strMode As String, ByRef lpInitV As Byte) As Long
Public Declare Function AES256_Hex Lib "diCryptoSys.dll" (ByVal strOutput As String, ByVal strInput As String, ByVal strHexKey As String, ByVal fEncrypt As Long) As Long
Public Declare Function AES256_HexMode Lib "diCryptoSys.dll" (ByVal strOutput As String, ByVal strInput As String, ByVal strHexKey As String, ByVal fEncrypt As Long, ByVal strMode As String, ByVal strHexIV As String) As Long
Public Declare Function AES256_B64Mode Lib "diCryptoSys.dll" (ByVal strOutput As String, ByVal strInput As String, ByVal strB64Key As String, ByVal fEncrypt As Long, ByVal strMode As String, ByVal strB64IV As String) As Long
Public Declare Function AES256_File Lib "diCryptoSys.dll" (ByVal strFileOut As String, ByVal strFileIn As String, ByRef lpKey As Byte, ByVal fEncrypt As Long, ByVal strMode As String, ByRef lpInitV As Byte) As Long
Public Declare Function AES256_FileExt Lib "diCryptoSys.dll" (ByVal strFileOut As String, ByVal strFileIn As String, ByRef lpKey As Byte, ByVal fEncrypt As Long, ByVal strMode As String, ByRef lpInitV As Byte, ByVal nOptions As Long) As Long
Public Declare Function AES256_FileHex Lib "diCryptoSys.dll" (ByVal strFileOut As String, ByVal strFileIn As String, ByVal strHexKey As String, ByVal fEncrypt As Long, ByVal strMode As String, ByVal strHexIV As String) As Long
Public Declare Function AES256_Init Lib "diCryptoSys.dll" (ByRef lpKey As Byte, ByVal fEncrypt As Long, ByVal strMode As String, ByRef lpInitV As Byte) As Long
Public Declare Function AES256_InitHex Lib "diCryptoSys.dll" (ByVal strHexKey As String, ByVal fEncrypt As Long, ByVal strMode As String, ByVal strHexIV As String) As Long
Public Declare Function AES256_Update Lib "diCryptoSys.dll" (ByVal hContext As Long, ByRef lpData As Byte, ByVal nBytes As Long) As Long
Public Declare Function AES256_UpdateHex Lib "diCryptoSys.dll" (ByVal hContext As Long, ByVal strHexData As String) As Long
Public Declare Function AES256_Final Lib "diCryptoSys.dll" (ByVal hContext As Long) As Long
Public Declare Function AES256_InitError Lib "diCryptoSys.dll" () As Long

' BLOWFISH BLOCK CIPHER FUNCTIONS
Public Declare Function BLF_Bytes Lib "diCryptoSys.dll" (ByRef lpOutput As Byte, ByRef lpData As Byte, ByVal nBytes As Long, ByRef lpKey As Byte, ByVal nKeyBytes As Long, ByVal fEncrypt As Long) As Long
Public Declare Function BLF_BytesMode Lib "diCryptoSys.dll" (ByRef lpOutput As Byte, ByRef lpData As Byte, ByVal nBytes As Long, ByRef lpKey As Byte, ByVal nKeyBytes As Long, ByVal fEncrypt As Long, ByVal strMode As String, ByRef lpInitV As Byte) As Long
Public Declare Function BLF_Hex Lib "diCryptoSys.dll" (ByVal strOutput As String, ByVal strInput As String, ByVal strHexKey As String, ByVal fEncrypt As Long) As Long
Public Declare Function BLF_HexMode Lib "diCryptoSys.dll" (ByVal strOutput As String, ByVal strInput As String, ByVal strHexKey As String, ByVal fEncrypt As Long, ByVal strMode As String, ByVal strHexIV As String) As Long
Public Declare Function BLF_B64Mode Lib "diCryptoSys.dll" (ByVal strOutput As String, ByVal strInput As String, ByVal strB64Key As String, ByVal fEncrypt As Long, ByVal strMode As String, ByVal strB64IV As String) As Long
Public Declare Function BLF_File Lib "diCryptoSys.dll" (ByVal strFileOut As String, ByVal strFileIn As String, ByRef lpKey As Byte, ByVal nKeyBytes As Long, ByVal fEncrypt As Long, ByVal strMode As String, ByRef lpInitV As Byte) As Long
Public Declare Function BLF_FileExt Lib "diCryptoSys.dll" (ByVal strFileOut As String, ByVal strFileIn As String, ByRef lpKey As Byte, ByVal nKeyBytes As Long, ByVal fEncrypt As Long, ByVal strMode As String, ByRef lpInitV As Byte, ByVal nOptions As Long) As Long
Public Declare Function BLF_FileHex Lib "diCryptoSys.dll" (ByVal strFileOut As String, ByVal strFileIn As String, ByVal strHexKey As String, ByVal fEncrypt As Long, ByVal strMode As String, ByVal strHexIV As String) As Long
Public Declare Function BLF_InitHex Lib "diCryptoSys.dll" (ByVal strHexKey As String, ByVal fEncrypt As Long, ByVal strMode As String, ByVal strHexIV As String) As Long
Public Declare Function BLF_Init Lib "diCryptoSys.dll" (ByRef lpKey As Byte, ByVal nKeyBytes As Long, ByVal fEncrypt As Long, ByVal strMode As String, ByRef lpInitV As Byte) As Long
Public Declare Function BLF_UpdateHex Lib "diCryptoSys.dll" (ByVal hContext As Long, ByVal strHexData As String) As Long
Public Declare Function BLF_Update Lib "diCryptoSys.dll" (ByVal hContext As Long, ByRef lpData As Byte, ByVal nBytes As Long) As Long
Public Declare Function BLF_Final Lib "diCryptoSys.dll" (ByVal hContext As Long) As Long
Public Declare Function BLF_InitError Lib "diCryptoSys.dll" () As Long
    
' DATA ENCRYPTION STANDARD (DES) BLOCK CIPHER
Public Declare Function DES_Bytes Lib "diCryptoSys.dll" (ByRef lpOutput As Byte, ByRef lpData As Byte, ByVal nBytes As Long, ByRef lpKey As Byte, ByVal fEncrypt As Long) As Long
Public Declare Function DES_BytesMode Lib "diCryptoSys.dll" (ByRef lpOutput As Byte, ByRef lpData As Byte, ByVal nBytes As Long, ByRef lpKey As Byte, ByVal fEncrypt As Long, ByVal strMode As String, ByRef lpInitV As Byte) As Long
Public Declare Function DES_Hex Lib "diCryptoSys.dll" (ByVal strOutput As String, ByVal strInput As String, ByVal strHexKey As String, ByVal fEncrypt As Long) As Long
Public Declare Function DES_HexMode Lib "diCryptoSys.dll" (ByVal strOutput As String, ByVal strInput As String, ByVal strHexKey As String, ByVal fEncrypt As Long, ByVal strMode As String, ByVal strHexIV As String) As Long
Public Declare Function DES_B64Mode Lib "diCryptoSys.dll" (ByVal strOutput As String, ByVal strInput As String, ByVal strB64Key As String, ByVal fEncrypt As Long, ByVal strMode As String, ByVal strB64IV As String) As Long
Public Declare Function DES_File Lib "diCryptoSys.dll" (ByVal strFileOut As String, ByVal strFileIn As String, ByRef lpKey As Byte, ByVal fEncrypt As Long, ByVal strMode As String, ByRef lpInitV As Byte) As Long
Public Declare Function DES_FileExt Lib "diCryptoSys.dll" (ByVal strFileOut As String, ByVal strFileIn As String, ByRef lpKey As Byte, ByVal fEncrypt As Long, ByVal strMode As String, ByRef lpInitV As Byte, ByVal nOptions As Long) As Long
Public Declare Function DES_FileHex Lib "diCryptoSys.dll" (ByVal strFileOut As String, ByVal strFileIn As String, ByVal strHexKey As String, ByVal fEncrypt As Long, ByVal strMode As String, ByVal strHexIV As String) As Long
Public Declare Function DES_Init Lib "diCryptoSys.dll" (ByRef lpKey As Byte, ByVal fEncrypt As Long, ByVal strMode As String, ByRef lpInitV As Byte) As Long
Public Declare Function DES_InitHex Lib "diCryptoSys.dll" (ByVal strHexKey As String, ByVal fEncrypt As Long, ByVal strMode As String, ByVal strHexIV As String) As Long
Public Declare Function DES_Update Lib "diCryptoSys.dll" (ByVal hContext As Long, ByRef lpData As Byte, ByVal nBytes As Long) As Long
Public Declare Function DES_UpdateHex Lib "diCryptoSys.dll" (ByVal hContext As Long, ByVal strHexData As String) As Long
Public Declare Function DES_Final Lib "diCryptoSys.dll" (ByVal hContext As Long) As Long
Public Declare Function DES_InitError Lib "diCryptoSys.dll" () As Long
    
' Checks for weak or invalid-length DES or TDEA keys
Public Declare Function DES_CheckKey Lib "diCryptoSys.dll" (ByRef lpKey As Byte, ByVal nKeyBytes As Long) As Long
Public Declare Function DES_CheckKeyHex Lib "diCryptoSys.dll" (ByVal strHexKey As String) As Long

' TRIPLE DATA ENCRYPTION ALGORITHM (TDEA) BLOCK CIPHER
Public Declare Function TDEA_Bytes Lib "diCryptoSys.dll" (ByRef lpOutput As Byte, ByRef lpData As Byte, ByVal nBytes As Long, ByRef lpKey As Byte, ByVal fEncrypt As Long) As Long
Public Declare Function TDEA_BytesMode Lib "diCryptoSys.dll" (ByRef lpOutput As Byte, ByRef lpData As Byte, ByVal nBytes As Long, ByRef lpKey As Byte, ByVal fEncrypt As Long, ByVal strMode As String, ByRef lpInitV As Byte) As Long
Public Declare Function TDEA_Hex Lib "diCryptoSys.dll" (ByVal strOutput As String, ByVal strInput As String, ByVal strHexKey As String, ByVal fEncrypt As Long) As Long
Public Declare Function TDEA_HexMode Lib "diCryptoSys.dll" (ByVal strOutput As String, ByVal strInput As String, ByVal strHexKey As String, ByVal fEncrypt As Long, ByVal strMode As String, ByVal strHexIV As String) As Long
Public Declare Function TDEA_B64Mode Lib "diCryptoSys.dll" (ByVal strOutput As String, ByVal strInput As String, ByVal strB64Key As String, ByVal fEncrypt As Long, ByVal strMode As String, ByVal strB64IV As String) As Long
Public Declare Function TDEA_File Lib "diCryptoSys.dll" (ByVal strFileOut As String, ByVal strFileIn As String, ByRef lpKey As Byte, ByVal fEncrypt As Long, ByVal strMode As String, ByRef lpInitV As Byte) As Long
Public Declare Function TDEA_FileExt Lib "diCryptoSys.dll" (ByVal strFileOut As String, ByVal strFileIn As String, ByRef lpKey As Byte, ByVal fEncrypt As Long, ByVal strMode As String, ByRef lpInitV As Byte, ByVal nOptions As Long) As Long
Public Declare Function TDEA_FileHex Lib "diCryptoSys.dll" (ByVal strFileOut As String, ByVal strFileIn As String, ByVal strHexKey As String, ByVal fEncrypt As Long, ByVal strMode As String, ByVal strHexIV As String) As Long
Public Declare Function TDEA_Init Lib "diCryptoSys.dll" (ByRef lpKey As Byte, ByVal fEncrypt As Long, ByVal strMode As String, ByRef lpInitV As Byte) As Long
Public Declare Function TDEA_InitHex Lib "diCryptoSys.dll" (ByVal strHexKey As String, ByVal fEncrypt As Long, ByVal strMode As String, ByVal strHexIV As String) As Long
Public Declare Function TDEA_Update Lib "diCryptoSys.dll" (ByVal hContext As Long, ByRef lpData As Byte, ByVal nBytes As Long) As Long
Public Declare Function TDEA_UpdateHex Lib "diCryptoSys.dll" (ByVal hContext As Long, ByVal strHexData As String) As Long
Public Declare Function TDEA_Final Lib "diCryptoSys.dll" (ByVal hContext As Long) As Long
Public Declare Function TDEA_InitError Lib "diCryptoSys.dll" () As Long

' GENERIC BLOCK CIPHER FUNCTIONS
' Added in [v6.20] (to get rid of that annoying 2)
Public Declare Function CIPHER_EncryptBytes Lib "diCryptoSys.dll" (ByRef lpOutput As Byte, ByVal nOutBytes As Long, ByRef lpInput As Byte, ByVal nInputLen As Long, ByRef lpKey As Byte, ByVal nKeyLen As Long, ByRef lpIV As Byte, ByVal nIvLen As Long, ByVal strAlgModePad As String, ByVal nOptions As Long) As Long
Public Declare Function CIPHER_DecryptBytes Lib "diCryptoSys.dll" (ByRef lpOutput As Byte, ByVal nOutBytes As Long, ByRef lpInput As Byte, ByVal nInputLen As Long, ByRef lpKey As Byte, ByVal nKeyLen As Long, ByRef lpIV As Byte, ByVal nIvLen As Long, ByVal strAlgModePad As String, ByVal nOptions As Long) As Long
' Aliases for backwards compatibility
Public Declare Function CIPHER_EncryptBytes2 Lib "diCryptoSys.dll" Alias "CIPHER_EncryptBytes" (ByRef lpOutput As Byte, ByVal nOutBytes As Long, ByRef lpInput As Byte, ByVal nInputLen As Long, ByRef lpKey As Byte, ByVal nKeyLen As Long, ByRef lpIV As Byte, ByVal nIvLen As Long, ByVal strAlgModePad As String, ByVal nOptions As Long) As Long
Public Declare Function CIPHER_DecryptBytes2 Lib "diCryptoSys.dll" Alias "CIPHER_DecryptBytes" (ByRef lpOutput As Byte, ByVal nOutBytes As Long, ByRef lpInput As Byte, ByVal nInputLen As Long, ByRef lpKey As Byte, ByVal nKeyLen As Long, ByRef lpIV As Byte, ByVal nIvLen As Long, ByVal strAlgModePad As String, ByVal nOptions As Long) As Long
Public Declare Function CIPHER_FileEncrypt Lib "diCryptoSys.dll" (ByVal strFileOut As String, ByVal strFileIn As String, ByRef lpKey As Byte, ByVal nKeyLen As Long, ByRef lpIV As Byte, ByVal nIvLen As Long, ByVal strAlgModePad As String, ByVal nOptions As Long) As Long
Public Declare Function CIPHER_FileDecrypt Lib "diCryptoSys.dll" (ByVal strFileOut As String, ByVal strFileIn As String, ByRef lpKey As Byte, ByVal nKeyLen As Long, ByRef lpIV As Byte, ByVal nIvLen As Long, ByVal strAlgModePad As String, ByVal nOptions As Long) As Long
' New in [v6.0]
Public Declare Function CIPHER_EncryptHex Lib "diCryptoSys.dll" (ByVal strOutput As String, ByVal nOutChars As Long, ByVal strInputHex As String, ByVal strKeyHex As String, ByVal strIvHex As String, ByVal strAlgModePad As String, ByVal nOptions As Long) As Long
Public Declare Function CIPHER_DecryptHex Lib "diCryptoSys.dll" (ByVal strOutput As String, ByVal nOutChars As Long, ByVal strInputHex As String, ByVal strKeyHex As String, ByVal strIvHex As String, ByVal strAlgModePad As String, ByVal nOptions As Long) As Long
' Stateful CIPHER functions added in [v6.0]
Public Declare Function CIPHER_Init Lib "diCryptoSys.dll" (ByVal fEncrypt As Integer, ByVal strAlgAndMode As String, ByRef lpKey As Byte, ByVal nKeyLen As Long, ByRef lpIV As Byte, ByVal nIvLen As Long, ByVal nOptions As Long) As Long
Public Declare Function CIPHER_InitHex Lib "diCryptoSys.dll" (ByVal fEncrypt As Integer, ByVal strAlgAndMode As String, ByVal strKeyHex As String, ByVal strIvHex As String, ByVal nOptions As Long) As Long
Public Declare Function CIPHER_Update Lib "diCryptoSys.dll" (ByVal hContext As Long, ByRef lpOutput As Byte, ByVal nOutBytes As Long, ByRef lpData As Byte, ByVal nDataLen As Long) As Long
Public Declare Function CIPHER_UpdateHex Lib "diCryptoSys.dll" (ByVal hContext As Long, ByVal strOutput As String, ByVal nOutChars As Long, ByVal strDataHex As String) As Long
Public Declare Function CIPHER_Final Lib "diCryptoSys.dll" (ByVal hContext As Long) As Long

' KEY WRAP FUNCTIONS
Public Declare Function CIPHER_KeyWrap Lib "diCryptoSys.dll" (ByRef lpOutput As Byte, ByVal nOutBytes As Long, ByRef lpData As Byte, ByVal nDataLen As Long, ByRef lpKek As Byte, ByVal nKekLen As Long, ByVal nOptions As Long) As Long
Public Declare Function CIPHER_KeyUnwrap Lib "diCryptoSys.dll" (ByRef lpOutput As Byte, ByVal nOutBytes As Long, ByRef lpData As Byte, ByVal nDataLen As Long, ByRef lpKek As Byte, ByVal nKekLen As Long, ByVal nOptions As Long) As Long

' STREAM CIPHER FUNCTIONS
Public Declare Function CIPHER_StreamBytes Lib "diCryptoSys.dll" (ByRef lpOutput As Byte, ByRef lpData As Byte, ByVal nDataLen As Long, ByRef lpKey As Byte, ByVal nKeyLen As Long, ByRef lpIV As Byte, ByVal nIvLen As Long, ByVal nCounter As Long, ByVal nOptions As Long) As Long
Public Declare Function CIPHER_StreamHex Lib "diCryptoSys.dll" (ByVal strOutput As String, ByVal nOutChars As Long, ByVal strInputHex As String, ByVal strKeyHex As String, ByVal strIvHex As String, ByVal nCounter As Long, ByVal nOptions As Long) As Long
Public Declare Function CIPHER_StreamFile Lib "diCryptoSys.dll" (ByVal strFileOut As String, ByVal strFileIn As String, ByRef lpKey As Byte, ByVal nKeyLen As Long, ByRef lpIV As Byte, ByVal nIvLen As Long, ByVal nCounter As Long, ByVal nOptions As Long) As Long
Public Declare Function CIPHER_StreamInit Lib "diCryptoSys.dll" (ByRef lpKey As Byte, ByVal nKeyLen As Long, ByRef lpIV As Byte, ByVal nIvLen As Long, ByVal nCounter As Long, ByVal nOptions As Long) As Long
Public Declare Function CIPHER_StreamUpdate Lib "diCryptoSys.dll" (ByVal hContext As Long, ByRef lpOutput As Byte, ByRef lpData As Byte, ByVal nDataLen As Long) As Long
Public Declare Function CIPHER_StreamFinal Lib "diCryptoSys.dll" (ByVal hContext As Long) As Long

' AEAD FUNCTIONS
Public Declare Function AEAD_Encrypt Lib "diCryptoSys.dll" (ByRef lpOutput As Byte, ByVal nOutLen As Long, ByRef lpTagOut As Byte, ByVal nTagLen As Long, ByRef lpData As Byte, ByVal nDataLen As Long, ByRef lpKey As Byte, ByVal nKeyLen As Long, ByRef lpNonce As Byte, ByVal nNonceLen As Long, ByRef lpAAD As Byte, ByVal nAadLen As Long, ByVal nOptions As Long) As Long
Public Declare Function AEAD_Decrypt Lib "diCryptoSys.dll" (ByRef lpOutput As Byte, ByVal nOutLen As Long, ByRef lpData As Byte, ByVal nDataLen As Long, ByRef lpKey As Byte, ByVal nKeyLen As Long, ByRef lpNonce As Byte, ByVal nNonceLen As Long, ByRef lpAAD As Byte, ByVal nAadLen As Long, ByRef lpTag As Byte, ByVal nTagLen As Long, ByVal nOptions As Long) As Long
Public Declare Function AEAD_InitKey Lib "diCryptoSys.dll" (ByRef lpKey As Byte, ByVal nKeyLen As Long, ByVal nOptions As Long) As Long
Public Declare Function AEAD_SetNonce Lib "diCryptoSys.dll" (ByVal hContext As Long, ByRef lpNonce As Byte, ByVal nNonceLen As Long) As Long
Public Declare Function AEAD_AddAAD Lib "diCryptoSys.dll" (ByVal hContext As Long, ByRef lpAAD As Byte, ByVal nAadLen As Long) As Long
Public Declare Function AEAD_StartEncrypt Lib "diCryptoSys.dll" (ByVal hContext As Long) As Long
Public Declare Function AEAD_StartDecrypt Lib "diCryptoSys.dll" (ByVal hContext As Long, ByRef lpTagToCheck As Byte, ByVal nTagLen As Long) As Long
Public Declare Function AEAD_Update Lib "diCryptoSys.dll" (ByVal hContext As Long, ByRef lpOutput As Byte, ByVal nOutLen As Long, ByRef lpData As Byte, ByVal nDataLen As Long) As Long
Public Declare Function AEAD_FinishEncrypt Lib "diCryptoSys.dll" (ByVal hContext As Long, ByRef lpTagOut As Byte, ByVal nTagLen As Long) As Long
Public Declare Function AEAD_FinishDecrypt Lib "diCryptoSys.dll" (ByVal hContext As Long) As Long
Public Declare Function AEAD_Destroy Lib "diCryptoSys.dll" (ByVal hContext As Long) As Long
' Added in [v5.4]
Public Declare Function AEAD_EncryptWithTag Lib "diCryptoSys.dll" (ByRef lpOutput As Byte, ByVal nOutLen As Long, ByRef lpData As Byte, ByVal nDataLen As Long, ByRef lpKey As Byte, ByVal nKeyLen As Long, ByRef lpNonce As Byte, ByVal nNonceLen As Long, ByRef lpAAD As Byte, ByVal nAadLen As Long, ByVal nOptions As Long) As Long
Public Declare Function AEAD_DecryptWithTag Lib "diCryptoSys.dll" (ByRef lpOutput As Byte, ByVal nOutLen As Long, ByRef lpData As Byte, ByVal nDataLen As Long, ByRef lpKey As Byte, ByVal nKeyLen As Long, ByRef lpNonce As Byte, ByVal nNonceLen As Long, ByRef lpAAD As Byte, ByVal nAadLen As Long, ByVal nOptions As Long) As Long

' GCM AUTHENTICATED EN/DECRYPTION FUNCTIONS
' Partly superseded by AEAD functions in [v5.1]
Public Declare Function GCM_Encrypt Lib "diCryptoSys.dll" (ByRef lpOutput As Byte, ByVal nOutLen As Long, ByRef lpTagOut As Byte, ByVal nTagLen As Long, ByRef lpData As Byte, ByVal nDataLen As Long, ByRef lpKey As Byte, ByVal nKeyLen As Long, ByRef lpIV As Byte, ByVal nIvLen As Long, ByRef lpAAD As Byte, ByVal nAadLen As Long, ByVal nOptions As Long) As Long
Public Declare Function GCM_Decrypt Lib "diCryptoSys.dll" (ByRef lpOutput As Byte, ByVal nOutLen As Long, ByRef lpData As Byte, ByVal nDataLen As Long, ByRef lpKey As Byte, ByVal nKeyLen As Long, ByRef lpIV As Byte, ByVal nIvLen As Long, ByRef lpAAD As Byte, ByVal nAadLen As Long, ByRef lpTag As Byte, ByVal nTagLen As Long, ByVal nOptions As Long) As Long
Public Declare Function GCM_InitKey Lib "diCryptoSys.dll" (ByRef lpKey As Byte, ByVal nKeyLen As Long, ByVal nOptions As Long) As Long
Public Declare Function GCM_NextEncrypt Lib "diCryptoSys.dll" (ByVal hContext As Long, ByRef lpOutput As Byte, ByVal nOutLen As Long, ByRef lpTagOut As Byte, ByVal nTagLen As Long, ByRef lpData As Byte, ByVal nDataLen As Long, ByRef lpIV As Byte, ByVal nIvLen As Long, ByRef lpAAD As Byte, ByVal nAadLen As Long) As Long
Public Declare Function GCM_NextDecrypt Lib "diCryptoSys.dll" (ByVal hContext As Long, ByRef lpOutput As Byte, ByVal nOutLen As Long, ByRef lpData As Byte, ByVal nDataLen As Long, ByRef lpIV As Byte, ByVal nIvLen As Long, ByRef lpAAD As Byte, ByVal nAadLen As Long, ByRef lpTag As Byte, ByVal nTagLen As Long) As Long
Public Declare Function GCM_FinishKey Lib "diCryptoSys.dll" (ByVal hContext As Long) As Long

' GENERIC MESSAGE DIGEST HASH FUNCTIONS
Public Declare Function HASH_Bytes Lib "diCryptoSys.dll" (ByRef lpOutput As Byte, ByVal nOutLen As Long, ByRef lpMessage As Byte, ByVal nMsgLen As Long, ByVal nOptions As Long) As Long
Public Declare Function HASH_File Lib "diCryptoSys.dll" (ByRef lpOutput As Byte, ByVal nOutLen As Long, ByVal strFileName As String, ByVal nOptions As Long) As Long
Public Declare Function HASH_HexFromBytes Lib "diCryptoSys.dll" (ByVal strOutput As String, ByVal nMaxChars As Long, ByRef lpMessage As Byte, ByVal nMsgLen As Long, ByVal nOptions As Long) As Long
Public Declare Function HASH_HexFromFile Lib "diCryptoSys.dll" (ByVal strOutput As String, ByVal nMaxChars As Long, ByVal strFileName As String, ByVal nOptions As Long) As Long
Public Declare Function HASH_HexFromHex Lib "diCryptoSys.dll" (ByVal strOutput As String, ByVal nMaxChars As Long, ByVal strMsgHex As String, ByVal nOptions As Long) As Long
Public Declare Function HASH_HexFromBits Lib "diCryptoSys.dll" (ByVal strOutput As String, ByVal nMaxChars As Long, ByRef lpData As Byte, ByVal nDataBitLen As Long, ByVal nOptions As Long) As Long
' Alias for VB6 strings
Public Declare Function HASH_HexFromString Lib "diCryptoSys.dll" Alias "HASH_HexFromBytes" (ByVal strOutput As String, ByVal nMaxChars As Long, ByVal strMessage As String, ByVal nStrLen As Long, ByVal nOptions As Long) As Long
' Stateful HASH functions added in [v6.0]
Public Declare Function HASH_Init Lib "diCryptoSys.dll" (ByVal nAlg As Long) As Long
Public Declare Function HASH_AddBytes Lib "diCryptoSys.dll" (ByVal hContext As Long, ByRef lpData As Byte, ByVal nDataLen As Long) As Long
Public Declare Function HASH_Final Lib "diCryptoSys.dll" (ByRef lpOutput As Byte, ByVal nOutLen As Long, ByVal hContext As Long) As Long
Public Declare Function HASH_DigestLength Lib "diCryptoSys.dll" (ByVal hContext As Long) As Long
Public Declare Function HASH_Reset Lib "diCryptoSys.dll" (ByVal hContext As Long) As Long

' GENERIC MAC FUNCTIONS (HMAC, CMAC, Poly1305 [v5.0], KMAC [v5.3])
Public Declare Function MAC_Bytes Lib "diCryptoSys.dll" (ByRef lpOutput As Byte, ByVal nOutLen As Long, ByRef lpMessage As Byte, ByVal nMsgLen As Long, ByRef lpKey As Byte, ByVal nKeyLen As Long, ByVal nOptions As Long) As Long
Public Declare Function MAC_HexFromBytes Lib "diCryptoSys.dll" (ByVal strOutput As String, ByVal nMaxChars As Long, ByRef lpMessage As Byte, ByVal nMsgLen As Long, ByRef lpKey As Byte, ByVal nKeyLen As Long, ByVal nOptions As Long) As Long
Public Declare Function MAC_HexFromHex Lib "diCryptoSys.dll" (ByVal strOutput As String, ByVal nMaxChars As Long, ByVal strMsgHex As String, ByVal strKeyHex As String, ByVal nOptions As Long) As Long
' Stateful MAC functions added in [v6.0] (HMAC only)
Public Declare Function MAC_Init Lib "diCryptoSys.dll" (ByRef lpKey As Byte, ByVal nKeyLen As Long, ByVal nAlg As Long) As Long
Public Declare Function MAC_AddBytes Lib "diCryptoSys.dll" (ByVal hContext As Long, ByRef lpMessage As Byte, ByVal nMsgLen As Long) As Long
Public Declare Function MAC_Final Lib "diCryptoSys.dll" (ByRef lpOutput As Byte, ByVal nOutLen As Long, ByVal hContext As Long) As Long
Public Declare Function MAC_CodeLength Lib "diCryptoSys.dll" (ByVal hContext As Long) As Long
Public Declare Function MAC_Reset Lib "diCryptoSys.dll" (ByVal hContext As Long) As Long

' SECURE HASH ALGORITHM 1 (SHA-1)
Public Declare Function SHA1_StringHexHash Lib "diCryptoSys.dll" (ByVal strDigest As String, ByVal strData As String) As Long
Public Declare Function SHA1_FileHexHash Lib "diCryptoSys.dll" (ByVal strDigest As String, ByVal strFileName As String, ByVal strMode As String) As Long
Public Declare Function SHA1_BytesHexHash Lib "diCryptoSys.dll" (ByVal strDigest As String, ByRef lpData As Byte, ByVal nBytes As Long) As Long
Public Declare Function SHA1_BytesHash Lib "diCryptoSys.dll" (ByRef lpDigest As Byte, ByRef lpData As Byte, ByVal nBytes As Long) As Long
Public Declare Function SHA1_Init Lib "diCryptoSys.dll" () As Long
Public Declare Function SHA1_AddString Lib "diCryptoSys.dll" (ByVal hContext As Long, ByVal strData As String) As Long
Public Declare Function SHA1_AddBytes Lib "diCryptoSys.dll" (ByVal hContext As Long, ByRef lpData As Byte, ByVal nBytes As Long) As Long
Public Declare Function SHA1_HexDigest Lib "diCryptoSys.dll" (ByVal strDigest As String, ByVal hContext As Long) As Long
Public Declare Function SHA1_Reset Lib "diCryptoSys.dll" (ByVal hContext As Long) As Long
Public Declare Function SHA1_Hmac Lib "diCryptoSys.dll" (ByVal strDigest As String, ByRef lpData As Byte, ByVal nBytes As Long, ByRef lpKey As Byte, ByVal nKeyBytes As Long) As Long
Public Declare Function SHA1_HmacHex Lib "diCryptoSys.dll" (ByVal strDigest As String, ByVal strHexData As String, ByVal strHexKey As String) As Long
    
' SECURE HASH ALGORITHM (SHA-256)
Public Declare Function SHA2_StringHexHash Lib "diCryptoSys.dll" (ByVal strDigest As String, ByVal strData As String) As Long
Public Declare Function SHA2_FileHexHash Lib "diCryptoSys.dll" (ByVal strDigest As String, ByVal strFileName As String, ByVal strMode As String) As Long
Public Declare Function SHA2_BytesHexHash Lib "diCryptoSys.dll" (ByVal strDigest As String, ByRef lpData As Byte, ByVal nBytes As Long) As Long
Public Declare Function SHA2_BytesHash Lib "diCryptoSys.dll" (ByRef lpDigest As Byte, ByRef lpData As Byte, ByVal nBytes As Long) As Long
Public Declare Function SHA2_Init Lib "diCryptoSys.dll" () As Long
Public Declare Function SHA2_AddString Lib "diCryptoSys.dll" (ByVal hContext As Long, ByVal strData As String) As Long
Public Declare Function SHA2_AddBytes Lib "diCryptoSys.dll" (ByVal hContext As Long, ByRef lpData As Byte, ByVal nBytes As Long) As Long
Public Declare Function SHA2_HexDigest Lib "diCryptoSys.dll" (ByVal strDigest As String, ByVal hContext As Long) As Long
Public Declare Function SHA2_Reset Lib "diCryptoSys.dll" (ByVal hContext As Long) As Long
Public Declare Function SHA2_Hmac Lib "diCryptoSys.dll" (ByVal strDigest As String, ByRef lpData As Byte, ByVal nBytes As Long, ByRef lpKey As Byte, ByVal nKeyBytes As Long) As Long
Public Declare Function SHA2_HmacHex Lib "diCryptoSys.dll" (ByVal strDigest As String, ByVal strHexData As String, ByVal strHexKey As String) As Long

' SECURE HASH ALGORITHM (SHA-3)
' New in [v5.3]
Public Declare Function SHA3_Init Lib "diCryptoSys.dll" (ByVal nHashBitLen As Long) As Long
Public Declare Function SHA3_AddString Lib "diCryptoSys.dll" (ByVal hContext As Long, ByVal strMessage As String) As Long
Public Declare Function SHA3_AddBytes Lib "diCryptoSys.dll" (ByVal hContext As Long, ByRef lpData As Byte, ByVal nDataLen As Long) As Long
Public Declare Function SHA3_HexDigest Lib "diCryptoSys.dll" (ByVal strOutput As String, ByVal nMaxChars As Long, ByVal hContext As Long) As Long
Public Declare Function SHA3_LengthInBytes Lib "diCryptoSys.dll" (ByVal hContext As Long) As Long
Public Declare Function SHA3_Reset Lib "diCryptoSys.dll" (ByVal hContext As Long) As Long

' RSA DATA SECURITY, INC. MD5 MESSAGE-DIGEST ALGORITHM
Public Declare Function MD5_StringHexHash Lib "diCryptoSys.dll" (ByVal strDigest As String, ByVal strData As String) As Long
Public Declare Function MD5_FileHexHash Lib "diCryptoSys.dll" (ByVal strDigest As String, ByVal strFileName As String, ByVal strMode As String) As Long
Public Declare Function MD5_BytesHexHash Lib "diCryptoSys.dll" (ByVal strDigest As String, ByRef lpData As Byte, ByVal nBytes As Long) As Long
Public Declare Function MD5_BytesHash Lib "diCryptoSys.dll" (ByRef lpDigest As Byte, ByRef lpData As Byte, ByVal nBytes As Long) As Long
Public Declare Function MD5_Init Lib "diCryptoSys.dll" () As Long
Public Declare Function MD5_AddString Lib "diCryptoSys.dll" (ByVal hContext As Long, ByVal strData As String) As Long
Public Declare Function MD5_AddBytes Lib "diCryptoSys.dll" (ByVal hContext As Long, ByRef lpData As Byte, ByVal nBytes As Long) As Long
Public Declare Function MD5_HexDigest Lib "diCryptoSys.dll" (ByVal strDigest As String, ByVal hContext As Long) As Long
Public Declare Function MD5_Reset Lib "diCryptoSys.dll" (ByVal hContext As Long) As Long
Public Declare Function MD5_Hmac Lib "diCryptoSys.dll" (ByVal strDigest As String, ByRef lpData As Byte, ByVal nBytes As Long, ByRef lpKey As Byte, ByVal nKeyBytes As Long) As Long
Public Declare Function MD5_HmacHex Lib "diCryptoSys.dll" (ByVal strDigest As String, ByVal strHexData As String, ByVal strHexKey As String) As Long
    
' RC4-COMPATIBLE PC1 FUNCTIONS (Superseded by CIPHER_Stream functions in [v5.0])
Public Declare Function PC1_Bytes Lib "diCryptoSys.dll" (ByRef lpOutput As Byte, ByRef lpData As Byte, ByVal nDataLen As Long, ByRef lpKey As Byte, ByVal nKeyLen As Long) As Long
Public Declare Function PC1_Hex Lib "diCryptoSys.dll" (ByVal strOutput As String, ByVal nMaxChars As Long, ByVal strInputHex As String, ByVal strKeyHex As String) As Long
Public Declare Function PC1_File Lib "diCryptoSys.dll" (ByVal strFileOut As String, ByVal strFileIn As String, ByRef lpKey As Byte, ByVal nKeyLen As Long) As Long

' RANDOM NUMBER GENERATOR (RNG) FUNCTIONS
Public Declare Function RNG_KeyBytes Lib "diCryptoSys.dll" (ByRef lpOutput As Byte, ByVal nBytes As Long, ByVal strSeed As String, ByVal nSeedLen As Long) As Long
Public Declare Function RNG_KeyHex Lib "diCryptoSys.dll" (ByVal strOutput As String, ByVal nMaxChars As Long, ByVal nBytes As Long, ByVal strSeed As String, ByVal nSeedLen As Long) As Long
Public Declare Function RNG_NonceData Lib "diCryptoSys.dll" (ByRef lpOutput As Byte, ByVal nBytes As Long) As Long
Public Declare Function RNG_NonceDataHex Lib "diCryptoSys.dll" (ByVal strOutput As String, ByVal nMaxChars As Long, ByVal nBytes As Long) As Long
Public Declare Function RNG_Test Lib "diCryptoSys.dll" (ByVal strFileName As String) As Long
Public Declare Function RNG_Number Lib "diCryptoSys.dll" (ByVal nLower As Long, ByVal nUpper As Long) As Long
Public Declare Function RNG_BytesWithPrompt Lib "diCryptoSys.dll" (ByRef lpOutput As Byte, ByVal nOutputLen As Long, ByVal strPrompt As String, ByVal nOptions As Long) As Long
Public Declare Function RNG_HexWithPrompt Lib "diCryptoSys.dll" (ByVal strOutput As String, ByVal nMaxChars As Long, ByVal nBytes As Long, ByVal strPrompt As String, ByVal nOptions As Long) As Long
Public Declare Function RNG_Initialize Lib "diCryptoSys.dll" (ByVal strSeedFile As String, ByVal nOptions As Long) As Long
Public Declare Function RNG_MakeSeedFile Lib "diCryptoSys.dll" (ByVal strSeedFile As String, ByVal strPrompt As String, ByVal nOptions As Long) As Long
Public Declare Function RNG_UpdateSeedFile Lib "diCryptoSys.dll" (ByVal strSeedFile As String, ByVal nOptions As Long) As Long
Public Declare Function RNG_TestDRBGVS Lib "diCryptoSys.dll" (ByVal strOutput As String, ByVal nMaxChars As Long, ByVal nReturnedBitsLen As Long, ByVal strEntropyInput As String, ByVal strNonce As String, ByVal strPersonalizationString As String, ByVal strAdditionalInput1 As String, ByVal strEntropyReseed As String, ByVal strAdditionalInputReseed As String, ByVal strAdditionalInput2 As String, ByVal nOptions As Long) As Long
    
' ZLIB COMPRESSION FUNCTIONS
Public Declare Function ZLIB_Deflate Lib "diCryptoSys.dll" (ByRef lpOutput As Byte, ByVal nOutputLen As Long, ByRef lpInput As Byte, ByVal nInputLen As Long) As Long
Public Declare Function ZLIB_Inflate Lib "diCryptoSys.dll" (ByRef lpOutput As Byte, ByVal nOutputLen As Long, ByRef lpInput As Byte, ByVal nInputLen As Long) As Long

' GENERIC COMPRESSION FUNCTIONS
' New in [v6.20]
Public Declare Function COMPR_Compress Lib "diCryptoSys.dll" (ByRef lpOutput As Byte, ByVal nOutBytes As Long, ByRef lpInput As Byte, ByVal nInputLen As Long, ByVal nOptions As Long) As Long
Public Declare Function COMPR_Uncompress Lib "diCryptoSys.dll" (ByRef lpOutput As Byte, ByVal nOutBytes As Long, ByRef lpInput As Byte, ByVal nInputLen As Long, ByVal nOptions As Long) As Long

' PASSWORD-BASED KEY DERIVATION FUNCTIONS
Public Declare Function PBE_Kdf2 Lib "diCryptoSys.dll" (ByRef lpDerivedKey As Byte, ByVal nKeyLen As Long, ByRef lpPwd As Byte, ByVal nPwdlen As Long, ByRef lpSalt As Byte, ByVal nSaltLen As Long, ByVal nCount As Long, ByVal nOptions As Long) As Long
Public Declare Function PBE_Kdf2Hex Lib "diCryptoSys.dll" (ByVal strOutput As String, ByVal nMaxChars As Long, ByVal nKeyBytes As Long, ByVal strPwd As String, ByVal strSaltHex As String, ByVal nCount As Long, ByVal nOptions As Long) As Long
' New in [v5.2]
Public Declare Function PBE_Scrypt Lib "diCryptoSys.dll" (ByRef lpDerivedKey As Byte, ByVal nKeyLen As Long, ByRef lpPwd As Byte, ByVal nPwdlen As Long, ByRef lpSalt As Byte, ByVal nSaltLen As Long, ByVal nParamN As Long, ByVal nParamR As Long, ByVal nParamP As Long, ByVal nOptions As Long) As Long
Public Declare Function PBE_ScryptHex Lib "diCryptoSys.dll" (ByVal strOutput As String, ByVal nMaxChars As Long, ByVal dkBytes As Long, ByVal strPwd As String, ByVal strSaltHex As String, ByVal nParamN As Long, ByVal nParamR As Long, ByVal nParamP As Long, ByVal nOptions As Long) As Long

' HEX ENCODING CONVERSION FUNCTIONS
' See cnvHexStrFromBytes, cnvBytesFromHexStr, cnvHexFilter below
Public Declare Function CNV_HexStrFromBytes Lib "diCryptoSys.dll" (ByVal strOutput As String, ByVal nMaxChars As Long, ByRef lpData As Byte, ByVal nBytes As Long) As Long
Public Declare Function CNV_BytesFromHexStr Lib "diCryptoSys.dll" (ByRef lpOutput As Byte, ByVal nOutLen As Long, ByVal strInput As String) As Long
Public Declare Function CNV_HexFilter Lib "diCryptoSys.dll" (ByVal strOutput As String, ByVal strInput As String, ByVal nStrLen As Long) As Long

' BASE64 ENCODING CONVERSION FUNCTIONS
' See cnvB64StrFromBytes, cnvBytesFromHexB64, cnvB64Filter below
Public Declare Function CNV_B64StrFromBytes Lib "diCryptoSys.dll" (ByVal strOutput As String, ByVal nMaxChars As Long, ByRef lpData As Byte, ByVal nBytes As Long) As Long
Public Declare Function CNV_BytesFromB64Str Lib "diCryptoSys.dll" (ByRef lpOutput As Byte, ByVal nOutLen As Long, ByVal strInput As String) As Long
Public Declare Function CNV_B64Filter Lib "diCryptoSys.dll" (ByVal strOutput As String, ByVal strInput As String, ByVal nStrLen As Long) As Long

' CRC FUNCTIONS
Public Declare Function CRC_Bytes Lib "diCryptoSys.dll" (ByRef lpInput As Byte, ByVal nBytes As Long, ByVal nOptions As Long) As Long
Public Declare Function CRC_String Lib "diCryptoSys.dll" (ByVal strInput As String, ByVal nOptions As Long) As Long
Public Declare Function CRC_File Lib "diCryptoSys.dll" (ByVal strFileName As String, ByVal nOptions As Long) As Long

' FUNCTIONS TO WIPE DATA
Public Declare Function WIPE_File Lib "diCryptoSys.dll" (ByVal strFileName As String, ByVal nOptions As Long) As Long
Public Declare Function WIPE_Data Lib "diCryptoSys.dll" (ByRef lpData As Byte, ByVal nBytes As Long) As Long
' Alternative Aliases to cope with Byte and String types explicitly...
Public Declare Function WIPE_Bytes Lib "diCryptoSys.dll" Alias "WIPE_Data" (ByRef lpData As Byte, ByVal nBytes As Long) As Long
Public Declare Function WIPE_String Lib "diCryptoSys.dll" Alias "WIPE_Data" (ByVal strData As String, ByVal nStrLen As Long) As Long

' PADDING FUNCTIONS
Public Declare Function PAD_BytesBlock Lib "diCryptoSys.dll" (ByRef lpOutput As Byte, ByVal nOutputLen As Long, ByRef lpInput As Byte, ByVal nInputLen As Long, ByVal nBlockLen As Long, ByVal nOptions As Long) As Long
Public Declare Function PAD_UnpadBytes Lib "diCryptoSys.dll" (ByRef lpOutput As Byte, ByVal nOutputLen As Long, ByRef lpInput As Byte, ByVal nInputLen As Long, ByVal nBlockLen As Long, ByVal nOptions As Long) As Long
Public Declare Function PAD_HexBlock Lib "diCryptoSys.dll" (ByVal strOutput As String, ByVal nMaxChars As Long, ByVal strInputHex As String, ByVal nBlockLen As Long, ByVal nOptions As Long) As Long
Public Declare Function PAD_UnpadHex Lib "diCryptoSys.dll" (ByVal strOutput As String, ByVal nMaxChars As Long, ByVal strInputHex As String, ByVal nBlockLen As Long, ByVal nOptions As Long) As Long

' XOF/PRF PROTOTYPES
Public Declare Function XOF_Bytes Lib "diCryptoSys.dll" (ByRef lpOutput As Byte, ByVal nOutBytes As Long, ByRef lpMessage As Byte, ByVal nMsgLen As Long, ByVal nOptions As Long) As Long
Public Declare Function PRF_Bytes Lib "diCryptoSys.dll" (ByRef lpOutput As Byte, ByVal nOutBytes As Long, ByRef lpMessage As Byte, ByVal nMsgLen As Long, ByRef lpKey As Byte, ByVal nKeyLen As Long, ByVal strCustom As String, ByVal nOptions As Long) As Long

#End If

' *** END OF CRYPTOSYS API DECLARATIONS

' *****************
' WRAPPER FUNCTIONS
' *****************
' Direct calls to the DLL begin with "XXX_", wrapper functions begin with "xxx"


'--------------------
' UTILITY FUNCTION
'--------------------
'/**
' Find length of byte array.
' @param  ab Input byte array.
' @return Number of bytes in array.
' @remark Safe to use even if array is empty.
' @example
' {@code
' Dim ab() As Byte
' Debug.Print cnvBytesLen(ab) ' Expecting 0
' ReDim ab(10)    ' NB actually 11 elements (0..10)
' Debug.Print cnvBytesLen(ab) ' 11
' ab = vbNullString   ' Set to empty array
' Debug.Print cnvBytesLen(ab) ' 0
' }
'**/
Public Function cnvBytesLen(ab() As Byte) As Long
    ' Trap error if array is empty
    On Error Resume Next
    cnvBytesLen = UBound(ab) - LBound(ab) + 1
End Function


'/**
' Encodes an array of bytes as a hexadecimal-encoded string.
' @param abData Input byte array.
' @return Hexadecimal-encoded string.
' @remark Same as {@link cnvToHex}.
'**/
Public Function cnvHexStrFromBytes(abData() As Byte) As String
    Dim strHex As String
    Dim nHexLen As Long
    Dim nDataLen As Long
    
    nDataLen = cnvBytesLen(abData)
    If nDataLen = 0 Then Exit Function
    nHexLen = CNV_HexStrFromBytes(vbNullString, 0, abData(0), nDataLen)
    If nHexLen <= 0 Then
        Exit Function
    End If
    strHex = String$(nHexLen, " ")
    nHexLen = CNV_HexStrFromBytes(strHex, nHexLen, abData(0), nDataLen)
    If nHexLen <= 0 Then
        Exit Function
    End If
    cnvHexStrFromBytes = Left$(strHex, nHexLen)
End Function

'/**
' Encodes an ANSI string as a hexadecimal-encoded string.
' @param strData String to be encoded.
' @return Hexadecimal-encoded string.
' @remark Expecting a string of 8-bit "ANSI" characters.
'**/
Public Function cnvHexStrFromString(strData As String) As String
    Dim strHex As String
    Dim nHexLen As Long
    Dim nDataLen As Long
    Dim abData() As Byte
    
    If Len(strData) = 0 Then Exit Function
    abData = StrConv(strData, vbFromUnicode)
    nDataLen = cnvBytesLen(abData)
    nHexLen = CNV_HexStrFromBytes(vbNullString, 0, abData(0), nDataLen)
    If nHexLen <= 0 Then
        Exit Function
    End If
    strHex = String$(nHexLen, " ")
    nHexLen = CNV_HexStrFromBytes(strHex, nHexLen, abData(0), nDataLen)
    If nHexLen <= 0 Then
        Exit Function
    End If
    cnvHexStrFromString = Left$(strHex, nHexLen)
End Function

'/**
' Decodes a hexadecimal-encoded string as an array of Bytes.
' @param strHex Hexadecimal data to be decoded.
' @return Array of bytes.
' @remark Same as {@link cnvFromHex}.
'**/
Public Function cnvBytesFromHexStr(strHex As String) As Byte()
    Dim abData() As Byte
    Dim nDataLen As Long
    
    ' Set default return value that won't cause a run-time error
    cnvBytesFromHexStr = vbNullString
    nDataLen = CNV_BytesFromHexStr(0, 0, strHex)
    If nDataLen <= 0 Then
        Exit Function
    End If
    ReDim abData(nDataLen - 1)
    nDataLen = CNV_BytesFromHexStr(abData(0), nDataLen, strHex)
    If nDataLen <= 0 Then
        Exit Function
    End If
    ReDim Preserve abData(nDataLen - 1)
    cnvBytesFromHexStr = abData
End Function

'/**
' Decodes a hexadecimal-encoded string as an ANSI string.
' @param strHex Hexadecimal data to be decoded.
' @return Decoded string. For example, "6162632E" will be converted to "abc."
' @remark Output is a string of "ANSI" characters of value between 0 and 255.
'**/
Public Function cnvStringFromHexStr(ByVal strHex As String) As String
    Dim abData() As Byte
    If Len(strHex) = 0 Then Exit Function
    abData = cnvBytesFromHexStr(strHex)
    cnvStringFromHexStr = StrConv(abData, vbUnicode)
End Function

'/**
' Strips any invalid hex characters from a hex string.
' @param strHex String to be filtered.
' @return Filtered string.
'**/
Public Function cnvHexFilter(strHex As String) As String
    Dim strFiltered As String
    Dim nLen As Long
    
    strFiltered = String(Len(strHex), " ")
    nLen = CNV_HexFilter(strFiltered, strHex, Len(strHex))
    If nLen > 0 Then
        strFiltered = Left$(strFiltered, nLen)
    Else
        strFiltered = ""
    End If
    cnvHexFilter = strFiltered
End Function

'/**
' Encodes an array of bytes as a base64-encoded string.
' @param abData Input byte array.
' @return Base64-encoded string.
' @remark Same as {@link cnvToBase64}.
'**/
Public Function cnvB64StrFromBytes(abData() As Byte) As String
    Dim strB64 As String
    Dim nB64Len As Long
    Dim nDataLen As Long
    
    nDataLen = cnvBytesLen(abData)
    nB64Len = CNV_B64StrFromBytes(vbNullString, 0, abData(0), nDataLen)
    If nB64Len <= 0 Then Exit Function
    strB64 = String$(nB64Len, " ")
    nB64Len = CNV_B64StrFromBytes(strB64, nB64Len, abData(0), nDataLen)
    If nB64Len <= 0 Then Exit Function
    cnvB64StrFromBytes = Left$(strB64, nB64Len)
End Function

'/**
' Encodes an ANSI string as a base64-encoded string.
' @param strData String to be encoded.
' @return Base64-encoded string.
' @remark Expecting a string of 8-bit "ANSI" characters.
'**/
Public Function cnvB64StrFromString(strData As String) As String
    Dim strB64 As String
    Dim nB64Len As Long
    Dim nDataLen As Long
    Dim abData() As Byte
    
    If Len(strData) = 0 Then Exit Function
    abData = StrConv(strData, vbFromUnicode)
    nDataLen = UBound(abData) - LBound(abData) + 1
    nB64Len = CNV_B64StrFromBytes(vbNullString, 0, abData(0), nDataLen)
    If nB64Len <= 0 Then Exit Function
    strB64 = String$(nB64Len, " ")
    nB64Len = CNV_B64StrFromBytes(strB64, nB64Len, abData(0), nDataLen)
    If nB64Len <= 0 Then Exit Function
    cnvB64StrFromString = Left$(strB64, nB64Len)
End Function

'/**
' Decodes a base64-encoded string as an array of Bytes.
' @param strB64 Base64 data to be decoded.
' @return Array of bytes.
' @remark Same as {@link cnvFromBase64}.
'**/
Public Function cnvBytesFromB64Str(strB64 As String) As Byte()
    Dim abData() As Byte
    Dim nDataLen As Long
    
    ' Set default return value that won't cause a run-time error
    cnvBytesFromB64Str = vbNullString
    nDataLen = CNV_BytesFromB64Str(0, 0, strB64)
    If nDataLen <= 0 Then Exit Function
    ReDim abData(nDataLen - 1)
    nDataLen = CNV_BytesFromB64Str(abData(0), nDataLen, strB64)
    If nDataLen <= 0 Then Exit Function
    ReDim Preserve abData(nDataLen - 1)
    cnvBytesFromB64Str = abData
End Function

'/**
' Strips any invalid base64 characters from a string.
' @param strB64 String to be filtered.
' @return Filtered string.
'**/
Public Function cnvB64Filter(strB64 As String) As String
    Dim strFiltered As String
    Dim nLen As Long
    
    strFiltered = String(Len(strB64), " ")
    nLen = CNV_B64Filter(strFiltered, strB64, Len(strB64))
    If nLen > 0 Then
        strFiltered = Left$(strFiltered, nLen)
    Else
        strFiltered = ""
    End If
    cnvB64Filter = strFiltered
End Function

'/**
' Re-encodes a hexadecimal-encoded binary value as base64.
' @param strHex Hex string representing a binary value.
' @return Binary value encoded in base64
'**/
Public Function cnvB64StrFromHexStr(strHex As String) As String
    cnvB64StrFromHexStr = cnvB64StrFromBytes(cnvBytesFromHexStr(strHex))
End Function

'/**
' Re-encodes a base64-encoded binary value as hexadecimal.
' @param strB64 Base64 string representing a binary value.
' @return Binary value encoded in hexadecimal
'**/
Public Function cnvHexStrFromB64Str(strB64 As String) As String
    cnvHexStrFromB64Str = cnvHexStrFromBytes(cnvBytesFromB64Str(strB64))
End Function

'/**
' Encodes a substring of an array of bytes as a hexadecimal-encoded string.
' @param abData Input byte array.
' @param nOffset Offset at which substring begins. First byte is at offset zero.
' @param nBytes Number of bytes in substring to encode.
' @return Hexadecimal-encoded string.
' @example
' {@code
' Debug.Print cnvHexFromBytesMid(cnvBytesFromHexStr("00112233445566"), 3, 2) ' 3344
' }
'**/
Public Function cnvHexFromBytesMid(abData() As Byte, nOffset As Long, nBytes As Long) As String
    Dim strHex As String
    ' Lazy but safe! Encode it all then grab the substring
    strHex = cnvHexStrFromBytes(abData)
    cnvHexFromBytesMid = Mid(strHex, nOffset * 2 + 1, nBytes * 2)
End Function

' New in [v6.20] more convenient synonyms

'/**
' Encodes an array of bytes as a hexadecimal-encoded string.
' @param lpData Input byte array
' @return Hexadecimal-encoded string
' @remark A shorter synonym for {@link cnvHexStrFromBytes}
'**/
Public Function cnvToHex(lpData() As Byte) As String
    cnvToHex = cnvHexStrFromBytes(lpData)
End Function

'/**
' Decodes a hexadecimal-encoded string as an array of bytes.
' @param strHex Hexadecimal-encoded data to be decoded.
' @return Array of bytes.
' @remark A shorter synonym for {@link cnvBytesFromHexStr}
'**/
Public Function cnvFromHex(strHex As String) As Byte()
    cnvFromHex = cnvBytesFromHexStr(strHex)
End Function

'/**
' Encodes an array of bytes as a base64-encoded string.
' @param lpData Input byte array
' @return Base64-encoded string
' @remark A shorter synonym for {@link cnvB64StrFromBytes}
'**/
Public Function cnvToBase64(lpData() As Byte) As String
    cnvToBase64 = cnvB64StrFromBytes(lpData)
End Function

'/**
' Decodes a base64-encoded string as an array of bytes.
' @param strBase64 Base64 data to be decoded.
' @return Array of bytes.
' @remark A shorter synonym for {@link cnvBytesFromB64Str}
'**/
Public Function cnvFromBase64(strBase64 As String) As Byte()
    cnvFromBase64 = cnvBytesFromB64Str(strBase64)
End Function

'/**
' Return a substring of bytes of specified length from within a given byte array
' @param Bytes Byte array from which to return a substring (of bytes)
' @param nOffset Offset at which substring begins. First byte is at offset zero.
' @param nBytes Number of bytes in substring (optional). If negative, copy to end of input.
'**/
Public Function cnvBytesMid(Bytes() As Byte, nOffset As Long, Optional nBytes As Long = -1) As Byte()
    cnvBytesMid = vbNullString
    Dim MyBytes() As Byte
    Dim nLen As Long
    Dim i As Long
    Dim nRest As Long
    Dim nToCopy As Long
    nLen = cnvBytesLen(Bytes)
    ' Cases with empty string output
    If nLen = 0 Then Exit Function
    If nBytes = 0 Then Exit Function
    If nOffset >= nLen Then Exit Function
    ' Max bytes to copy to end of string
    nRest = nLen - nOffset
    If nBytes < 0 Then
        nToCopy = nRest
    Else
        nToCopy = nBytes
    End If
    If nToCopy > nRest Then nToCopy = nRest
    ReDim MyBytes(nToCopy - 1)
    For i = 0 To nToCopy - 1
        MyBytes(i) = Bytes(i + nOffset)
    Next
    cnvBytesMid = MyBytes
End Function

'/**
' Retrieves the error message associated with a given error code.
' @param nCode Error code for which the message is required.
' @return Error message, or empty string if no corresponding error code.
'**/
Public Function apiErrorLookup(nCode As Long) As String
    Dim strMsg As String
    Dim nRet As Long
    
    strMsg = String(128, " ")
    nRet = API_ErrorLookup(strMsg, Len(strMsg), nCode)
    apiErrorLookup = Left(strMsg, nRet)
End Function

'/**
' Returns the error code of the error that occurred when calling the last function.
' @return Error code (see {@link apiErrorLookup}).
' @remark Not all functions set this value.
'**/
Public Function apiErrorCode() As Long
    apiErrorCode = API_ErrorCode()
End Function


'/**
' Return an error message string for the last error.
' @param nErrCode Error code returned by last call.
' @param szMsg Optional message to add.
' @return Error message as a string including previous ErrorCode, if available.
' @example
' {@code
' Error (11): Parameter out of range (RANGE_ERROR)
' }
'**/
Public Function errFormatErrorMessage(Optional nErrCode As Long = 0, Optional szMsg As String = "") As String
    Dim nLastCode As Long
    errFormatErrorMessage = vbNullString
    If nErrCode < 0 Then nErrCode = -nErrCode
    ' Get previous error code, if available
    nLastCode = apiErrorCode()
    If nErrCode = 0 And nLastCode = 0 Then
        Exit Function
    End If
    ' Compose error message
    If Len(szMsg) > 0 Then
        errFormatErrorMessage = errFormatErrorMessage & szMsg & ": "
    End If
    errFormatErrorMessage = errFormatErrorMessage & "Error"
    If nErrCode <> 0 Then
        errFormatErrorMessage = errFormatErrorMessage & "(" & nErrCode & ")"
    End If
    ' Get error message for code errCode
    If (nErrCode <> 0) Then
        errFormatErrorMessage = errFormatErrorMessage & ": " & apiErrorLookup(nErrCode)
    End If
    If nLastCode <> 0 And nErrCode <> nLastCode Then
        errFormatErrorMessage = errFormatErrorMessage & ": " & apiErrorLookup(nLastCode)
    End If
    
End Function


'/**
' Generate a hex-encoded sequence of bytes.
' @param nBytes Required number of random bytes.
' @return Hex-encoded random bytes.
'**/
Public Function rngNonceHex(nBytes As Long) As String
    Dim strHex As String
    Dim lngRet As Long
    
    strHex = String(nBytes * 2, " ")
    lngRet = RNG_NonceDataHex(strHex, Len(strHex), nBytes)
    If lngRet = 0 Then
        rngNonceHex = strHex
    End If
End Function

'/**
' Add PKCS5 padding to a hex string up to next multiple of block length [DEPRECATED].
' @param strInputHex Hexadecimal-encoded data to be padded.
' @param nBlockLen Cipher block length in bytes (8 or 16).
' @return Padded hex string or empty string on error.
' @deprecated Use `padHexBlock()`.
'**/
Public Function padHexString(ByVal strInputHex As String, nBlockLen As Long) As String
    Dim nOutChars As Long
    Dim strOutputHex As String
    
    ' In VB6 an uninitialised empty string is passed to a DLL as a NULL,
    ' so we append a non-null empty string!
    strInputHex = strInputHex & ""
    
    nOutChars = PAD_HexBlock("", 0, strInputHex, nBlockLen, 0)
    'Debug.Print "Required length is " & nOutChars & " characters"
    ' Check for error
    If (nOutChars <= 0) Then Exit Function
    
    ' Pre-dimension output
    strOutputHex = String(nOutChars, " ")
    
    nOutChars = PAD_HexBlock(strOutputHex, Len(strOutputHex), strInputHex, nBlockLen, 0)
    If (nOutChars <= 0) Then Exit Function
    'Debug.Print "Padded data='" & strOutputHex & "'"
    
    padHexString = strOutputHex
    
End Function

'/**
' Strips PKCS5 padding from a hex string [DEPRECATED].
' @param strInputHex Hexadecimal-encoded padded data.
' @param nBlockLen Cipher block length in bytes (8 or 16).
' @return  Unpadded data in hex string or _unchanged_ data on error.
' @remark An error is indicated by returning the _original_ data
' which will always be longer than the expected unpadded result.
' @deprecated Use {@link padUnpadHex}.
'**/
Public Function unpadHexString(strInputHex As String, nBlockLen As Long) As String
' Strips PKCS5 padding from a hex string.
' Returns unpadded hex string or, on error, the original input string
' -- we do this because an empty string is a valid result.
' To check for error: a valid output string is *always* shorter than the input.

    Dim nOutChars As Long
    Dim strOutputHex As String
    
    ' No need to query for length because we know the output will be shorter than input
    ' so make sure output is as long as the input
    strOutputHex = String(Len(strInputHex), " ")
    nOutChars = PAD_UnpadHex(strOutputHex, Len(strOutputHex), strInputHex, nBlockLen, 0)
    'Debug.Print "Unpadded length is " & nOutChars & " characters"
    
    ' Check for error
    If (nOutChars < 0) Then
        ' Return unchanged input to indicate error
        unpadHexString = strInputHex
        Exit Function
    End If
    
    ' Re-dimension the output to the correct length
    strOutputHex = Left$(strOutputHex, nOutChars)
    'Debug.Print "Unpadded data='" & strOutputHex & "'"
    
    unpadHexString = strOutputHex
    
End Function


'/**
' Decrypt data using specified AEAD algorithm in one-off operation. The authentication tag is expected to be appended to the input ciphertext.
' @param  lpData Input data to be decrypted.
' @param  lpKey Key of exact length for algorithm (16 or 32 bytes).
' @param  lpNonce Initialization Vector (IV) (aka nonce) exactly 12 bytes long.
' @param  lpAAD Additional authenticated data (optional) - set variable as `vbNullString` to ignore.
' @param  nOptions Algorithm to be used. Select one from
' {@code
' API_AEAD_AES_128_GCM
' API_AEAD_AES_256_GCM
' API_AEAD_CHACHA20_POLY1305
' }
' Add `API_IV_PREFIX` to expect the IV to be prepended at the start of the input (use the `Or` operator).
' @return Plaintext in a byte array, or empty array on error (an empty array may also be the correct result).
' @remark The input must include the 16-byte tag appended to the ciphertext.
'**/
Public Function aeadDecryptWithTag(lpData() As Byte, lpKey() As Byte, lpNonce() As Byte, lpAAD() As Byte, nOptions As Long) As Byte()
    aeadDecryptWithTag = vbNullString
    Dim n1 As Long
    n1 = cnvBytesLen(lpData)
    If n1 = 0 Then Exit Function
    Dim n2 As Long
    n2 = cnvBytesLen(lpKey)
    If n2 = 0 Then Exit Function
    Dim n3 As Long
    n3 = cnvBytesLen(lpNonce)
    If n3 = 0 Then Exit Function
    Dim n4 As Long
    n4 = cnvBytesLen(lpAAD)
    ' Fudge to allow an empty input array
    If n4 = 0 Then ReDim lpAAD(0)
    Dim abMyData() As Byte
    Dim nb As Long
    nb = AEAD_DecryptWithTag(ByVal 0&, 0, lpData(0), n1, lpKey(0), n2, lpNonce(0), n3, lpAAD(0), n4, nOptions)
    If nb <= 0 Then GoTo CleanUp
    ReDim abMyData(nb - 1)
    nb = AEAD_DecryptWithTag(abMyData(0), nb, lpData(0), n1, lpKey(0), n2, lpNonce(0), n3, lpAAD(0), n4, nOptions)
    If nb <= 0 Then GoTo CleanUp
    ReDim Preserve abMyData(nb - 1)
    aeadDecryptWithTag = abMyData
CleanUp:
    If n4 = 0 Then lpAAD = vbNullString
End Function

'/**
' Encrypt data using specified AEAD algorithm in one-off operation. The authentication tag is appended to the output.
' @param  lpData Input data to be encrypted.
' @param  lpKey Key of exact length for algorithm (16 or 32 bytes).
' @param  lpNonce Initialization Vector (IV) (aka nonce) exactly 12 bytes long.
' @param  lpAAD Additional authenticated data (optional) - set variable as `vbNullString` to ignore.
' @param  nOptions Algorithm to be used. Select one from
' {@code
' API_AEAD_AES_128_GCM
' API_AEAD_AES_256_GCM
' API_AEAD_CHACHA20_POLY1305
' }
' Add `API_IV_PREFIX` to prepend the IV (nonce) before the ciphertext in the output (use the `Or` operator).
' @return Ciphertext with tag appended in a byte array, or empty array on error.
' @remark The output will either be exactly 16 bytes longer than the input, or exactly 28 bytes longer if `API_IV_PREFIX` is used.
'**/
Public Function aeadEncryptWithTag(lpData() As Byte, lpKey() As Byte, lpNonce() As Byte, lpAAD() As Byte, nOptions As Long) As Byte()
    aeadEncryptWithTag = vbNullString
    Dim n1 As Long
    n1 = cnvBytesLen(lpData)
    If n1 = 0 Then Exit Function
    Dim n2 As Long
    n2 = cnvBytesLen(lpKey)
    If n2 = 0 Then Exit Function
    Dim n3 As Long
    n3 = cnvBytesLen(lpNonce)
    If n3 = 0 Then Exit Function
    Dim n4 As Long
    n4 = cnvBytesLen(lpAAD)
    ' Fudge to allow an empty input array
    If n4 = 0 Then ReDim lpAAD(0)
    Dim abMyData() As Byte
    Dim nb As Long
    nb = AEAD_EncryptWithTag(ByVal 0&, 0, lpData(0), n1, lpKey(0), n2, lpNonce(0), n3, lpAAD(0), n4, nOptions)
    If nb <= 0 Then GoTo CleanUp
    ReDim abMyData(nb - 1)
    nb = AEAD_EncryptWithTag(abMyData(0), nb, lpData(0), n1, lpKey(0), n2, lpNonce(0), n3, lpAAD(0), n4, nOptions)
    If nb <= 0 Then GoTo CleanUp
    ReDim Preserve abMyData(nb - 1)
    aeadEncryptWithTag = abMyData
CleanUp:
    If n4 = 0 Then lpAAD = vbNullString
End Function

'/**
' Initialize the AEAD context with the key and algorithm ready for repeated incremental operations.
' @param  lpKey Key of exact length for algorithm (16 or 32 bytes).
' @param  nOptions Algorithm to be used. Select one from
' {@code
' API_AEAD_AES_128_GCM
' API_AEAD_AES_256_GCM
' API_AEAD_CHACHA20_POLY1305
' }
' @return Nonzero handle of the AEAD context, or _zero_ if an error occurs.
'**/
Public Function aeadInitKey(lpKey() As Byte, nOptions As Long) As Long
    aeadInitKey = AEAD_InitKey(lpKey(0), cnvBytesLen(lpKey), nOptions)
End Function

'/**
' Set the nonce for the AEAD context (in incremental mode).
' @param hContext Handle to the AEAD context set up by an earlier call to {@link aeadInitKey}.
' @param  lpNonce Nonce of exact length required for given algorithm (currently always 12 bytes).
' @return  Zero (0) on success, or a nonzero error code (use {@link apiErrorLookup} to check).
'**/
Public Function aeadSetNonce(hContext As Long, lpNonce() As Byte) As Long
    aeadSetNonce = AEAD_SetNonce(hContext, lpNonce(0), cnvBytesLen(lpNonce))
End Function

'/**
' Add a chunk of additional authenticated data (in incremental mode).
' @param hContext Handle to the AEAD context set up by an earlier call to {@link aeadInitKey}.
' @param  lpAAD Chunk of Additional Authenticated Data (AAD) to add.
' @return  Zero (0) on success, or a nonzero error code (use {@link apiErrorLookup} to check).
' @remark May be repeated to add additional data in chunks.
' Must eventually be followed by either {@link aeadStartEncrypt} or {@link aeadStartDecrypt}.
' Returns `MISUSE_ERROR` if called out of sequence.
'**/
Public Function aeadAddAAD(hContext As Long, lpAAD() As Byte) As Long
    aeadAddAAD = AEAD_AddAAD(hContext, lpAAD(0), cnvBytesLen(lpAAD))
End Function

'/**
' Start authenticated encryption (in incremental mode).
' @param hContext Handle to the AEAD context set up by an earlier call to {@link aeadInitKey}.
' @return  Zero (0) on success, or a nonzero error code (use {@link apiErrorLookup} to check).
' @remark May be followed by zero or more calls to {@link aeadUpdate} to encrypt data in chunks.
' Must eventually be followed by {@link aeadFinishEncrypt}.
' Returns `MISUSE_ERROR` if called out of sequence.
'**/
Public Function aeadStartEncrypt(hContext As Long) As Long
    aeadStartEncrypt = AEAD_StartEncrypt(hContext)
End Function

'/**
' Start authenticated decryption (in incremental mode).
' @param hContext Handle to the AEAD context set up by an earlier call to {@link aeadInitKey}.
' @param lpTagToCheck Byte array containing the tag to be checked.
' @return  Zero (0) on success, or a nonzero error code (use {@link apiErrorLookup} to check).
' @remark May be followed by zero or more calls to {@link aeadUpdate} to decrypt data in chunks.
' Must eventually be followed by {@link aeadFinishDecrypt}.
' Returns `MISUSE_ERROR` if called out of sequence.
' __Caution__: do not trust decrypted data until final authentication.
'**/
Public Function aeadStartDecrypt(hContext As Long, lpTagToCheck() As Byte) As Long
    aeadStartDecrypt = AEAD_StartDecrypt(hContext, lpTagToCheck(0), cnvBytesLen(lpTagToCheck))
End Function

'/**
' Encrypts or decrypts a chunk of input (in incremental mode).
' @param hContext Handle to the AEAD context set up by an earlier call to {@link aeadInitKey}.
' @param lpData Data to be encrypted or decrypted.
' @return Encrypted or decrypted data in array of exactly the same length as input; or an empty array on error
' @remark This function may be repeated to add data in chunks.
' The input data is encrypted or decrypted depending on the start mode set by a preceding call to
' {@link aeadStartEncrypt} or {@link aeadStartDecrypt}, respectively.
' It must eventually be followed by either {@link aeadFinishEncrypt} or {@link aeadFinishDecrypt}, which must match the start mode.
'**/
Public Function aeadUpdate(hContext As Long, lpData() As Byte) As Byte()
    aeadUpdate = vbNullString
    Dim nb As Long
    nb = cnvBytesLen(lpData)
    If nb = 0 Then Exit Function
    Dim abMyData() As Byte
    ReDim abMyData(nb - 1)
    nb = AEAD_Update(hContext, abMyData(0), nb, lpData(0), nb)
    If nb <> 0 Then Exit Function
    aeadUpdate = abMyData
End Function

'/**
' Finish the authenticated encryption (in incremental mode).
' @param hContext Handle to the AEAD context set up by an earlier call to {@link aeadInitKey}.
' @return  Authentication tag value.
' @remark Must be preceded by {@link aeadStartEncrypt} and zero or more calls to {@link aeadUpdate}.
' May be followed by {@link aeadSetNonce} to begin processing another packet with the same key and algorithm;
' otherwise should be followed by {@link aeadDestroy}.
'**/
Public Function aeadFinishEncrypt(hContext As Long) As Byte()
    aeadFinishEncrypt = vbNullString
    Dim nb As Long
    nb = API_AEAD_TAG_MAX_BYTES
    Dim abMyData() As Byte
    ReDim abMyData(nb - 1)
    nb = AEAD_FinishEncrypt(hContext, abMyData(0), nb)
    If nb <> 0 Then Exit Function
    aeadFinishEncrypt = abMyData
End Function

'/**
' Finish the authenticated decryption (in incremental mode).
' @param hContext Handle to the AEAD context set up by an earlier call to {@link aeadInitKey}.
' @return  Zero (0) on success, or a nonzero error code (use {@link apiErrorLookup} to check).
' @remark Returns the nonzero error code `AUTH_FAILED_ERROR` (-40) if the inputs are not authentic.
' Must be preceded by {@link aeadStartDecrypt} and zero or more calls to {@link aeadUpdate}.
' May be followed by {@link aeadSetNonce} to begin processing another packet with the same key and algorithm;
' otherwise should be followed by {@link aeadDestroy}.
' Returns `MISUSE_ERROR` if called out of sequence.
'**/
Public Function aeadFinishDecrypt(hContext As Long) As Long
    aeadFinishDecrypt = AEAD_FinishDecrypt(hContext)
End Function

'/**
' Close the AEAD context and destroy the key (in incremental mode).
' @param hContext Handle to the AEAD context set up by an earlier call to {@link aeadInitKey}.
' @return  Zero (0) on success, or a nonzero error code (use {@link apiErrorLookup} to check).
'**/
Public Function aeadDestroy(hContext As Long) As Long
    aeadDestroy = AEAD_Destroy(hContext)
End Function


'/**
' Gets date and time the CryptoSys DLL module was last compiled.
' @return Date and time string.
'**/
Public Function apiCompileTime() As String
    Dim nc As Long
    nc = API_CompileTime(vbNullString, 0)
    If nc <= 0 Then Exit Function
    apiCompileTime = String(nc, " ")
    nc = API_CompileTime(apiCompileTime, nc)
    apiCompileTime = Left$(apiCompileTime, nc)
End Function

'/**
' Retrieves the name of the current process's module.
' @param  nOptions For future use.
' @return File path to current DLL module.
'**/
Public Function apiModuleName(Optional nOptions As Long = 0) As String
    Dim nc As Long
    nc = API_ModuleName(vbNullString, 0, nOptions)
    If nc <= 0 Then Exit Function
    apiModuleName = String(nc, " ")
    nc = API_ModuleName(apiModuleName, nc, nOptions)
    apiModuleName = Left$(apiModuleName, nc)
End Function

'/**
' Returns the ASCII value of the licence type.
' @param  nOptions For future use.
' @return `D`=Developer `T`=Trial.
'**/
Public Function apiLicenceType(Optional nOptions As Long = 0) As String
    Dim n As Long
    n = API_LicenceType(nOptions)
    apiLicenceType = Chr(n)
End Function

'/**
' Get version number of native core DLL.
' @return Version number as an integer in form `Major*100*100 + Minor*100 + Revision`. For example, version 6.1.2 would return `60102`.
'**/
Public Function apiVersion() As Long
    apiVersion = API_Version()
End Function

'/**
' Decrypts data in a byte array using the specified block cipher algorithm, mode and padding.
' The key and initialization vector are passed as byte arrays.
' @param  lpInput  Input data to be decrypted.
' @param  lpKey    Key of exact length for block cipher algorithm.
' @param  lpIV     Initialization Vector (IV) of exactly the block size (if not provided in input) or empty array for ECB mode.
' @param  szAlgModePad     String with block cipher algorithm, mode and padding,
' e.g. <code>"aes128/cbc/pkcs5"</code>
' {@code
' Alg:  aes128|aes192|aes256|tdea|3des|desede3
' Mode: ecb|cbc|ofb|cfb|ctr
' Pad:  pkcs5|nopad|oneandzeroes|ansix923|w3c
' }
' @param  nOptions  Add `API_IV_PREFIX` to expect the IV to be prepended at the start of the input
' (ignored for ECB mode).
' @return Decrypted plaintext in byte array or empty array on error.
' @remark Default padding is `Pkcs5` for ECB and CBC mode and `NoPad` for all other modes.
'**/
Public Function cipherDecryptBytes(lpInput() As Byte, lpKey() As Byte, lpIV() As Byte, szAlgModePad As String, Optional nOptions As Long = 0) As Byte()
    cipherDecryptBytes = vbNullString
    Dim n1 As Long
    n1 = cnvBytesLen(lpInput)
    If n1 = 0 Then Exit Function
    Dim n2 As Long
    n2 = cnvBytesLen(lpKey)
    If n2 = 0 Then Exit Function
    Dim n3 As Long
    n3 = cnvBytesLen(lpIV)
    ' Fudge to allow an empty input array
    If n3 = 0 Then ReDim lpIV(0)
    Dim abMyData() As Byte
    Dim nb As Long
    nb = CIPHER_DecryptBytes(ByVal 0&, 0, lpInput(0), n1, lpKey(0), n2, lpIV(0), n3, szAlgModePad, nOptions)
    If nb <= 0 Then GoTo CleanUp
    ReDim abMyData(nb - 1)
    nb = CIPHER_DecryptBytes(abMyData(0), nb, lpInput(0), n1, lpKey(0), n2, lpIV(0), n3, szAlgModePad, nOptions)
    If nb <= 0 Then GoTo CleanUp
    ReDim Preserve abMyData(nb - 1)
    cipherDecryptBytes = abMyData
CleanUp:
    If n3 = 0 Then lpIV = vbNullString
End Function

' @deprecated Use cipherDecryptBytes()
Public Function cipherDecryptBytes2(lpInput() As Byte, lpKey() As Byte, lpIV() As Byte, szAlgModePad As String, Optional nOptions As Long = 0) As Byte()
    cipherDecryptBytes2 = cipherDecryptBytes(lpInput, lpKey, lpIV, szAlgModePad, nOptions)
End Function

'/**
' Decrypt hex-encoded data using the specified block cipher algorithm, mode and padding.
' The input data, key and initialization vector are all represented as hexadecimal strings.
' @param  szInputHex  Hex-encoded input data.
' @param  szKeyHex    Hex-encoded key of exact key length.
' @param  szIvHex     Hex-encoded IV of exact block length, ignored for ECB mode or if `API_IV_PREFIX` is used (use `""`).
' @param  szAlgModePad  String with block cipher algorithm, mode and padding,
' e.g. `"aes128/cbc/pkcs5"`
' {@code
' Alg:  aes128|aes192|aes256|tdea|3des|desede3
' Mode: ecb|cbc|ofb|cfb|ctr
' Pad:  pkcs5|nopad|oneandzeroes|ansix923|w3c
' }
' @param  nOptions  Add `API_IV_PREFIX` to expect the IV to be prepended before the ciphertext in the input (not applicable for ECB mode).
' @return Decrypted plaintext in hex-encoded string or empty string on error.
' @remark Input data may be any even number of hex characters, but not zero.
' @remark Default padding is `Pkcs5` for ECB and CBC mode and `NoPad` for all other modes.
'**/
Public Function cipherDecryptHex(szInputHex As String, szKeyHex As String, szIvHex As String, szAlgModePad As String, Optional nOptions As Long = 0) As String
    Dim nc As Long
    nc = CIPHER_DecryptHex(vbNullString, 0, szInputHex, szKeyHex, szIvHex, szAlgModePad, nOptions)
    If nc <= 0 Then Exit Function
    cipherDecryptHex = String(nc, " ")
    nc = CIPHER_DecryptHex(cipherDecryptHex, nc, szInputHex, szKeyHex, szIvHex, szAlgModePad, nOptions)
    cipherDecryptHex = Left$(cipherDecryptHex, nc)
End Function

'/**
' Encrypts data in a byte array using the specified block cipher algorithm, mode and padding.
' The key and initialization vector are passed as byte arrays.
' @param  lpInput  Input data to be encrypted.
' @param  lpKey    Key of exact length for block cipher algorithm.
' @param  lpIV     Initialization Vector (IV) of exactly the block size or empty array for ECB mode.
' @param  szAlgModePad     String with block cipher algorithm, mode and padding,
' e.g. <code>"aes128/cbc/pkcs5"</code>
' {@code
' Alg:  aes128|aes192|aes256|tdea|3des|desede3
' Mode: ecb|cbc|ofb|cfb|ctr
' Pad:  pkcs5|nopad|oneandzeroes|ansix923|w3c
' }
' @param  nOptions  Add `API_IV_PREFIX` to prepend the IV before the ciphertext in the output
' (ignored for ECB mode).
' @return Ciphertext in byte array or empty array on error.
' @remark Default padding is `Pkcs5` for ECB and CBC mode and `NoPad` for all other modes.
'**/
Public Function cipherEncryptBytes(lpInput() As Byte, lpKey() As Byte, lpIV() As Byte, szAlgModePad As String, Optional nOptions As Long = 0) As Byte()
    cipherEncryptBytes = vbNullString
    Dim n1 As Long
    n1 = cnvBytesLen(lpInput)
    ' Fudge to allow an empty input array
    If n1 = 0 Then ReDim lpInput(0)
    Dim n2 As Long
    n2 = cnvBytesLen(lpKey)
    If n2 = 0 Then GoTo CleanUp
    Dim n3 As Long
    n3 = cnvBytesLen(lpIV)
    If n3 = 0 Then ReDim lpIV(0)
    Dim abMyData() As Byte
    Dim nb As Long
    nb = CIPHER_EncryptBytes(ByVal 0&, 0, lpInput(0), n1, lpKey(0), n2, lpIV(0), n3, szAlgModePad, nOptions)
    If nb <= 0 Then GoTo CleanUp
    ReDim abMyData(nb - 1)
    nb = CIPHER_EncryptBytes(abMyData(0), nb, lpInput(0), n1, lpKey(0), n2, lpIV(0), n3, szAlgModePad, nOptions)
    If nb <= 0 Then GoTo CleanUp
    ReDim Preserve abMyData(nb - 1)
    cipherEncryptBytes = abMyData
CleanUp:
    If n1 = 0 Then lpInput = vbNullString
    If n3 = 0 Then lpIV = vbNullString
End Function

' @deprecated Use cipherEncryptBytes()
Public Function cipherEncryptBytes2(lpInput() As Byte, lpKey() As Byte, lpIV() As Byte, szAlgModePad As String, Optional nOptions As Long = 0) As Byte()
    cipherEncryptBytes2 = cipherEncryptBytes(lpInput, lpKey, lpIV, szAlgModePad, nOptions)
End Function

'/**
' Encrypt hex-encoded data using the specified block cipher algorithm, mode and padding.
' The key and initialization vector are passed as hex-encoded strings.
' @param  szInputHex  Input data to be encrypted.
' @param  szKeyHex    Hex-encoded key of exact key length.
' @param  szIvHex     Hex-encoded IV of exact block length or `""` for ECB mode.
' @param  szAlgModePad  String with block cipher algorithm, mode and padding,
' e.g. `"aes128/cbc/pkcs5"`
' {@code
' Alg:  aes128|aes192|aes256|tdea|3des|desede3
' Mode: ecb|cbc|ofb|cfb|ctr
' Pad:  pkcs5|nopad|oneandzeroes|ansix923|w3c
' }
' @param  nOptions  Add `API_IV_PREFIX` to prepend the IV before the ciphertext in the output
' (ignored for ECB mode).
' @return Encrypted ciphertext in hex-encoded string or empty string on error.
' @remark Input data may be any even number of hex characters, but not zero.
' @remark Default padding is `Pkcs5` for ECB and CBC mode and `NoPad` for all other modes.
'**/
Public Function cipherEncryptHex(szInputHex As String, szKeyHex As String, szIvHex As String, szAlgModePad As String, Optional nOptions As Long = 0) As String
    Dim nc As Long
    nc = CIPHER_EncryptHex(vbNullString, 0, szInputHex, szKeyHex, szIvHex, szAlgModePad, nOptions)
    If nc <= 0 Then Exit Function
    cipherEncryptHex = String(nc, " ")
    nc = CIPHER_EncryptHex(cipherEncryptHex, nc, szInputHex, szKeyHex, szIvHex, szAlgModePad, nOptions)
    cipherEncryptHex = Left$(cipherEncryptHex, nc)
End Function

'/**
' Encrypt a file with block cipher.
' @param  szFileOut  Name of output file to be created or overwritten.
' @param  szFileIn  Name of input file.
' @param  lpKey    Key of exact length for block cipher algorithm.
' @param  lpIV     Initialization Vector (IV) of exactly the block size or empty array for ECB mode.
' @param  szAlgModePad     String with block cipher algorithm, mode and padding,
' e.g. <code>"aes128/cbc/pkcs5"</code>
' {@code
' Alg:  aes128|aes192|aes256|tdea|3des|desede3
' Mode: ecb|cbc|ofb|cfb|ctr
' Pad:  pkcs5|nopad|oneandzeroes|ansix923|w3c
' }
' @param  nOptions  Add `API_IV_PREFIX` to prepend the IV before the ciphertext in the output
' (ignored for ECB mode).
' @return Zero (0) on success, or a nonzero error code (use {@link apiErrorLookup} to check).
' @remark `szFileOut` and `szFileIn` must __not__ be the same.
'**/
Public Function cipherFileEncrypt(szFileOut As String, szFileIn As String, lpKey() As Byte, lpIV() As Byte, szAlgModePad As String, Optional nOptions As Long = 0) As Long
    Dim n2 As Long
    n2 = cnvBytesLen(lpKey)
    Dim n3 As Long
    n3 = cnvBytesLen(lpIV)
    If n3 = 0 Then ReDim lpIV(0)
    cipherFileEncrypt = CIPHER_FileEncrypt(szFileOut, szFileIn, lpKey(0), n2, lpIV(0), n3, szAlgModePad, nOptions)
CleanUp:
    If n3 = 0 Then lpIV = vbNullString
End Function

'/**
' Decrypt a file with block cipher.
' @param  szFileOut  Name of output file to be created or overwritten.
' @param  szFileIn  Name of input file.
' @param  lpKey    Key of exact length for block cipher algorithm.
' @param  lpIV     Initialization Vector (IV) of exactly the block size, or empty array for ECB mode or if IV already prefixed to input.
' @param  szAlgModePad     String with block cipher algorithm, mode and padding,
' e.g. <code>"aes128/cbc/pkcs5"</code>
' {@code
' Alg:  aes128|aes192|aes256|tdea|3des|desede3
' Mode: ecb|cbc|ofb|cfb|ctr
' Pad:  pkcs5|nopad|oneandzeroes|ansix923|w3c
' }
' @param  nOptions  Add `API_IV_PREFIX` to expect the IV to be prepended before the ciphertext in the input
' (ignored for ECB mode).
' @return Zero (0) on success, or a nonzero error code (use {@link apiErrorLookup} to check).
' @remark `szFileOut` and `szFileIn` must __not__ be the same.
'**/
Public Function cipherFileDecrypt(szFileOut As String, szFileIn As String, lpKey() As Byte, lpIV() As Byte, szAlgModePad As String, Optional nOptions As Long = 0) As Long
    Dim n2 As Long
    n2 = cnvBytesLen(lpKey)
    Dim n3 As Long
    n3 = cnvBytesLen(lpIV)
    If n3 = 0 Then ReDim lpIV(0)
    cipherFileDecrypt = CIPHER_FileDecrypt(szFileOut, szFileIn, lpKey(0), n2, lpIV(0), n3, szAlgModePad, nOptions)
CleanUp:
    If n3 = 0 Then lpIV = vbNullString
End Function

'/**
' Closes and clears the CIPHER context.
' @param  hContext Handle to the CIPHER context.
' @return Zero (0) on success, or a nonzero error code (use {@link apiErrorLookup} to check)
'**/
Public Function cipherFinal(hContext As Long) As Long
    cipherFinal = CIPHER_Final(hContext)
End Function

'/**
' Initializes the CIPHER context with the key, algorithm and mode ready for repeated incremental operations.
' @param  fEncrypt Direction flag: set as `ENCRYPT` (True) to encrypt or `DECRYPT` (False) to decrypt.
' @param szAlgAndMode String with block cipher algorithm and mode, e.g. `"aes128/cbc"`
' @param  lpKey Key of exact length for block cipher algorithm.
' @param  lpIV Initialization Vector (IV) of exactly the block size or empty array for ECB mode.
' @param  nOptions  Option flags, set as zero for defaults.
' @return Nonzero handle of the CIPHER context, or _zero_ if an error occurs.
'**/
Public Function cipherInit(fEncrypt As Integer, szAlgAndMode As String, lpKey() As Byte, lpIV() As Byte, Optional nOptions As Long = 0) As Long
    Dim n1 As Long
    n1 = cnvBytesLen(lpKey)
    Dim n2 As Long
    n2 = cnvBytesLen(lpIV)
    ' Fudge to allow an empty input array
    If n2 = 0 Then ReDim lpIV(0)
    cipherInit = CIPHER_Init(fEncrypt, szAlgAndMode, lpKey(0), n1, lpIV(0), n2, nOptions)
CleanUp:
    If n2 = 0 Then lpIV = vbNullString
End Function

'/**
' Initializes the CIPHER context with hex-encoded key, algorithm and mode ready for repeated incremental operations.
' @param  fEncrypt Direction flag: set as `ENCRYPT` (True) to encrypt or `DECRYPT` (False) to decrypt.
' @param szAlgAndMode String with block cipher algorithm and mode, e.g. `"aes128/cbc"`
' @param  szKeyHex Hex-encoded key of exact length for block cipher algorithm.
' @param  szIvHex  Hex-encoded initialization vector (IV) of exactly the block size or `""` for ECB mode.
' @param  nOptions Option flags, set as zero for defaults.
' @return Nonzero handle of the CIPHER context, or _zero_ if an error occurs.
'**/
Public Function cipherInitHex(fEncrypt As Integer, szAlgAndMode As String, szKeyHex As String, szIvHex As String, Optional nOptions As Long = 0) As Long
    cipherInitHex = CIPHER_InitHex(fEncrypt, szAlgAndMode, szKeyHex, szIvHex, nOptions)
End Function

'/**
' Encrypts or decrypts a chunk of input (in incremental mode).
' @param  hContext Handle to the CIPHER context.
' @param  lpData  Input data to be processed.
' @return Encrypted/decrypted block the same length as the input, or an empty array on error.
' @remark The input byte array must be a length exactly a multiple of the block size, except for the last chunk in CTR/OFB/CFB mode.
' Input in ECB/CBC mode must be suitably padded to the correct length.
'**/
Public Function cipherUpdate(hContext As Long, lpData() As Byte) As Byte()
    cipherUpdate = vbNullString
    Dim abMyData() As Byte
    Dim nb As Long
    nb = cnvBytesLen(lpData)
    If nb = 0 Then Exit Function
    ReDim abMyData(nb - 1)
    Dim r As Long
    r = CIPHER_Update(hContext, abMyData(0), nb, lpData(0), nb)
    If r <> 0 Then Exit Function
    cipherUpdate = abMyData
End Function

'/**
' Encrypts or decrypts a chunk of hex-encoded input (in incremental mode).
' @param  hContext Handle to the CIPHER context.
' @param  szInputHex Hex-encoded input data.
' @return Hex-encoded encrypted/decrypted chunk the same length as the input, or empty string `""` on error.
' @remark The input must represent data of a length exactly a multiple of the block size, except for the last chunk in CTR/OFB/CFB mode.
' Input in ECB/CBC mode must be suitably padded to the correct length.
'**/
Public Function cipherUpdateHex(hContext As Long, szInputHex As String) As String
    Dim nc As Long
    nc = Len(szInputHex)
    cipherUpdateHex = String(nc, " ")
    Call CIPHER_UpdateHex(hContext, cipherUpdateHex, nc, szInputHex)
End Function

'/**
' Unwraps (decrypts) key material with a key-encryption key.
' @param  lpData Wrapped key.
' @param  lpKek Key encryption key.
' @param  nOptions Algorithm to be used. Select one from:
' {@code
' API_BC_AES128
' API_BC_AES192
' API_BC_AES256
' API_BC_3DES
' }
' @return Unwrapped key material (or empty array on error).
'**/
Public Function cipherKeyUnwrap(lpData() As Byte, lpKek() As Byte, nOptions As Long) As Byte()
    cipherKeyUnwrap = vbNullString
    Dim n1 As Long
    n1 = cnvBytesLen(lpData)
    If n1 = 0 Then Exit Function
    Dim n2 As Long
    n2 = cnvBytesLen(lpKek)
    If n2 = 0 Then Exit Function
    Dim abMyData() As Byte
    Dim nb As Long
    nb = CIPHER_KeyUnwrap(ByVal 0&, 0, lpData(0), n1, lpKek(0), n2, nOptions)
    If nb <= 0 Then Exit Function
    ReDim abMyData(nb - 1)
    nb = CIPHER_KeyUnwrap(abMyData(0), nb, lpData(0), n1, lpKek(0), n2, nOptions)
    If nb <= 0 Then Exit Function
    ReDim Preserve abMyData(nb - 1)
    cipherKeyUnwrap = abMyData
End Function

'/**
' Wraps (encrypts) key material with a key-encryption key.
' @param  lpData Key material to be wrapped.
' @param  lpKek Key encryption key.
' @param  nOptions Algorithm to be used. Select one from:
' {@code
' API_BC_AES128
' API_BC_AES192
' API_BC_AES256
' API_BC_3DES
' }
' @return Wrapped key (or empty array on error).
'**/
Public Function cipherKeyWrap(lpData() As Byte, lpKek() As Byte, nOptions As Long) As Byte()
    cipherKeyWrap = vbNullString
    Dim n1 As Long
    n1 = cnvBytesLen(lpData)
    If n1 = 0 Then Exit Function
    Dim n2 As Long
    n2 = cnvBytesLen(lpKek)
    If n2 = 0 Then Exit Function
    Dim abMyData() As Byte
    Dim nb As Long
    nb = CIPHER_KeyWrap(ByVal 0&, 0, lpData(0), n1, lpKek(0), n2, nOptions)
    If nb <= 0 Then Exit Function
    ReDim abMyData(nb - 1)
    nb = CIPHER_KeyWrap(abMyData(0), nb, lpData(0), n1, lpKek(0), n2, nOptions)
    If nb <= 0 Then Exit Function
    ReDim Preserve abMyData(nb - 1)
    cipherKeyWrap = abMyData
End Function

'/**
' Encipher data in array of bytes using specified stream cipher.
' @param  lpInput Input data.
' @param  lpKey Key.
' @param  lpIV Initialization Vector (IV, nonce). Use an empty array for Arcfour.
' @param  nOptions Algorithm to be used. Select one from:
' {@code
' API_SC_ARCFOUR
' API_SC_SALSA20
' API_SC_CHACHA20
' }
' @param  nCounter Counter value for ChaCha20 only, otherwise ignored.
' @return Ciphertext in byte array, or empty array on error.
' @remark _Arcfour:_ any length key; use an empty array for IV.<br>
' @remark _Salsa20:_ key must be exactly 16 or 32 bytes and IV exactly 8 bytes long.<br>
' @remark _ChaCha20:_ key must be exactly 16 or 32 bytes and IV exactly 8, 12, or 16 bytes long. Counter is ignored if IV is 16 bytes.<br>
' @remark Note different order of parameters from core function.
'**/
Public Function cipherStreamBytes(lpInput() As Byte, lpKey() As Byte, lpIV() As Byte, nOptions As Long, Optional nCounter As Long = 0) As Byte()
    cipherStreamBytes = vbNullString
    Dim n1 As Long
    n1 = cnvBytesLen(lpInput)
    If n1 = 0 Then Exit Function
    Dim n2 As Long
    n2 = cnvBytesLen(lpKey)
    If n2 = 0 Then Exit Function
    Dim n3 As Long
    n3 = cnvBytesLen(lpIV)
    ' Fudge to allow an empty input array
    If n3 = 0 Then ReDim lpIV(0)
    ' Output is always same length as input
    Dim abMyData() As Byte
    ReDim abMyData(n1 - 1)
    Dim r As Long
    ' NB order of parameters (nCounter <=> nOptions)
    r = CIPHER_StreamBytes(abMyData(0), lpInput(0), n1, lpKey(0), n2, lpIV(0), n3, nCounter, nOptions)
    If r <> 0 Then GoTo CleanUp
    cipherStreamBytes = abMyData
CleanUp:
    If n3 = 0 Then lpIV = vbNullString
End Function


'/**
' Encipher data in a hex-encoded string using specified stream cipher.
' @param  szInputHex Hex-encoded input data.
' @param  szKeyHex Hex-encoded key.
' @param  szIvHex Hex-encoded Initialization Vector (IV, nonce). Use "" for Arcfour.
' @param  nOptions Algorithm to be used. Select one from:
' {@code
' API_SC_ARCFOUR
' API_SC_SALSA20
' API_SC_CHACHA20
' }
' @param  nCounter Counter value for ChaCha20 only, otherwise ignored.
' @return Ciphertext in hex-encoded string or empty string on error.
' @remark _Arcfour:_ any length key; specify "" for IV.<br>
' @remark _Salsa20:_ key must be exactly 16 or 32 bytes and IV exactly 8 bytes long.<br>
' @remark _ChaCha20:_ key must be exactly 16 or 32 bytes and IV exactly 8, 12, or 16 bytes long. Counter is ignored if IV is 16 bytes.<br>
' @remark Note different order of parameters from core function.
' @example
' {@code
' ' Test vector 3
' Debug.Print cipherStreamHex("00000000000000000000", "ef012345", "", API_SC_ARCFOUR)
' ' OK=D6A141A7EC3C38DFBD61
' }
'**/
Public Function cipherStreamHex(szInputHex As String, szKeyHex As String, szIvHex As String, nOptions As Long, Optional nCounter As Long = 0) As String
    Dim nc As Long
    nc = Len(szInputHex)
    cipherStreamHex = String(nc, " ")
    ' NB order of parameters (nCounter <=> nOptions)
    Call CIPHER_StreamHex(cipherStreamHex, nc, szInputHex, szKeyHex, szIvHex, nCounter, nOptions)
End Function

'/**
' Encipher data in a file using specified stream cipher.
' @param  szFileOut Output file to be created.
' @param  szFileIn Input file to be processed.
' @param  lpKey Key.
' @param  lpIV Initialization Vector (IV, nonce). Use an empty array for Arcfour.
' @param  nOptions Algorithm to be used. Select one from:
' {@code
' API_SC_ARCFOUR
' API_SC_SALSA20
' API_SC_CHACHA20
' }
' @param  nCounter Counter value for ChaCha20 only, otherwise ignored.
' @return Zero (0) on success, or a nonzero error code (use {@link apiErrorLookup} to check)
' @remark _Arcfour:_ any length key; use an empty array for IV.<br>
' @remark _Salsa20:_ key must be exactly 16 or 32 bytes and IV exactly 8 bytes long.<br>
' @remark _ChaCha20:_ key must be exactly 16 or 32 bytes and IV exactly 8, 12, or 16 bytes long. Counter is ignored if IV is 16 bytes.<br>
' @remark Note different order of parameters from core function.
'**/
Public Function cipherStreamFile(szFileOut As String, szFileIn As String, lpKey() As Byte, lpIV() As Byte, nOptions As Long, Optional nCounter As Long = 0) As Long
    Dim n2 As Long
    n2 = cnvBytesLen(lpKey)
    If n2 = 0 Then Exit Function
    Dim n3 As Long
    n3 = cnvBytesLen(lpIV)
    ' Fudge to allow an empty input array
    If n3 = 0 Then ReDim lpIV(0)
    Dim r As Long
    ' NB order of parameters (nCounter <=> nOptions)
    cipherStreamFile = CIPHER_StreamFile(szFileOut, szFileIn, lpKey(0), n2, lpIV(0), n3, nCounter, nOptions)
    If n3 = 0 Then lpIV = vbNullString
End Function

'/**
' Initialize the CIPHERSTREAM context ready for repeated operations of {@link cipherStreamUpdate}.
' @param  lpKey Key.
' @param  lpIV Initialization Vector (IV, nonce). Use an empty array for Arcfour.
' @param  nOptions Algorithm to be used. Select one from:
' {@code
' API_SC_ARCFOUR
' API_SC_SALSA20
' API_SC_CHACHA20
' }
' @param  nCounter Counter value for ChaCha20 only, otherwise ignored.
' @return Nonzero handle of the CIPHERSTREAM context, or _zero_ if an error occurs.
' @remark _Arcfour:_ any length key; use an empty array for IV.<br>
' @remark _Salsa20:_ key must be exactly 16 or 32 bytes and IV exactly 8 bytes long.<br>
' @remark _ChaCha20:_ key must be exactly 16 or 32 bytes and IV exactly 8, 12, or 16 bytes long. Counter is ignored if IV is 16 bytes.<br>
' @remark Note different order of parameters from core function.
'**/
Public Function cipherStreamInit(lpKey() As Byte, lpIV() As Byte, nOptions As Long, Optional nCounter As Long = 0) As Long
    Dim n2 As Long
    n2 = cnvBytesLen(lpKey)
    If n2 = 0 Then Exit Function
    Dim n3 As Long
    n3 = cnvBytesLen(lpIV)
    ' Fudge to allow an empty input array
    If n3 = 0 Then ReDim lpIV(0)
    Dim r As Long
    ' NB order of parameters (nCounter <=> nOptions)
    cipherStreamInit = CIPHER_StreamInit(lpKey(0), n2, lpIV(0), n3, nCounter, nOptions)
    If n3 = 0 Then lpIV = vbNullString
End Function

'/**
' Encrypt input using current CIPHERSTREAM context (in incremental mode).
' @param hContext Handle to the CIPHERSTREAM context.
' @param lpData Data to be encrypted.
' @return Encrypted array of exactly the same length as input; or an empty array on error
'**/
Public Function cipherStreamUpdate(hContext As Long, lpData() As Byte) As Byte()
    cipherStreamUpdate = vbNullString
    Dim nb As Long
    nb = cnvBytesLen(lpData)
    If nb = 0 Then Exit Function
    Dim abMyData() As Byte
    ReDim abMyData(nb - 1)
    nb = CIPHER_StreamUpdate(hContext, abMyData(0), lpData(0), nb)
    If nb <> 0 Then Exit Function
    cipherStreamUpdate = abMyData
End Function

'/**
' Close the CIPHERSTREAM context and destroy the key.
' @param hContext Handle to the CIPHERSTREAM context.
'**/
Public Function cipherStreamFinal(hContext As Long) As Long
    cipherStreamFinal = CIPHER_StreamFinal(hContext)
End Function

'/**
' Compute the CRC-32 checksum of an array of bytes.
' @param  lpInput Input data.
' @return CRC-32 checksum as a 32-bit integer value.
'**/
Public Function crcBytes(lpInput() As Byte) As Long
    crcBytes = CRC_Bytes(lpInput(0), cnvBytesLen(lpInput), 0)
End Function

'/**
' Compute the CRC-32 checksum of an ANSI string.
' @param  szInput Input data.
' @return CRC-32 checksum as a 32-bit integer value.
'**/
Public Function crcString(szInput As String) As Long
    crcString = CRC_String(szInput, 0)
End Function

'/**
' Compute the CRC-32 checksum of a file.
' @param  szFileName Name of input file.
' @return CRC-32 checksum as a 32-bit integer value.
'**/
Public Function crcFile(szFileName As String) As Long
    crcFile = CRC_File(szFileName, 0)
End Function

'/**
' Compute hash digest in byte format of byte input.
' @param  lpMessage Message to be digested in byte array.
' @param  nOptions  Algorithm to be used. Select one from:
' {@code
' API_HASH_SHA1
' API_HASH_SHA224
' API_HASH_SHA256
' API_HASH_SHA384
' API_HASH_SHA512
' API_HASH_SHA3_224
' API_HASH_SHA3_256
' API_HASH_SHA3_384
' API_HASH_SHA3_512
' API_HASH_MD5
' API_HASH_MD2
' API_HASH_RMD160
' }
' @return Message digest in byte array.
'**/
Public Function hashBytes(lpMessage() As Byte, nOptions As Long) As Byte()
    hashBytes = vbNullString
    Dim n1 As Long
    n1 = cnvBytesLen(lpMessage)
    ' Fudge to allow an empty input array
    If n1 = 0 Then ReDim lpMessage(0)
    Dim abMyData() As Byte
    Dim nb As Long
    nb = HASH_Bytes(ByVal 0&, 0, lpMessage(0), n1, nOptions)
    If nb <= 0 Then GoTo CleanUp
    ReDim abMyData(nb - 1)
    nb = HASH_Bytes(abMyData(0), nb, lpMessage(0), n1, nOptions)
    If nb <= 0 Then GoTo CleanUp
    ReDim Preserve abMyData(nb - 1)
    hashBytes = abMyData
CleanUp:
    Dim abMyDummy() As Byte
    If n1 = 0 Then lpMessage = abMyDummy
End Function

'/**
' Compute hash digest in byte format of a file.
' @param  szFileName Name of file containing message data.
' @param  nOptions  Algorithm to be used. Select one from:
' {@code
' API_HASH_SHA1
' API_HASH_SHA224
' API_HASH_SHA256
' API_HASH_SHA384
' API_HASH_SHA512
' API_HASH_SHA3_224
' API_HASH_SHA3_256
' API_HASH_SHA3_384
' API_HASH_SHA3_512
' API_HASH_MD5
' API_HASH_MD2
' API_HASH_RMD160
' }
' Add `API_HASH_MODE_TEXT` to hash in "text" mode instead of default "binary" mode.
' @return Message digest in byte array.
' @remark The default mode is "binary" where each byte is treated individually.
' In "text" mode CR-LF pairs will be treated as a single newline (LF) character.
'**/
Public Function hashFile(szFileName As String, nOptions As Long) As Byte()
    hashFile = vbNullString
    Dim abMyData() As Byte
    Dim nb As Long
    nb = HASH_File(ByVal 0&, 0, szFileName, nOptions)
    If nb <= 0 Then Exit Function
    ReDim abMyData(nb - 1)
    nb = HASH_File(abMyData(0), nb, szFileName, nOptions)
    If nb <= 0 Then Exit Function
    ReDim Preserve abMyData(nb - 1)
    hashFile = abMyData
End Function

'/**
' Compute hash digest in hex format from bit-oriented input.
' @param  lpData Bit-oriented message data in byte array.
' @param  nDataBitLen length of the message data in _bits_.
' @param  nOptions  Algorithm to be used. Select one from:
' {@code
' API_HASH_SHA1
' API_HASH_SHA224
' API_HASH_SHA256
' API_HASH_SHA384
' API_HASH_SHA512
' API_HASH_SHA3_224
' API_HASH_SHA3_256
' API_HASH_SHA3_384
' API_HASH_SHA3_512
' }
' @return Message digest in hex-encoded format.
' @remark Only the SHA family of hash functions is supported in bit-oriented mode.
' @remark Pass a bitstring as an array of bytes in `lpData` in big-endian order with the most-significant bit first.
' The bitstring will be truncated to the number of bits specified in `nDataBitLen` and extraneous bits on the right will be ignored.
' @example
' {@code
' ' NIST SHAVS CAVS 11.0 "SHA-1 ShortMsg" information
' lpData = cnvBytesFromHexStr("5180")  ' 9-bit bitstring = 0101 0001 1
' strDigest = hashHexFromBits(lpData, 9, API_HASH_SHA1)
' Debug.Print "MD = " & strDigest
' Debug.Print "OK = 0f582fa68b71ecdf1dcfc4946019cf5a18225bd2"
' }
'**/
Public Function hashHexFromBits(lpData() As Byte, nDataBitLen As Long, nOptions As Long) As String
    Dim nc As Long
    Dim s As String
    nc = API_MAX_HASH_CHARS
    ' Check bits length is not too long
    If cnvBytesLen(lpData) < (nDataBitLen + 7) \ 8 Then
        nDataBitLen = -1    ' Fudge to cause an error
    End If
    s = String(nc, " ")
    nc = HASH_HexFromBits(s, nc, lpData(0), nDataBitLen, nOptions)
    If nc <= 0 Then Exit Function
    hashHexFromBits = Left$(s, nc)
End Function

'/**
' Compute hash digest in hex format of byte input.
' @param  lpMessage Message to be digested in byte array.
' @param  nOptions  Algorithm to be used. Select one from:
' {@code
' API_HASH_SHA1
' API_HASH_SHA224
' API_HASH_SHA256
' API_HASH_SHA384
' API_HASH_SHA512
' API_HASH_SHA3_224
' API_HASH_SHA3_256
' API_HASH_SHA3_384
' API_HASH_SHA3_512
' API_HASH_MD5
' API_HASH_MD2
' API_HASH_RMD160
' }
' @return Message digest in hex-encoded format.
'**/
Public Function hashHexFromBytes(lpMessage() As Byte, nOptions As Long) As String
    Dim n1 As Long
    n1 = cnvBytesLen(lpMessage)
    ' Fudge to allow empty input array
    If n1 = 0 Then ReDim lpMessage(0)
    Dim nc As Long
    nc = HASH_HexFromBytes(vbNullString, 0, lpMessage(0), n1, nOptions)
    If nc <= 0 Then GoTo CleanUp
    hashHexFromBytes = String(nc, " ")
    nc = HASH_HexFromBytes(hashHexFromBytes, nc, lpMessage(0), n1, nOptions)
    hashHexFromBytes = Left$(hashHexFromBytes, nc)
CleanUp:
    If n1 = 0 Then lpMessage = vbNullString
End Function

'/**
' Compute hash digest in hex format of a file.
' @param  szFileName Name of file containing message data.
' @param  nOptions  Algorithm to be used. Select one from:
' {@code
' API_HASH_SHA1
' API_HASH_SHA224
' API_HASH_SHA256
' API_HASH_SHA384
' API_HASH_SHA512
' API_HASH_SHA3_224
' API_HASH_SHA3_256
' API_HASH_SHA3_384
' API_HASH_SHA3_512
' API_HASH_MD5
' API_HASH_MD2
' API_HASH_RMD160
' }
' Add `API_HASH_MODE_TEXT` to hash in "text" mode instead of default "binary" mode.
' @return Message digest in hex-encoded format.
' @remark The default mode is "binary" where each byte is treated individually.
' In "text" mode CR-LF pairs will be treated as a single newline (LF) character.
'**/
Public Function hashHexFromFile(szFileName As String, nOptions As Long) As String
    Dim nc As Long
    nc = HASH_HexFromFile(vbNullString, 0, szFileName, nOptions)
    If nc <= 0 Then Exit Function
    hashHexFromFile = String(nc, " ")
    nc = HASH_HexFromFile(hashHexFromFile, nc, szFileName, nOptions)
    hashHexFromFile = Left$(hashHexFromFile, nc)
End Function

'/**
' Compute hash digest in hex-encoded format from hex-encoded input.
' @param  szMsgHex Message to be digested in hex-encoded format.
' @param  nOptions  Algorithm to be used. Select one from:
' {@code
' API_HASH_SHA1
' API_HASH_SHA224
' API_HASH_SHA256
' API_HASH_SHA384
' API_HASH_SHA512
' API_HASH_SHA3_224
' API_HASH_SHA3_256
' API_HASH_SHA3_384
' API_HASH_SHA3_512
' API_HASH_MD5
' API_HASH_MD2
' API_HASH_RMD160
' }
' @return Message digest in hex-encoded format.
'**/
Public Function hashHexFromHex(szMsgHex As String, nOptions As Long) As String
    Dim nc As Long
    nc = HASH_HexFromHex(vbNullString, 0, szMsgHex, nOptions)
    If nc <= 0 Then Exit Function
    hashHexFromHex = String(nc, " ")
    nc = HASH_HexFromHex(hashHexFromHex, nc, szMsgHex, nOptions)
    hashHexFromHex = Left$(hashHexFromHex, nc)
End Function

'/**
' Compute hash digest in hex-encoded format from hex-encoded input.
' @param  szMessage Message data string.
' @param  nOptions  Algorithm to be used. Select one from:
' {@code
' API_HASH_SHA1
' API_HASH_SHA224
' API_HASH_SHA256
' API_HASH_SHA384
' API_HASH_SHA512
' API_HASH_SHA3_224
' API_HASH_SHA3_256
' API_HASH_SHA3_384
' API_HASH_SHA3_512
' API_HASH_MD5
' API_HASH_MD2
' API_HASH_RMD160
' }
' @return Message digest in hex-encoded format.
'**/
Public Function hashHexFromString(szMessage As String, nOptions As Long) As String
    Dim nc As Long
    nc = API_MAX_HASH_CHARS
    hashHexFromString = String(nc, " ")
    nc = HASH_HexFromString(hashHexFromString, nc, szMessage, Len(szMessage), nOptions)
    If nc <= 0 Then Exit Function
    hashHexFromString = Left$(hashHexFromString, nc)
End Function

'/**
' Initialise the HASH context ready for incremental operations.
' @param  nAlg  Algorithm to be used. Select one from:
' {@code
' API_HASH_SHA1
' API_HASH_SHA224
' API_HASH_SHA256
' API_HASH_SHA384
' API_HASH_SHA512
' API_HASH_SHA3_224
' API_HASH_SHA3_256
' API_HASH_SHA3_384
' API_HASH_SHA3_512
' }
' @return Nonzero handle of the HASH context, or _zero_ if an error occurs.
' @remark Only the SHA-1, SHA-2 and SHA-3 families of hash algorithms are supported in incremental mode.
' While the context handle is valid, add data to be digested in blocks of any length using
' {@link hashAddBytes} or {@link hashAddString}.
'**/
Public Function hashInit(Optional nAlg As Long = 0) As Long
    hashInit = HASH_Init(nAlg)
End Function


'/**
' Add an array of bytes to be digested.
' @param  hContext Handle to the HASH context.
' @param  lpData  Byte array containing the next part of the message.
' @return Zero (0) on success, or a nonzero error code (use {@link apiErrorLookup} to check)
'**/
Public Function hashAddBytes(hContext As Long, lpData() As Byte) As Long
    Dim n1 As Long
    n1 = cnvBytesLen(lpData)
    ' Fudge to allow an empty input array
    If n1 = 0 Then ReDim lpData(0)
    hashAddBytes = HASH_AddBytes(hContext, lpData(0), n1)
CleanUp:
    If n1 = 0 Then lpData = vbNullString
End Function

'/**
' Add a string to be digested.
' @param  hContext Handle to the HASH context.
' @param  szData  String containing the next part of the message.
' @return Zero (0) on success, or a nonzero error code (use {@link apiErrorLookup} to check)
'**/
Public Function hashAddString(hContext As Long, szData As String) As Long
    hashAddString = hashAddBytes(hContext, StrConv(szData, vbFromUnicode))
End Function

'/**
' Return the final message digest value.
' @param  hContext Handle to the HASH context.
' @return Digest in byte array or empty array on error.
' @remark Computes the result of all `hashAddBytes` and `hashAddString` calls since `HASH_Init`.
'**/
Public Function hashFinal(hContext As Long) As Byte()
    hashFinal = vbNullString
    Dim abMyData() As Byte
    Dim nb As Long
    nb = API_MAX_HASH_BYTES
    ReDim abMyData(nb - 1)
    nb = HASH_Final(abMyData(0), nb, hContext)
    If nb <= 0 Then Exit Function
    ReDim Preserve abMyData(nb - 1)
    hashFinal = abMyData
End Function

'/**
' Computes a keyed MAC in byte format from byte input.
' @param  lpMessage Message to be signed in byte format.
' @param  lpKey     Key in byte format.
' @param  nOptions  Algorithm to be used. Select one from:
' {@code
' API_HMAC_SHA1
' API_HMAC_SHA224
' API_HMAC_SHA256
' API_HMAC_SHA384
' API_HMAC_SHA512
' API_HMAC_MD5
' API_HMAC_RMD160
' API_HMAC_SHA3_224
' API_HMAC_SHA3_256
' API_HMAC_SHA3_384
' API_HMAC_SHA3_512
' API_CMAC_TDEA
' API_CMAC_AES128
' API_CMAC_AES192
' API_CMAC_AES256
' API_MAC_POLY1305
' API_KMAC_128
' API_KMAC_256
' }
' @return MAC value in hex-encoded format.
'**/
Public Function macBytes(lpMessage() As Byte, lpKey() As Byte, nOptions As Long) As Byte()
    macBytes = vbNullString
    Dim n1 As Long
    n1 = cnvBytesLen(lpMessage)
    ' Fudge to allow an empty input array
    If n1 = 0 Then ReDim lpMessage(0)
    Dim n2 As Long
    n2 = cnvBytesLen(lpKey)
    If n2 = 0 Then GoTo CleanUp
    Dim abMyData() As Byte
    Dim nb As Long
    nb = MAC_Bytes(ByVal 0&, 0, lpMessage(0), n1, lpKey(0), n2, nOptions)
    If nb <= 0 Then GoTo CleanUp
    ReDim abMyData(nb - 1)
    nb = MAC_Bytes(abMyData(0), nb, lpMessage(0), n1, lpKey(0), n2, nOptions)
    If nb <= 0 Then GoTo CleanUp
    ReDim Preserve abMyData(nb - 1)
    macBytes = abMyData
CleanUp:
    If n1 = 0 Then lpMessage = vbNullString
End Function

'/**
' Computes a keyed MAC in hex-encoded format from byte input.
' @param  lpMessage Message to be signed in byte format.
' @param  lpKey     Key in byte format.
' @param  nOptions  Algorithm to be used. Select one from:
' {@code
' API_HMAC_SHA1
' API_HMAC_SHA224
' API_HMAC_SHA256
' API_HMAC_SHA384
' API_HMAC_SHA512
' API_HMAC_MD5
' API_HMAC_RMD160
' API_HMAC_SHA3_224
' API_HMAC_SHA3_256
' API_HMAC_SHA3_384
' API_HMAC_SHA3_512
' API_CMAC_TDEA
' API_CMAC_AES128
' API_CMAC_AES192
' API_CMAC_AES256
' API_MAC_POLY1305
' API_KMAC_128
' API_KMAC_256
' }
' @return MAC value in hex-encoded format.
'**/
Public Function macHexFromBytes(lpMessage() As Byte, lpKey() As Byte, nOptions As Long) As String
    Dim n1 As Long
    n1 = cnvBytesLen(lpMessage)
    ' Fudge to allow empty input array
    If n1 = 0 Then ReDim lpMessage(0)
    Dim n2 As Long
    n2 = cnvBytesLen(lpKey)
    If n2 = 0 Then ReDim lpKey(0)
    Dim nc As Long
    nc = MAC_HexFromBytes(vbNullString, 0, lpMessage(0), n1, lpKey(0), n2, nOptions)
    If nc <= 0 Then GoTo CleanUp
    macHexFromBytes = String(nc, " ")
    nc = MAC_HexFromBytes(macHexFromBytes, nc, lpMessage(0), n1, lpKey(0), n2, nOptions)
    macHexFromBytes = Left$(macHexFromBytes, nc)
CleanUp:
    If n1 = 0 Then lpMessage = vbNullString
End Function

'/**
' Computes a keyed MAC in hex-encoded format from hex-encoded input.
' @param  szMsgHex Message to be signed in hex-encoded format.
' @param  szKeyHex Key in hex-encoded format.
' @param  nOptions  Algorithm to be used. Select one from:
' {@code
' API_HMAC_SHA1
' API_HMAC_SHA224
' API_HMAC_SHA256
' API_HMAC_SHA384
' API_HMAC_SHA512
' API_HMAC_MD5
' API_HMAC_RMD160
' API_HMAC_SHA3_224
' API_HMAC_SHA3_256
' API_HMAC_SHA3_384
' API_HMAC_SHA3_512
' API_CMAC_TDEA
' API_CMAC_AES128
' API_CMAC_AES192
' API_CMAC_AES256
' API_MAC_POLY1305
' API_KMAC_128
' API_KMAC_256
' }
' @return MAC value in hex-encoded format.
'**/
Public Function macHexFromHex(szMsgHex As String, szKeyHex As String, nOptions As Long) As String
    Dim nc As Long
    nc = MAC_HexFromHex(vbNullString, 0, szMsgHex, szKeyHex, nOptions)
    If nc <= 0 Then Exit Function
    macHexFromHex = String(nc, " ")
    nc = MAC_HexFromHex(macHexFromHex, nc, szMsgHex, szKeyHex, nOptions)
    macHexFromHex = Left$(macHexFromHex, nc)
End Function

'/**
' Initialises the MAC context ready to receive data to authenticate.
' @param  lpKey Key in byte format.
' @param  nAlg  Algorithm to be used. Select one from:
' {@code
' API_HMAC_SHA1
' API_HMAC_SHA224
' API_HMAC_SHA256
' API_HMAC_SHA384
' API_HMAC_SHA512
' }
' @return Nonzero handle of the MAC context, or _zero_ if an error occurs.
' @remark Once initialized, use the context for subsequent calls to `macAddBytes`, `macAddString` and `macFinal`.
'**/
Public Function macInit(lpKey() As Byte, nAlg As Long) As Long
    Dim n1 As Long
    n1 = cnvBytesLen(lpKey)
    macInit = MAC_Init(lpKey(0), n1, nAlg)
End Function

'/**
' Adds an array of bytes to be authenticated.
' @param  hContext Handle to the MAC context.
' @param  lpData  Byte array containing the next part of the message.
' @return Zero (0) on success, or a nonzero error code (use {@link apiErrorLookup} to check)
'**/
Public Function macAddBytes(hContext As Long, lpData() As Byte) As Long
    Dim n1 As Long
    n1 = cnvBytesLen(lpData)
    ' Fudge to allow an empty input array
    If n1 = 0 Then ReDim lpData(0)
    macAddBytes = MAC_AddBytes(hContext, lpData(0), n1)
CleanUp:
    If n1 = 0 Then lpData = vbNullString
End Function

'/**
' Adds a string to be authenticated.
' @param  hContext Handle to the MAC context.
' @param  szData  String containing the next part of the message.
' @return Zero (0) on success, or a nonzero error code (use {@link apiErrorLookup} to check)
'**/
Public Function macAddString(hContext As Long, szData As String) As Long
    macAddString = macAddBytes(hContext, StrConv(szData, vbFromUnicode))
End Function

'/**
' Returns the final MAC value.
' @param  hContext Handle to the MAC context.
' @return MAC in byte array or empty array on error.
' @remark Computes the result of all `macAddBytes` and `macAddString` calls since `macInit`.
'**/
Public Function macFinal(hContext As Long) As Byte()
    macFinal = vbNullString
    Dim abMyData() As Byte
    Dim nb As Long
    nb = API_MAX_MAC_CHARS
    ReDim abMyData(nb - 1)
    nb = MAC_Final(abMyData(0), nb, hContext)
    If nb <= 0 Then Exit Function
    ReDim Preserve abMyData(nb - 1)
    macFinal = abMyData
End Function

'/**
' Creates an input block suitably padded for encryption by a block cipher in ECB or CBC mode.
' @param  lpInput Plaintext bytes to be padded.
' @param  nBlkLen Cipher block length in bytes (8 or 16).
' @param  nOptions Use 0 for default PKCS5 padding or select one of:
' {@code
' API_PAD_1ZERO
' API_PAD_AX923
' API_PAD_W3C
' }
' @return Padded data in byte array.
'**/
Public Function padBytesBlock(lpInput() As Byte, nBlkLen As Long, Optional nOptions As Long = 0) As Byte()
    padBytesBlock = vbNullString
    Dim n1 As Long
    n1 = cnvBytesLen(lpInput)
    If n1 = 0 Then Exit Function
    Dim abMyData() As Byte
    Dim nb As Long
    nb = PAD_BytesBlock(ByVal 0&, 0, lpInput(0), n1, nBlkLen, nOptions)
    If nb <= 0 Then Exit Function
    ReDim abMyData(nb - 1)
    nb = PAD_BytesBlock(abMyData(0), nb, lpInput(0), n1, nBlkLen, nOptions)
    If nb <= 0 Then Exit Function
    ReDim Preserve abMyData(nb - 1)
    padBytesBlock = abMyData
End Function

'/**
' Creates a hex-encoded input block suitably padded for encryption by a block cipher in ECB or CBC mode.
' @param  szInput Hexadecimal-encoded data to be padded.
' @param  nBlkLen Cipher block length in bytes (8 or 16).
' @param  nOptions Use 0 for default PKCS5 padding or select one of:
' {@code
' API_PAD_1ZERO
' API_PAD_AX923
' API_PAD_W3C
' }
' @return Padded data in byte array.
'**/
Public Function padHexBlock(szInput As String, nBlkLen As Long, Optional nOptions As Long = 0) As String
    Dim nc As Long
    nc = PAD_HexBlock(vbNullString, 0, szInput & "", nBlkLen, nOptions)
    If nc <= 0 Then Exit Function
    padHexBlock = String(nc, " ")
    nc = PAD_HexBlock(padHexBlock, nc, szInput & "", nBlkLen, nOptions)
    padHexBlock = Left$(padHexBlock, nc)
End Function

'/**
' Removes the padding from an encryption block.
' @param  lpInput Padded data.
' @param  nBlkLen Cipher block length in bytes (8 or 16).
' @param  nOptions Use 0 for default PKCS5 padding or select one of:
' {@code
' API_PAD_1ZERO
' API_PAD_AX923
' API_PAD_W3C
' }
' @return Unpadded data in byte array or _unchanged_ data on error.
' @remark An error is indicated by returning the _original_ data which will always be longer than the expected unpadded result.
'**/
Public Function padUnpadBytes(lpInput() As Byte, nBlkLen As Long, Optional nOptions As Long = 0) As Byte()
    Dim lpOutput() As Byte
    Dim nb As Long
    padUnpadBytes = vbNullString
    ' No need to query for length because we know the output will be shorter than input
    ' so make sure output is as long as the input
    nb = cnvBytesLen(lpInput)
    If nb = 0 Then Exit Function
    ReDim lpOutput(nb - 1)
    nb = PAD_UnpadBytes(lpOutput(0), nb, lpInput(0), nb, nBlkLen, nOptions)
    If nb <= 0 Then Exit Function
    ' Re-dimension the output to the correct length
    ReDim Preserve lpOutput(nb - 1)
    padUnpadBytes = lpOutput
End Function

'/**
' Removes the padding from a hex-encoded encryption block.
' @param  szInput Hex-encoded padded data.
' @param  nBlkLen Cipher block length in bytes (8 or 16).
' @param  nOptions Use 0 for default PKCS5 padding or select one of:
' {@code
' API_PAD_1ZERO
' API_PAD_AX923
' API_PAD_W3C
' }
' @return Unpadded data in hex-encoded string or _unchanged_ data on error.
' @remark An error is indicated by returning the _original_ data which will always be longer than the expected unpadded result.
'**/
Public Function padUnpadHex(szInput As String, nBlkLen As Long, Optional nOptions As Long = 0) As String
    Dim nc As Long
    ' No need to query for length because we know the output will be shorter than input
    nc = Len(szInput)
    padUnpadHex = String(nc, " ")
    nc = PAD_UnpadHex(padUnpadHex, nc, szInput, nBlkLen, nOptions)
    If nc < 0 Then
        ' Return original input on error - user to catch
        padUnpadHex = szInput
        Exit Function
    End If
    padUnpadHex = Left$(padUnpadHex, nc)
End Function

'/**
' Derives a key of any length from a password using the PBKDF2 algorithm from PKCS#5 v2.1.
' @param  dkBytes Required length of key in bytes.
' @param  lpPwd Password encoded as byte array.
' @param  lpSalt Salt in a byte array.
' @param  nCount Iteration count.
' @param  nOptions  Hash algorithm to use in HMAC PRF. Select one from:
' {@code
' API_HMAC_SHA1
' API_HMAC_SHA224
' API_HMAC_SHA256
' API_HMAC_SHA384
' API_HMAC_SHA512
' API_HMAC_MD5
' }
' @return Key in byte array.
'**/
Public Function pbeKdf2(dkBytes As Long, lpPwd() As Byte, lpSalt() As Byte, nCount As Long, nOptions As Long) As Byte()
    Dim n1 As Long
    n1 = cnvBytesLen(lpPwd)
    If n1 = 0 Then ReDim lpPwd(0)
    Dim n2 As Long
    n2 = cnvBytesLen(lpSalt)
    If n2 = 0 Then ReDim lpSalt(0)
    Dim abMyData() As Byte
    pbeKdf2 = vbNullString
    ReDim abMyData(dkBytes - 1)
    Dim r As Long
    r = PBE_Kdf2(abMyData(0), dkBytes, lpPwd(0), n1, lpSalt(0), n2, nCount, nOptions)
    If r <> 0 Then Exit Function
    pbeKdf2 = abMyData
End Function

'/**
' Derives a hex-encoded key of any length from a password using the PBKDF2 algorithm from PKCS#5 v2.1. The salt and derived key are encoded in hexadecimal.
' @param  dkBytes Required length of key in bytes.
' @param  szPwd Password (as normal text).
' @param  szSaltHex Salt in hex-encoded format.
' @param  nCount Iteration count.
' @param  nOptions  Hash algorithm to use in HMAC PRF. Select one from:
' {@code
' API_HMAC_SHA1
' API_HMAC_SHA224
' API_HMAC_SHA256
' API_HMAC_SHA384
' API_HMAC_SHA512
' API_HMAC_MD5
' }
' @return Key in hex format.
'**/
Public Function pbeKdf2Hex(dkBytes As Long, szPwd As String, szSaltHex As String, nCount As Long, nOptions As Long) As String
    Dim dk As String
    dk = String(2 * dkBytes, " ")
    Call PBE_Kdf2Hex(dk, Len(dk), dkBytes, szPwd, szSaltHex, nCount, nOptions)
    pbeKdf2Hex = dk
End Function

'/**
' Derives a key of any length from a password using the SCRYPT algorithm from RFC7914.
' @param  dkBytes Required length of key in bytes.
' @param  lpPwd Password encoded as byte array.
' @param  lpSalt Salt in a byte array.
' @param  nParamN CPU/Memory cost parameter `N` (`"costParameter"`), a number greater than one and a power of 2.
' @param  nParamR Block size `r` (`"blockSize"`).
' @param  nParamP Parallelization parameter `p` (`"parallelizationParameter"`).
' @param  nOptions For future use.
' @return Key in byte array.
'**/
Public Function pbeScrypt(dkBytes As Long, lpPwd() As Byte, lpSalt() As Byte, nParamN As Long, nParamR As Long, nParamP As Long, Optional nOptions As Long = 0) As Byte()
    Dim n1 As Long
    n1 = cnvBytesLen(lpPwd)
    If n1 = 0 Then ReDim lpPwd(0)
    Dim n2 As Long
    n2 = cnvBytesLen(lpSalt)
    If n2 = 0 Then ReDim lpSalt(0)
    Dim abMyData() As Byte
    pbeScrypt = vbNullString
    ReDim abMyData(dkBytes - 1)
    Dim r As Long
    r = PBE_Scrypt(abMyData(0), dkBytes, lpPwd(0), n1, lpSalt(0), n2, nParamN, nParamR, nParamP, nOptions)
    If r <> 0 Then Exit Function
    pbeScrypt = abMyData
End Function

'/**
' Derives a hex-encoded key of any length from a password using the SCRYPT algorithm from RFC7914. The salt and derived key are encoded in hexadecimal.
' @param  dkBytes Required length of key in bytes.
' @param  szPwd Password (as normal text).
' @param  szSaltHex Salt in hex-encoded format.
' @param  nParamN CPU/Memory cost parameter `N` (`"costParameter"`), a number greater than one and a power of 2.
' @param  nParamR Block size `r` (`"blockSize"`).
' @param  nParamP Parallelization parameter `p` (`"parallelizationParameter"`).
' @param  nOptions For future use.
' @return Key in hex format.
' @example
' {@code
' Debug.Print pbeScryptHex(64, "password", "4E61436C", 1024, 8, 16)
' ' FDBABE1C9D3472007856E7190D01E9FE7C6AD7CBC8237830E77376634B3731622EAF30D92E22A3886FF109279D9830DAC727AFB94A83EE6D8360CBDFA2CC0640
' }
'**/
Public Function pbeScryptHex(dkBytes As Long, szPwd As String, szSaltHex As String, nParamN As Long, nParamR As Long, nParamP As Long, Optional nOptions As Long = 0) As String
    Dim dk As String
    dk = String(2 * dkBytes, " ")
    Call PBE_ScryptHex(dk, Len(dk), dkBytes, szPwd, szSaltHex, nParamN, nParamR, nParamP, nOptions)
    pbeScryptHex = dk
End Function

'/**
' Generate output bytes using a pseudorandom function (PRF).
' @param  nBytes Required number of output bytes.
' @param  lpMessage Input message data.
' @param  lpKey Key (expected 128 or 256 bits long).
' @param  nOptions  PRF function to be used. Select one from:
' {@code
' API_KMAC_128
' API_KMAC_256
' }
' @param  szCustom Customization string (optional).
' @return Output data in byte array.
' @remark The KMAC128 and KMAC256 PRF functions are described in NIST SP800-185 (_SHA-3 Derived Functions_),
' and use SHAKE128 and SHAKE256, respectively.
' @remark Note different order of parameters from core function.
'**/
Public Function prfBytes(nBytes As Long, lpMessage() As Byte, lpKey() As Byte, nOptions As Long, Optional szCustom As String = "") As Byte()
    Dim n1 As Long
    n1 = cnvBytesLen(lpMessage)
    If n1 = 0 Then ReDim lpMessage(0)
    Dim n2 As Long
    n2 = cnvBytesLen(lpKey)
    If n2 = 0 Then ReDim lpKey(0)
    Dim abMyData() As Byte
    prfBytes = vbNullString
    ReDim abMyData(nBytes - 1)
    nBytes = PRF_Bytes(abMyData(0), nBytes, lpMessage(0), n1, lpKey(0), n2, szCustom, nOptions)
    If nBytes <= 0 Then Exit Function
    prfBytes = abMyData
End Function

'/**
' Generate a random key value.
' @param  nBytes Required number of random bytes.
' @param szSeed User-supplied entropy in string format (optional).
' @return Array of random bytes.
'**/
Public Function rngKeyBytes(nBytes As Long, Optional szSeed As String = "") As Byte()
    Dim abMyData() As Byte
    rngKeyBytes = vbNullString
    If nBytes <= 0 Then Exit Function
    ReDim abMyData(nBytes - 1)
    Call RNG_KeyBytes(abMyData(0), nBytes, szSeed, Len(szSeed))
    rngKeyBytes = abMyData
End Function

'/**
' Generate a random key in hex format.
' @param  nBytes Required number of random bytes.
' @param szSeed User-supplied entropy in string format (optional).
' @return Random bytes in hex format.
'**/
Public Function rngKeyHex(nBytes As Long, Optional szSeed As String = "") As String
    Dim nc As Long
    If nBytes <= 0 Then Exit Function
    nc = nBytes * 2
    rngKeyHex = String(nc, " ")
    Call RNG_KeyHex(rngKeyHex, nc, nBytes, szSeed, Len(szSeed))
End Function

'/**
' Generate a random nonce.
' @param  nBytes Required number of random bytes.
' @return Array of random bytes.
'**/
Public Function rngNonce(nBytes As Long) As Byte()
    Dim abMyData() As Byte
    rngNonce = vbNullString
    If nBytes <= 0 Then Exit Function
    ReDim abMyData(nBytes - 1)
    Call RNG_NonceData(abMyData(0), nBytes)
    rngNonce = abMyData
End Function


'/**
' Generate bytes using an extendable-output function (XOF).
' @param  nBytes Required number of output bytes.
' @param  lpMessage Input message data.
' @param  nOptions  XOF algorithm to be used. Select one from:
' {@code
' API_XOF_SHAKE128
' API_XOF_SHAKE256
' }
' @return Output data in byte array.
'**/
Public Function xofBytes(nBytes As Long, lpMessage() As Byte, nOptions As Long) As Byte()
    Dim n1 As Long
    n1 = cnvBytesLen(lpMessage)
    If n1 = 0 Then ReDim lpMessage(0)
    Dim abMyData() As Byte
    xofBytes = vbNullString
    ReDim abMyData(nBytes - 1)
    nBytes = XOF_Bytes(abMyData(0), nBytes, lpMessage(0), n1, nOptions)
    If nBytes <= 0 Then Exit Function
    xofBytes = abMyData
End Function

'/**
' Compress data using the ZLIB deflate algorithm.
' @param  lpInput Data to be compressed.
' @return Compressed data.
'**/
Public Function zlibDeflate(lpInput() As Byte) As Byte()
    zlibDeflate = vbNullString
    Dim n1 As Long
    n1 = cnvBytesLen(lpInput)
    If n1 = 0 Then Exit Function
    Dim abMyData() As Byte
    Dim nb As Long
    nb = ZLIB_Deflate(ByVal 0&, 0, lpInput(0), n1)
    If nb <= 0 Then Exit Function
    ReDim abMyData(nb - 1)
    nb = ZLIB_Deflate(abMyData(0), nb, lpInput(0), n1)
    If nb <= 0 Then Exit Function
    ReDim Preserve abMyData(nb - 1)
    zlibDeflate = abMyData
End Function

'/**
' Inflate compressed data using the ZLIB algorithm.
' @param  lpInput Compressed data to be inflated.
' @return Uncompressed data, or an empty array on error.
' @remark An empty array may also be returned in the trivial case where the original data was the empty array itself.
'**/
Public Function zlibInflate(lpInput() As Byte) As Byte()
    zlibInflate = vbNullString
    Dim n1 As Long
    n1 = cnvBytesLen(lpInput)
    If n1 = 0 Then Exit Function
    Dim abMyData() As Byte
    Dim nb As Long
    nb = ZLIB_Inflate(ByVal 0&, 0, lpInput(0), n1)
    If nb <= 0 Then Exit Function
    ReDim abMyData(nb - 1)
    nb = ZLIB_Inflate(abMyData(0), nb, lpInput(0), n1)
    If nb <= 0 Then Exit Function
    ReDim Preserve abMyData(nb - 1)
    zlibInflate = abMyData
End Function


' ---------------------
' NEW in [v6.20]
' ---------------------

'/**
' Compress data using compression algorithm.
' @param  lpInput Data to be compressed.
' @param  nOptions  Compression algorithm to be used. Select one from:
' {@code
' API_COMPR_ZLIB (0)
' API_COMPR_ZSTD
' }
' @return Compressed data, or an empty array on error.
'**/
Public Function comprCompress(lpInput() As Byte, Optional nOptions As Long = API_COMPR_ZLIB) As Byte()
    comprCompress = vbNullString
    Dim n1 As Long
    n1 = cnvBytesLen(lpInput)
    If n1 = 0 Then Exit Function
    Dim abMyData() As Byte
    Dim nb As Long
    nb = COMPR_Compress(ByVal 0&, 0, lpInput(0), n1, nOptions)
    If nb <= 0 Then Exit Function
    ReDim abMyData(nb - 1)
    nb = COMPR_Compress(abMyData(0), nb, lpInput(0), n1, nOptions)
    If nb <= 0 Then Exit Function
    ReDim Preserve abMyData(nb - 1)
    comprCompress = abMyData
End Function

'/**
' Uncompress data using compression algorithm.
' @param  lpInput Data to be uncompressed.
' @param  nOptions  Compression algorithm to be used. Select one from:
' {@code
' API_COMPR_ZLIB (0)
' API_COMPR_ZSTD
' }
' @return Uncompressed data, or an empty array on error.
'**/
Public Function comprUncompress(lpInput() As Byte, Optional nOptions As Long = API_COMPR_ZLIB) As Byte()
    comprUncompress = vbNullString
    Dim n1 As Long
    n1 = cnvBytesLen(lpInput)
    If n1 = 0 Then Exit Function
    Dim abMyData() As Byte
    Dim nb As Long
    nb = COMPR_Uncompress(ByVal 0&, 0, lpInput(0), n1, nOptions)
    If nb <= 0 Then Exit Function
    ReDim abMyData(nb - 1)
    nb = COMPR_Uncompress(abMyData(0), nb, lpInput(0), n1, nOptions)
    If nb <= 0 Then Exit Function
    ReDim Preserve abMyData(nb - 1)
    comprUncompress = abMyData
End Function

'/**
' Encrypt or decrypt a block of data using Blowfish algorithm.
' @param  fEncrypt ENCRYPT (True) to encrypt, DECRYPT (False) to decrypt
' @param  lpInput Input data, length must be an exact multiple of 8.
' @param  lpKey Key of length between 1 and 56 bytes (448 bits)
' @param  lpIV Initialization vector of exactly 8 bytes
' @param  szMode Encryption mode. Select one from:
' {@code
' "ECB" "CBC" "CTR" "OFB" "CFB"
' }
' @return Output data, the same length as the input, or an empty array on error.
'**/
Public Function blfBytesBlock(fEncrypt As Integer, lpInput() As Byte, lpKey() As Byte, lpIV() As Byte, Optional szMode As String = "ECB") As Byte()
    blfBytesBlock = vbNullString
    Dim n1 As Long
    n1 = cnvBytesLen(lpInput)
    If n1 = 0 Then GoTo Done
    Dim n2 As Long
    n2 = cnvBytesLen(lpKey)
    If n2 = 0 Then GoTo Done
    Dim n3 As Long
    n3 = cnvBytesLen(lpIV)
    If n3 <> 8 Then GoTo Done
    ' Fudge for empty IV
    If n3 = 0 Then ReDim lpIV(0)
    Dim abMyData() As Byte
    ReDim abMyData(n1 - 1)
    Dim nb As Long
    nb = BLF_BytesMode(abMyData(0), lpInput(0), n1, lpKey(0), n2, fEncrypt, szMode, lpIV(0))
    If nb <> 0 Then GoTo CleanUp
    blfBytesBlock = abMyData
CleanUp:
    If n3 = 0 Then lpIV = vbNullString
Done:
    
End Function

'/**
' Zeroise data in memory.
' @param  lpData Data to be wiped.
' @return Zero (0) on success, or a nonzero error code (use {@link apiErrorLookup} to check).
' @remark This function does not free any memory; it just zeroises it.
'**/
Public Function wipeBytes(lpData() As Byte) As Long
    wipeBytes = WIPE_Bytes(lpData(0), cnvBytesLen(lpData))
End Function

'/**
' Zeroise a string.
' @param  szData String to be wiped.
' @return An empty string.
' @remark On its own this just zeroizes the string.
' To clear the string securely, do the following
' {@code
' strData = wipeString(strData)
' }
'**/
Public Function wipeString(szData As String) As String
    Call WIPE_String(szData, Len(szData))
    wipeString = vbNullString
End Function

'/**
' Securely wipes and deletes a file using 7-pass DOD standards.
' @param  szFileName File to be deleted.
' @param  nOptions Option flags. Select one of
' {@code
' API_WIPEFILE_DOD7 (0)
' API_WIPEFILE_SIMPLE
' }
' @return Zero (0) on success, or a nonzero error code (use {@link apiErrorLookup} to check).
' @remark The default option uses the 7-pass DOD Standard according to [NISPOM] before deleting.
' `API_WIPEFILE_SIMPLE` overwrites the file with a single pass of zero bytes (quicker but less secure).
'**/
Public Function wipeFile(szFileName As String, Optional nOptions As Long = 0) As Long
    wipeFile = WIPE_File(szFileName, nOptions)
End Function


' ... END OF MODULE
' *******************************************************************


