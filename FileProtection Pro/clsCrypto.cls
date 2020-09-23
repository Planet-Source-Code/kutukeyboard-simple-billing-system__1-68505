VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCrypto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' Algorithm classes
Private Const ALG_CLASS_ANY = 0
Private Const ALG_CLASS_SIGNATURE = 8192
Private Const ALG_CLASS_MSG_ENCRYPT = 16384
Private Const ALG_CLASS_DATA_ENCRYPT = 24576
Private Const ALG_CLASS_HASH = 32768
Private Const ALG_CLASS_KEY_EXCHANGE = 40960

' Algorithm types
Private Const ALG_TYPE_ANY = 0
Private Const ALG_TYPE_DSS = 512
Private Const ALG_TYPE_RSA = 1024
Private Const ALG_TYPE_BLOCK = 1536
Private Const ALG_TYPE_STREAM = 2048
Private Const ALG_TYPE_DH = 2560
Private Const ALG_TYPE_SECURECHANNEL = 3072

' RC2 sub-ids
Private Const ALG_SID_RC2 = 2

' Stream cipher sub-ids
Private Const ALG_SID_RC4 = 1
Private Const ALG_SID_SEAL = 2

' Diffie-Hellman sub-ids
Private Const ALG_SID_DH_SANDF = 1
Private Const ALG_SID_DH_EPHEM = 2
Private Const ALG_SID_AGREED_KEY_ANY = 3
Private Const ALG_SID_KEA = 4

' Hash sub ids
Private Const ALG_SID_MD2 = 1
Private Const ALG_SID_MD4 = 2
Private Const ALG_SID_MD5 = 3
Private Const ALG_SID_SHA = 4
Private Const ALG_SID_SHA1 = 4
Private Const ALG_SID_MAC = 5
Private Const ALG_SID_RIPEMD = 6
Private Const ALG_SID_RIPEMD160 = 7
Private Const ALG_SID_SSL3SHAMD5 = 8
Private Const ALG_SID_HMAC = 9

' algorithm identifier definitions
Private Const CALG_MD5 = ((ALG_CLASS_HASH Or ALG_TYPE_ANY) Or ALG_SID_MD5)
Private Const CALG_RC2 = ((ALG_CLASS_DATA_ENCRYPT Or ALG_TYPE_BLOCK) Or ALG_SID_RC2)
Private Const CALG_SSL3_SHAMD5 = ((ALG_CLASS_HASH Or ALG_TYPE_ANY) Or ALG_SID_SSL3SHAMD5)

' dwFlags definitions for CryptAcquireContext
Private Const CRYPT_VERIFYCONTEXT = &HF0000000
Private Const CRYPT_NEWKEYSET = &H8
Private Const CRYPT_DELETEKEYSET = &H10
Private Const CRYPT_MACHINE_KEYSET = &H20
Private Const HP_ALGID = &H1         '- Hash algorithm
Private Const HP_HASHVAL = &H2       '- Hash value
Private Const HP_HASHSIZE = &H4      '- Hash value size
Private Const HP_HMAC_INFO = &H5     '- Information for creating an HMAC

' dwFlag definitions for CryptGenKey
Private Const CRYPT_EXPORTABLE = &H1
Private Const CRYPT_USER_PROTECTED = &H2
Private Const CRYPT_CREATE_SALT = &H4
Private Const CRYPT_UPDATE_KEY = &H8
Private Const CRYPT_NO_SALT = &H10
Private Const CRYPT_PREGEN = &H40
Private Const CRYPT_RECIPIENT = &H10
Private Const CRYPT_INITIATOR = &H40
Private Const CRYPT_ONLINE = &H80
Private Const CRYPT_SF = &H100
Private Const CRYPT_CREATE_IV = &H200
Private Const CRYPT_KEK = &H400
Private Const CRYPT_DATA_KEY = &H800

' dwFlags definitions for CryptDeriveKey
Private Const CRYPT_SERVER = &H400
Private Const KEY_LENGTH_MASK = &HFFFF0000

' dwFlag definitions for CryptSetProviderEx and CryptGetDefaultProvider
Private Const CRYPT_MACHINE_DEFAULT = &H1
Private Const CRYPT_USER_DEFAULT = &H2
Private Const CRYPT_DELETE_DEFAULT = &H4
Private Const CRYPT_FAILED = 0
Private Const CRYPT_SUCCEED = -1

' CryptSetProvParam
Private Const PROV_RSA_FULL = 1
Private Const PROV_RSA_SIG = 2
Private Const PROV_DSS = 3
Private Const PROV_FORTEZZA = 4
Private Const PROV_MS_EXCHANGE = 5
Private Const PROV_SSL = 6
Private Const PROV_RSA_SCHANNEL = 12
Private Const PROV_DSS_DH = 13
Private Const PROV_EC_ECDSA_SIG = 14
Private Const PROV_EC_ECNRA_SIG = 15
Private Const PROV_EC_ECDSA_FULL = 16
Private Const PROV_EC_ECNRA_FULL = 17
Private Const PROV_SPYRUS_LYNKS = 20

' Provider friendly names
Private Const MS_DEF_PROV = "Microsoft Base Cryptographic Provider v1.0"
Private Const MS_ENHANCED_PROV = "Microsoft Enhanced Cryptographic Provider v1.0"
Private Const MS_DEF_RSA_SIG_PROV = "Microsoft RSA Signature Cryptographic Provider"
Private Const MS_DEF_RSA_SCHANNEL_PROV = "Microsoft Base RSA SChannel Cryptographic Provider"
Private Const MS_ENHANCED_RSA_SCHANNEL_PROV = "Microsoft Enhanced RSA SChannel Cryptographic Provider"
Private Const MS_DEF_DSS_PROV = "Microsoft Base DSS Cryptographic Provider"
Private Const MS_DEF_DSS_DH_PROV = "Microsoft Base DSS and Diffie-Hellman Cryptographic Provider"


'WinCrypt API Declarations
Private Declare Function CryptAcquireContext Lib "advapi32.dll" Alias "CryptAcquireContextA" ( _
    phProv As Long, _
    ByVal pszContainer As String, _
    ByVal pszProvider As String, _
    ByVal dwProvType As Long, _
    ByVal dwFlags As Long) As Long

Private Declare Function CryptReleaseContext Lib "advapi32.dll" ( _
    ByVal hProv As Long, _
    ByVal dwFlags As Long) As Long

Private Declare Function CryptDeriveKey Lib "advapi32.dll" ( _
    ByVal hProv As Long, _
    ByVal Algid As Long, _
    ByVal hBaseData As Long, _
    ByVal dwFlags As Long, _
    phKey As Long) As Long

Private Declare Function CryptDestroyKey Lib "advapi32.dll" ( _
    ByVal hKey As Long) As Long

Private Declare Function CryptEncrypt Lib "advapi32.dll" ( _
    ByVal hKey As Long, _
    ByVal hHash As Long, _
    ByVal Final As Long, _
    ByVal dwFlags As Long, _
    ByVal pbData As String, _
    pdwDataLen As Long, _
    ByVal dwBufLen As Long) As Long

Private Declare Function CryptDecrypt Lib "advapi32.dll" ( _
    ByVal hKey As Long, _
    ByVal hHash As Long, _
    ByVal Final As Long, _
    ByVal dwFlags As Long, _
    ByVal pbData As String, _
    pdwDataLen As Long) As Long

Private Declare Function CryptCreateHash Lib "advapi32.dll" ( _
    ByVal hProv As Long, _
    ByVal Algid As Long, _
    ByVal hKey As Long, _
    ByVal dwFlags As Long, _
    phHash As Long) As Long

Private Declare Function CryptHashData Lib "advapi32.dll" ( _
    ByVal hHash As Long, _
    ByVal pbData As String, _
    ByVal dwDataLen As Long, _
    ByVal dwFlags As Long) As Long

Private Declare Function CryptDestroyHash Lib "advapi32.dll" ( _
    ByVal hHash As Long) As Long
    
Private Declare Function CryptGetHashParam Lib "advapi32.dll" ( _
    ByVal hHash As Long, ByVal dwParam As Long, _
    ByVal pbData As String, pdwDataLen As Long, _
    ByVal dwFlags As Long) As Long
    
Private Declare Function CryptGetHashDWParam Lib "advapi32.dll" _
    Alias "CryptGetHashParam" (ByVal hHash As Long, _
    ByVal dwParam As Long, pbData As Long, _
    pdwDataLen As Long, ByVal dwFlags As Long) As Long

'constants from WinErr.h
Private Const ERROR_INVALID_HANDLE As Long = 6
Private Const ERROR_INVALID_PARAMETER As Long = 87
Private Const NTE_BAD_KEY As Long = &H80090003
Private Const NTE_BAD_UID As Long = &H80090001
Private Const NTE_NO_KEY As Long = &H8009000D
Private Const NTE_BAD_SIGNATURE As Long = &H80090006

' Private property buffers
Private m_sPassword As String   ' Password used to create encryption key
Private m_sInBuffer As String   ' Used as an input buffer for all data to be encrypted or decrypted
Private m_sOutBuffer As String  ' Used as an output buffer for all data that has been encrypted or decrypted
Private m_sErrorMsg As String   ' Error message string

' Private class-level variables
Private m_lHCryptProv As Long    ' Handle for the cryptographic service provider (CSP)
Private m_lHSessionKey As Long   ' Session key for encrypting and decrypting data
Private m_sHashBuffer As String  ' Used as a buffer for all the hash value
Private m_iHashLength As Integer ' Used to hold the hash length

Public Property Get InBuffer() As String
InBuffer = m_sInBuffer
End Property

Public Property Let InBuffer(vNewValue As String)
m_sInBuffer = vNewValue
End Property

Public Property Get OutBuffer() As String
OutBuffer = m_sOutBuffer
End Property

Public Property Get ErrorMsg() As String
ErrorMsg = m_sErrorMsg
End Property

Private Sub Class_Initialize()
If Not InitUser Then
  m_sErrorMsg = "Unable to initialize CryptoAPI."
  MsgBox m_sErrorMsg, vbOKOnly, "VB Crypto"
End If
End Sub

Private Function InitUser() As Boolean
On Error Resume Next
InitUser = False
Dim lDataSize As Long
Dim lResult As Long
Dim sResult As String
Dim sContainer As String
Dim sProvider As String
sContainer = vbNullChar ' Prepare string buffers
sProvider = MS_DEF_PROV & vbNullChar
' Attempt to acquire a handle to the default key container.
If Not CBool(CryptAcquireContext(m_lHCryptProv, sContainer, sProvider, PROV_RSA_FULL, 0)) Then
  ' Create default key container.
  If Not CBool(CryptAcquireContext(m_lHCryptProv, sContainer, sProvider, PROV_RSA_FULL, CRYPT_NEWKEYSET)) Then
    m_sErrorMsg = "Error creating key container - " & CStr(err.LastDllError)
    MsgBox m_sErrorMsg, vbOKOnly, "VB Crypto"
    Exit Function
  End If
End If
InitUser = True
Exit Function
End Function

Private Sub Class_Terminate()
Dim lResult As Long
If (m_lHCryptProv <> 0) Then lResult = CryptReleaseContext(m_lHCryptProv, 0)
End Sub

Public Function GeneratePasswordKey() As Boolean
On Error Resume Next
Dim lHash As Long
Dim lResult As Long
GeneratePasswordKey = False
' Create an empty hash object.
If Not CBool(CryptCreateHash(m_lHCryptProv, CALG_MD5, 0, 0, lHash)) Then
  m_sErrorMsg = "Error " & CStr(err.LastDllError) & " during CryptCreateHash!"
  MsgBox m_sErrorMsg, vbOKOnly, "VB Crypto"
  Exit Function
End If
' Hash the password string.
If Not CBool(CryptHashData(lHash, m_sPassword, Len(m_sPassword), 0)) Then
  m_sErrorMsg = "Error " & CStr(err.LastDllError) & " during CryptHashData!"
  MsgBox m_sErrorMsg, vbOKOnly, "VB Crypto"
  Exit Function
End If
' Create a derived block cipher session key.
If Not CBool(CryptDeriveKey(m_lHCryptProv, CALG_RC2, lHash, 0, m_lHSessionKey)) Then
  m_sErrorMsg = "Error " & CStr(err.LastDllError) & " during CryptDeriveKey!"
  MsgBox m_sErrorMsg, vbOKOnly, "VB Crypto"
  Exit Function
End If
' Destroy the hash object
If Not CBool(CryptDestroyHash(lHash)) Then
  m_sErrorMsg = "Error " & CStr(err.LastDllError) & " during CryptDestroyHash!"
  MsgBox m_sErrorMsg, vbOKOnly, "VB Crypto"
  Exit Function
End If
GeneratePasswordKey = True
End Function

Public Property Get Password() As String
Password = m_sPassword
End Property

Public Property Let Password(ByVal sNewValue As String)
m_sPassword = sNewValue
End Property

Public Function EncryptMessageData()
On Error Resume Next
EncryptMessageData = False
Dim lDataSize As Long
Dim lResult As Long
Dim sCryptBuffer As String
Dim lCryptLength As Long
Dim lCryptBufLen As Long
' Determine the size of the buffer needed for encrypting the data in the InBuffer property
lCryptLength = Len(m_sInBuffer)
lResult = CryptEncrypt(m_lHSessionKey, 0, 1, 0, vbNullString, lCryptLength, lCryptBufLen)
lCryptBufLen = lCryptLength * 2 ' Prepare a string buffer for the CryptEncrypt function
lCryptLength = Len(m_sInBuffer)
sCryptBuffer = String(lCryptBufLen, vbNullChar)
LSet sCryptBuffer = m_sInBuffer ' Copy in the contents of the InBuffer property
If Not CBool(CryptEncrypt(m_lHSessionKey, 0, 1, 0, sCryptBuffer, lCryptLength, lCryptBufLen)) Then ' Encrypt data
  m_sErrorMsg = "Error " & CStr(err.LastDllError) & " during CryptEncrypt!"
  MsgBox m_sErrorMsg, vbOKOnly, "VB Crypto"
  Exit Function
End If
m_sOutBuffer = Mid$(sCryptBuffer, 1, lCryptLength) ' Copy the encrypted data to the OutBuffer property
EncryptMessageData = True
End Function

Public Function DestroySessionKey() As Boolean
On Error Resume Next
Dim lResult As Long
DestroySessionKey = False
If Not CBool(CryptDestroyKey(m_lHSessionKey)) Then ' Destroy the session key
  m_sErrorMsg = "Error " & CStr(err.LastDllError) & " during CryptDestroyKey!"
  MsgBox m_sErrorMsg, vbOKOnly, "VB Crypto"
  Exit Function
End If
DestroySessionKey = True
End Function

Public Function DecryptMessageData() As Boolean
On Error Resume Next
DecryptMessageData = False
Dim lDataSize As Long
Dim lResult As Long
Dim sCryptBuffer As String
Dim lCryptLength As Long
Dim lCryptBufLen As Long
lCryptBufLen = Len(m_sInBuffer) ' Prepare sCryptBuffer for CryptDecrypt
sCryptBuffer = String(lCryptBufLen, vbNullChar)
LSet sCryptBuffer = m_sInBuffer
If Not CBool(CryptDecrypt(m_lHSessionKey, 0, 1, 0, sCryptBuffer, lCryptBufLen)) Then ' Decrypt data
  m_sErrorMsg = "Error " & CStr(err.LastDllError) & " during CryptDecrypt!"
  MsgBox m_sErrorMsg, vbOKOnly, "VB Crypto"
  Exit Function
End If
m_sOutBuffer = Mid$(sCryptBuffer, 1, lCryptBufLen) ' Apply decrypted string from sCryptBuffer to private buffer for OutBuffer property
DecryptMessageData = True
End Function

Public Function HashFile() As Boolean
On Error Resume Next
Dim lHash As Long
Dim lResult As Long
Dim sCryptBuffer As String
Dim lCryptBufLen As Long
Dim lCryptHashSize As Long
HashFile = False
If Not CBool(CryptCreateHash(m_lHCryptProv, CALG_MD5, 0, 0, lHash)) Then ' Create an empty hash object.
  m_sErrorMsg = "Error " & CStr(err.LastDllError) & " during CryptCreateHash!"
  MsgBox m_sErrorMsg, vbOKOnly, "VB Crypto"
  Exit Function
End If
If Not CBool(CryptHashData(lHash, m_sInBuffer, Len(m_sInBuffer), 0)) Then ' Hash the password string.
  m_sErrorMsg = "Error " & CStr(err.LastDllError) & " during CryptHashData!"
  MsgBox m_sErrorMsg, vbOKOnly, "VB Crypto"
  Exit Function
End If
lCryptBufLen = 4 ' Extract the hash value size.
If Not CBool(CryptGetHashDWParam(lHash, HP_HASHSIZE, lCryptHashSize, lCryptBufLen, 0)) Then
  m_sErrorMsg = "Error " & CStr(err.LastDllError) & " during CryptGetHashParam!"
  MsgBox m_sErrorMsg, vbOKOnly, "VB Crypto"
  Exit Function
End If
sCryptBuffer = String(lCryptHashSize, vbNullChar) ' Prepare sCryptBuffer for CryptDecrypt
lCryptBufLen = lCryptHashSize
If Not CBool(CryptGetHashParam(lHash, HP_HASHVAL, sCryptBuffer, lCryptBufLen, 0)) Then ' Extract the hash value.
  m_sErrorMsg = "Error " & CStr(err.LastDllError) & " during CryptGetHashParam!"
  MsgBox m_sErrorMsg, vbOKOnly, "VB Crypto"
  Exit Function
End If
If Not CBool(CryptDestroyHash(lHash)) Then ' Destroy the Hash object
  m_sErrorMsg = "Error " & CStr(err.LastDllError) & " during CryptDestroyHash!"
  MsgBox m_sErrorMsg, vbOKOnly, "VB Crypto"
  Exit Function
End If
m_sHashBuffer = Mid$(sCryptBuffer, 1, lCryptBufLen) ' Apply hash string from sCryptBuffer to private buffer for HashBuffer variables
m_iHashLength = lCryptBufLen
HashFile = True
End Function

Public Function EncryptFileData() As Boolean
On Error Resume Next
EncryptFileData = False
Dim strInputData As String
strInputData = Trim(Str(m_iHashLength)) + m_sHashBuffer + m_sInBuffer ' concatenate length of hash, hash value, and origianl file
m_sInBuffer = strInputData ' copy new strin to in buffer
EncryptFileData = EncryptMessageData ' call encryptmessagedata
End Function

Public Function DecryptFileData() As Boolean
On Error Resume Next
DecryptFileData = False
Dim strInputData As String
Dim strHash As String
Dim strHashLength As String
Dim lHashLength As Long
If Not DecryptMessageData Then Exit Function ' decrypt file data
strHashLength = Left(m_sOutBuffer, 2) ' extract original hash length
lHashLength = CLng(strHashLength)
strHash = Mid(m_sOutBuffer, 3, lHashLength) ' extract origianl hash value
strInputData = Mid(m_sOutBuffer, (3 + lHashLength)) ' extract original file contents
m_sInBuffer = strInputData ' copy origianl file contents to the input buffer
If Not HashFile Then Exit Function ' hash origianl file contents
If (strHash = m_sHashBuffer) Then ' compare tht original hash value to the new hash value
  m_sOutBuffer = strInputData ' file was un-altered
Else
  MsgBox "File has been corrupted since original encryption!", vbInformation, "File Error" ' hash values don't match
  m_sOutBuffer = ""
  Exit Function
End If
DecryptFileData = True
End Function
