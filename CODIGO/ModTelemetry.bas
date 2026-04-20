Attribute VB_Name = "ModTelemetry"
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

' ============================================================================
' Enums
' ============================================================================
Public Enum e_OtelStatus
    OTEL_STATUS_UNSET = 0
    OTEL_STATUS_OK = 1
    OTEL_STATUS_ERROR = 2
End Enum

Public Enum e_OtelSpanKind
    OTEL_SPAN_KIND_INTERNAL = 1
    OTEL_SPAN_KIND_CLIENT = 3
End Enum

' ============================================================================
' Constants
' ============================================================================
Private Const MAX_SPAN_ATTRIBUTES As Long = 8
Private Const MAX_ACTIVE_SPANS As Long = 16
Private Const MAX_BUFFER_SPANS As Long = 100

' ============================================================================
' UDTs
' ============================================================================
Public Type t_OtelAttribute
    Key As String
    Value As String
End Type

Public Type t_OtelSpan
    TraceId As String
    SpanId As String
    ParentSpanId As String
    Name As String
    Kind As e_OtelSpanKind
    StartTimeUnixNano As String
    EndTimeUnixNano As String
    StatusCode As e_OtelStatus
    StatusMessage As String
    Attributes(0 To 7) As t_OtelAttribute
    AttributeCount As Long
    Active As Boolean
End Type

' ============================================================================
' Module-level private state
' ============================================================================
Private m_Enabled As Boolean
Private m_CollectorURL As String
Private m_ServiceName As String
Private m_ServiceVersion As String
Private m_MachineId As String
Private m_ActiveSpans(0 To MAX_ACTIVE_SPANS - 1) As t_OtelSpan
Private m_Buffer(0 To MAX_BUFFER_SPANS - 1) As t_OtelSpan
Private m_BufferCount As Long
Private m_BatchSize As Long
Private m_FlushIntervalMs As Long
Private m_QPCFrequency As Currency
Private m_UnixEpochQPC As Currency
Private m_UnixEpochTime As Double
Private m_RndSeeded As Boolean

' ============================================================================
' ID Generation
' ============================================================================

''
' Generates a 32-character lowercase hex string (16 random bytes) for use as a trace ID.
'
' @return   32-char lowercase hex string
Public Function Telemetry_GenerateTraceId() As String
    On Error GoTo Telemetry_GenerateTraceId_Err
    Telemetry_GenerateTraceId = GenerateHexBytes(16)
    Exit Function
Telemetry_GenerateTraceId_Err:
    Call RegistrarError(Err.Number, Err.Description, "ModTelemetry.Telemetry_GenerateTraceId", Erl)
    Resume Next
End Function

''
' Generates a 16-character lowercase hex string (8 random bytes) for use as a span ID.
'
' @return   16-char lowercase hex string
Public Function Telemetry_GenerateSpanId() As String
    On Error GoTo Telemetry_GenerateSpanId_Err
    Telemetry_GenerateSpanId = GenerateHexBytes(8)
    Exit Function
Telemetry_GenerateSpanId_Err:
    Call RegistrarError(Err.Number, Err.Description, "ModTelemetry.Telemetry_GenerateSpanId", Erl)
    Resume Next
End Function

''
' Generates N random bytes as a lowercase hex string.
'
' @param    NumBytes    Number of random bytes to generate
' @return   Hex string of length NumBytes * 2
Private Function GenerateHexBytes(ByVal NumBytes As Long) As String
    Dim i As Long
    Dim ByteVal As Long
    Dim HexByte As String
    Dim Result As String
    
    Result = vbNullString
    
    For i = 1 To NumBytes
        ByteVal = Int(Rnd * 256)
        HexByte = Hex$(ByteVal)
        If Len(HexByte) = 1 Then
            HexByte = "0" & HexByte
        End If
        Result = Result & LCase$(HexByte)
    Next i
    
    GenerateHexBytes = Result
End Function

' ============================================================================
' Timestamp
' ============================================================================

''
' Returns the current time as a string of nanoseconds since Unix epoch.
' Uses QueryPerformanceCounter for sub-millisecond precision.
'
' @return   String representation of Unix epoch nanoseconds
Public Function Telemetry_GetUnixNano() As String
    On Error GoTo Telemetry_GetUnixNano_Err
    
    Dim CurrentQPC As Currency
    Dim DeltaQPC As Currency
    Dim DeltaSeconds As Double
    Dim UnixSeconds As Double
    Dim NanoStr As String
    
    Call QueryPerformanceCounter(CurrentQPC)
    
    ' Compute delta from calibration point
    DeltaQPC = CurrentQPC - m_UnixEpochQPC
    
    ' Convert QPC delta to seconds (Currency is scaled by 10000)
    DeltaSeconds = CDbl(DeltaQPC) / CDbl(m_QPCFrequency)
    
    ' Current Unix time in seconds
    UnixSeconds = m_UnixEpochTime + DeltaSeconds
    
    ' Convert to nanoseconds and return as string to avoid precision loss
    ' Use Decimal via CDec for full precision multiplication
    Dim NanoDec As Variant
    NanoDec = CDec(UnixSeconds) * CDec(1000000000#)
    
    ' Format without decimals
    NanoStr = CStr(Fix(NanoDec))
    
    ' Remove any negative sign that might appear from Fix on very small negatives
    If Left$(NanoStr, 1) = "-" Then NanoStr = Mid$(NanoStr, 2)
    
    Telemetry_GetUnixNano = NanoStr
    Exit Function
    
Telemetry_GetUnixNano_Err:
    Call RegistrarError(Err.Number, Err.Description, "ModTelemetry.Telemetry_GetUnixNano", Erl)
    Telemetry_GetUnixNano = "0"
    Resume Next
End Function

''
' Calibrates the QPC-to-Unix-time mapping.
' Records the current QPC value and the corresponding Unix epoch seconds.
' Called during Telemetry_Init.
Private Sub CalibrateTimestamp()
    On Error GoTo CalibrateTimestamp_Err
    
    ' Get QPC frequency
    Call QueryPerformanceFrequency(m_QPCFrequency)
    
    ' Record current Unix time in seconds since epoch
    m_UnixEpochTime = CDbl(DateDiff("s", #1/1/1970#, Now()))
    
    ' Record corresponding QPC value
    Call QueryPerformanceCounter(m_UnixEpochQPC)
    
    Exit Sub
    
CalibrateTimestamp_Err:
    Call RegistrarError(Err.Number, Err.Description, "ModTelemetry.CalibrateTimestamp", Erl)
    Resume Next
End Sub

' ============================================================================
' Settings
' ============================================================================

''
' Reads telemetry configuration from the [TELEMETRY] section of Configuracion.ini
' using the existing GetSetting function from ModSettings.bas.
Private Sub ReadTelemetrySettings()
    On Error GoTo ReadTelemetrySettings_Err
    
    Dim sValue As String
    
    ' Read Enabled flag — compare to "1"
    sValue = GetSetting("TELEMETRY", "Enabled")
    m_Enabled = (sValue = "1")
    
    ' Read CollectorURL
    m_CollectorURL = GetSetting("TELEMETRY", "CollectorURL")
    
    ' Read FlushIntervalMs with default 30000
    sValue = GetSetting("TELEMETRY", "FlushIntervalMs")
    If Len(sValue) > 0 Then
        m_FlushIntervalMs = Val(sValue)
    Else
        m_FlushIntervalMs = 30000
    End If
    
    ' Read BatchSize with default 10
    sValue = GetSetting("TELEMETRY", "BatchSize")
    If Len(sValue) > 0 Then
        m_BatchSize = Val(sValue)
    Else
        m_BatchSize = 10
    End If
    
    Exit Sub
    
ReadTelemetrySettings_Err:
    Call RegistrarError(Err.Number, Err.Description, "ModTelemetry.ReadTelemetrySettings", Erl)
    Resume Next
End Sub

' ============================================================================
' Initialization / Shutdown
' ============================================================================

''
' Initializes the telemetry subsystem. Reads settings, populates resource
' attributes, seeds Rnd, calibrates QPC, and prepares the span buffer.
' Called once at client startup.
Public Sub Telemetry_Init()
    On Error GoTo Telemetry_Init_Err
    
    ' Step 1: Read settings from Configuracion.ini
    Call ReadTelemetrySettings
    
    ' Step 2: If CollectorURL is empty, disable and log warning
    If Len(m_CollectorURL) = 0 Then
        m_Enabled = False
        Call RegistrarError(-1, "Telemetry CollectorURL not configured, telemetry disabled", "ModTelemetry.Telemetry_Init")
    End If
    
    ' Always initialize login span handle to inactive
    ModAuth.m_LoginSpanHandle = -1
    
    ' Step 3: If not enabled, nothing more to do
    If Not m_Enabled Then Exit Sub
    
    ' Step 4: Set service name
    m_ServiceName = "argentum-client"
    
    ' Step 5: Set service version from App version info
    m_ServiceVersion = App.Major & "." & App.Minor & "." & App.Revision
    
    ' Step 6: Set machine identifier
    m_MachineId = Environ$("COMPUTERNAME")
    
    ' Step 7: Seed the random number generator
    Randomize Timer
    m_RndSeeded = True
    
    ' Step 8: Calibrate QPC-to-Unix-time mapping
    Call CalibrateTimestamp
    
    ' Step 9: Initialize buffer count
    m_BufferCount = 0
    
    ' Step 10: Enable the flush timer on frmMain
    Call frmMain.EnableTelemetryFlushTimer(m_FlushIntervalMs)
    
    Exit Sub
    
Telemetry_Init_Err:
    Call RegistrarError(Err.Number, Err.Description, "ModTelemetry.Telemetry_Init", Erl)
    Resume Next
End Sub

''
' Shuts down the telemetry subsystem. Flushes any remaining buffered spans
' and disables further telemetry collection.
' Called once at client exit.
Public Sub Telemetry_Shutdown()
    On Error GoTo Telemetry_Shutdown_Err
    
    ' Flush any remaining spans in the buffer
    Call Telemetry_FlushBuffer
    
    ' Disable telemetry
    m_Enabled = False
    
    Exit Sub
    
Telemetry_Shutdown_Err:
    Call RegistrarError(Err.Number, Err.Description, "ModTelemetry.Telemetry_Shutdown", Erl)
    Resume Next
End Sub

' ============================================================================
' Span Lifecycle
' ============================================================================

''
' Starts a new span by allocating a slot in m_ActiveSpans.
' Generates trace ID and span ID, records start time, attaches optional
' parent span ID and up to two initial attributes.
'
' @param    SpanName    The name of the span (e.g. "client.login")
' @param    ParentSpanId Optional parent span ID for trace correlation
' @param    Attr1Key    Optional first attribute key
' @param    Attr1Val    Optional first attribute value
' @param    Attr2Key    Optional second attribute key
' @param    Attr2Val    Optional second attribute value
' @return   Span handle (index into m_ActiveSpans), or -1 if disabled/no slot
Public Function Telemetry_StartSpan( _
    ByVal SpanName As String, _
    Optional ByVal ParentSpanId As String = "", _
    Optional ByVal Attr1Key As String = "", _
    Optional ByVal Attr1Val As String = "", _
    Optional ByVal Attr2Key As String = "", _
    Optional ByVal Attr2Val As String = "" _
) As Long
    On Error GoTo Telemetry_StartSpan_Err
    
    ' If telemetry is disabled, return null handle
    If Not m_Enabled Then
        Telemetry_StartSpan = -1
        Exit Function
    End If
    
    Dim SlotIndex As Long
    Dim FreeSlot As Long
    FreeSlot = -1
    
    ' Find first free slot in m_ActiveSpans
    For SlotIndex = 0 To MAX_ACTIVE_SPANS - 1
        If Not m_ActiveSpans(SlotIndex).Active Then
            FreeSlot = SlotIndex
            Exit For
        End If
    Next SlotIndex
    
    ' No free slot available
    If FreeSlot = -1 Then
        Telemetry_StartSpan = -1
        Exit Function
    End If
    
    ' Initialize the slot
    With m_ActiveSpans(FreeSlot)
        .TraceId = Telemetry_GenerateTraceId()
        .SpanId = Telemetry_GenerateSpanId()
        .ParentSpanId = ParentSpanId
        .Name = SpanName
        .Kind = OTEL_SPAN_KIND_CLIENT
        .StartTimeUnixNano = Telemetry_GetUnixNano()
        .EndTimeUnixNano = ""
        .StatusCode = OTEL_STATUS_UNSET
        .StatusMessage = ""
        .AttributeCount = 0
        .Active = True
        
        ' Attach optional first attribute
        If Len(Attr1Key) > 0 Then
            .Attributes(.AttributeCount).Key = Attr1Key
            .Attributes(.AttributeCount).Value = Attr1Val
            .AttributeCount = .AttributeCount + 1
        End If
        
        ' Attach optional second attribute
        If Len(Attr2Key) > 0 Then
            .Attributes(.AttributeCount).Key = Attr2Key
            .Attributes(.AttributeCount).Value = Attr2Val
            .AttributeCount = .AttributeCount + 1
        End If
    End With
    
    ' Return the slot index as the span handle
    Telemetry_StartSpan = FreeSlot
    Exit Function
    
Telemetry_StartSpan_Err:
    Call RegistrarError(Err.Number, Err.Description, "ModTelemetry.Telemetry_StartSpan", Erl)
    Telemetry_StartSpan = -1
    Resume Next
End Function

''
' Ends an active span by recording end time, setting status, copying to
' the buffer, and triggering a flush if the buffer threshold is met.
'
' @param    SpanHandle  The span handle returned by Telemetry_StartSpan
' @param    StatusCode  Optional status code (OTEL_STATUS_UNSET by default)
' @param    Attr1Key    Optional attribute key to add before ending
' @param    Attr1Val    Optional attribute value to add before ending
Public Sub Telemetry_EndSpan( _
    ByVal SpanHandle As Long, _
    Optional ByVal StatusCode As e_OtelStatus = OTEL_STATUS_UNSET, _
    Optional ByVal Attr1Key As String = "", _
    Optional ByVal Attr1Val As String = "" _
)
    On Error GoTo Telemetry_EndSpan_Err
    
    ' Validate span handle range
    If SpanHandle < 0 Or SpanHandle >= MAX_ACTIVE_SPANS Then Exit Sub
    
    ' Validate span is active
    If Not m_ActiveSpans(SpanHandle).Active Then Exit Sub
    
    With m_ActiveSpans(SpanHandle)
        ' Record end time
        .EndTimeUnixNano = Telemetry_GetUnixNano()
        
        ' Set status code
        .StatusCode = StatusCode
        
        ' Add optional attribute if provided and room available
        If Len(Attr1Key) > 0 Then
            If .AttributeCount < MAX_SPAN_ATTRIBUTES Then
                .Attributes(.AttributeCount).Key = Attr1Key
                .Attributes(.AttributeCount).Value = Attr1Val
                .AttributeCount = .AttributeCount + 1
            End If
        End If
        
        ' Mark as inactive
        .Active = False
    End With
    
    ' Copy completed span to buffer if there is room
    If m_BufferCount < MAX_BUFFER_SPANS Then
        m_Buffer(m_BufferCount) = m_ActiveSpans(SpanHandle)
        m_BufferCount = m_BufferCount + 1
    End If
    
    ' Check flush threshold
    If m_BufferCount >= m_BatchSize Or m_BufferCount >= MAX_BUFFER_SPANS Then
        Call Telemetry_FlushBuffer
    End If
    
    Exit Sub
    
Telemetry_EndSpan_Err:
    Call RegistrarError(Err.Number, Err.Description, "ModTelemetry.Telemetry_EndSpan", Erl)
    Resume Next
End Sub

''
' Adds a key/value attribute to an active span.
' Silently drops the attribute if the span is invalid, inactive,
' or the attribute limit has been reached.
'
' @param    SpanHandle  The span handle returned by Telemetry_StartSpan
' @param    Key         The attribute key
' @param    Value       The attribute value
Public Sub Telemetry_AddSpanAttribute( _
    ByVal SpanHandle As Long, _
    ByVal Key As String, _
    ByVal Value As String _
)
    On Error GoTo Telemetry_AddSpanAttribute_Err
    
    ' Validate span handle range
    If SpanHandle < 0 Or SpanHandle >= MAX_ACTIVE_SPANS Then Exit Sub
    
    ' Validate span is active
    If Not m_ActiveSpans(SpanHandle).Active Then Exit Sub
    
    ' Check attribute limit
    If m_ActiveSpans(SpanHandle).AttributeCount >= MAX_SPAN_ATTRIBUTES Then Exit Sub
    
    ' Add the attribute
    With m_ActiveSpans(SpanHandle)
        .Attributes(.AttributeCount).Key = Key
        .Attributes(.AttributeCount).Value = Value
        .AttributeCount = .AttributeCount + 1
    End With
    
    Exit Sub
    
Telemetry_AddSpanAttribute_Err:
    Call RegistrarError(Err.Number, Err.Description, "ModTelemetry.Telemetry_AddSpanAttribute", Erl)
    Resume Next
End Sub

' ============================================================================
' Serialization
' ============================================================================

''
' Escapes special JSON characters in a string.
'
' @param    Text    The string to escape
' @return   The escaped string safe for JSON inclusion
Private Function EscapeJsonString(ByRef Text As String) As String
    Dim Result As String
    
    Result = Text
    
    ' Backslash must be replaced first to avoid double-escaping
    Result = Replace$(Result, "\", "\\")
    Result = Replace$(Result, """", "\""")
    Result = Replace$(Result, vbCr, "\r")
    Result = Replace$(Result, vbLf, "\n")
    Result = Replace$(Result, vbTab, "\t")
    
    EscapeJsonString = Result
End Function

''
' Builds the full OTLP ExportTraceServiceRequest JSON string from an array
' of completed spans using cStringBuilder for performance.
'
' @param    Spans       Array of t_OtelSpan structures to serialize
' @param    SpanCount   Number of spans in the array to serialize
' @return   Valid OTLP JSON string
Public Function Telemetry_SerializeSpansToJson( _
    ByRef Spans() As t_OtelSpan, _
    ByVal SpanCount As Long _
) As String
    On Error GoTo Telemetry_SerializeSpansToJson_Err
    
    ' Early exit for empty span collection
    If SpanCount = 0 Then
        Telemetry_SerializeSpansToJson = "{""resourceSpans"":[]}"
        Exit Function
    End If
    
    Dim sb As New cStringBuilder
    Dim i As Long
    Dim j As Long
    
    ' Open resourceSpans and resource attributes
    sb.Append "{""resourceSpans"":[{""resource"":{""attributes"":["
    
    ' Resource attribute: service.name
    sb.Append "{""key"":""service.name"",""value"":{""stringValue"":"""
    sb.Append EscapeJsonString(m_ServiceName)
    sb.Append """}},"
    
    ' Resource attribute: service.version
    sb.Append "{""key"":""service.version"",""value"":{""stringValue"":"""
    sb.Append EscapeJsonString(m_ServiceVersion)
    sb.Append """}},"
    
    ' Resource attribute: host.id
    sb.Append "{""key"":""host.id"",""value"":{""stringValue"":"""
    sb.Append EscapeJsonString(m_MachineId)
    sb.Append """}}"
    
    ' Close resource attributes, open scopeSpans
    sb.Append "]},""scopeSpans"":[{""scope"":{""name"":""ao20-telemetry"",""version"":""1.0.0""},""spans"":["
    
    ' Serialize each span
    For i = 0 To SpanCount - 1
        ' Comma separator between spans
        If i > 0 Then sb.Append ","
        
        ' traceId, spanId, parentSpanId
        sb.Append "{""traceId"":"""
        sb.Append Spans(i).TraceId
        sb.Append """,""spanId"":"""
        sb.Append Spans(i).SpanId
        sb.Append """,""parentSpanId"":"""
        sb.Append Spans(i).ParentSpanId
        sb.Append ""","
        
        ' name, kind
        sb.Append """name"":"""
        sb.Append EscapeJsonString(Spans(i).Name)
        sb.Append """,""kind"":"
        sb.Append CStr(CLng(Spans(i).Kind))
        sb.Append ","
        
        ' startTimeUnixNano, endTimeUnixNano (as JSON strings)
        sb.Append """startTimeUnixNano"":"""
        sb.Append Spans(i).StartTimeUnixNano
        sb.Append """,""endTimeUnixNano"":"""
        sb.Append Spans(i).EndTimeUnixNano
        sb.Append ""","
        
        ' status
        sb.Append """status"":{""code"":"
        sb.Append CStr(CLng(Spans(i).StatusCode))
        sb.Append "},"
        
        ' attributes array
        sb.Append """attributes"":["
        
        For j = 0 To Spans(i).AttributeCount - 1
            ' Comma separator between attributes
            If j > 0 Then sb.Append ","
            
            sb.Append "{""key"":"""
            sb.Append EscapeJsonString(Spans(i).Attributes(j).Key)
            sb.Append """,""value"":{""stringValue"":"""
            sb.Append EscapeJsonString(Spans(i).Attributes(j).Value)
            sb.Append """}}"
        Next j
        
        ' Close attributes array and span object
        sb.Append "]}"
    Next i
    
    ' Close spans array, scopeSpans, and resourceSpans
    sb.Append "]}]}]}"
    
    Telemetry_SerializeSpansToJson = sb.toString
    Exit Function
    
Telemetry_SerializeSpansToJson_Err:
    Call RegistrarError(Err.Number, Err.Description, "ModTelemetry.Telemetry_SerializeSpansToJson", Erl)
    Telemetry_SerializeSpansToJson = "{""resourceSpans"":[]}"
    Resume Next
End Function

' ============================================================================
' Buffer Management
' ============================================================================

''
' Flushes all buffered spans by serializing and transmitting them.
Public Sub Telemetry_FlushBuffer()
    On Error GoTo Telemetry_FlushBuffer_Err
    
    ' Nothing to flush if disabled or buffer empty
    If Not m_Enabled Then Exit Sub
    If m_BufferCount = 0 Then Exit Sub
    
    ' Serialize all buffered spans to OTLP JSON
    Dim JsonPayload As String
    JsonPayload = Telemetry_SerializeSpansToJson(m_Buffer, m_BufferCount)
    
    ' Clear the buffer immediately (before sending, to avoid re-sending on error)
    m_BufferCount = 0
    
    ' POST to collector endpoint via frmConnect.InetTelemetry
    Dim TargetURL As String
    TargetURL = m_CollectorURL & "/v1/traces"
    
    ' Set headers and execute async POST
    frmConnect.InetTelemetry.RequestTimeout = 10
    Dim Headers As String
    Headers = "Content-Type: application/json" & vbCrLf
    frmConnect.InetTelemetry.Execute TargetURL, "POST", JsonPayload, Headers
    
    Exit Sub
    
Telemetry_FlushBuffer_Err:
    Call RegistrarError(Err.Number, Err.Description, "ModTelemetry.Telemetry_FlushBuffer", Erl)
    Resume Next
End Sub
