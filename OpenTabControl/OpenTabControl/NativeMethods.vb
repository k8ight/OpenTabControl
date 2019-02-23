#Region "   Imports"
Imports System.Drawing
Imports System.Reflection
Imports System.Runtime.InteropServices
Imports System.Security.Permissions
Imports System.Windows.Forms
#End Region

Friend NotInheritable Class NativeMethods

#Region "Windows Constants"

    <System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Performance", "CA1823:AvoidUnusedPrivateFields")> _
    Public Const WM_GETTABRECT As Integer = &H130A
    <System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Performance", "CA1823:AvoidUnusedPrivateFields")> _
    Public Const WS_EX_TRANSPARENT As Integer = &H20
    <System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Performance", "CA1823:AvoidUnusedPrivateFields")> _
    Public Const WM_SETFONT As Integer = &H30
    <System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Performance", "CA1823:AvoidUnusedPrivateFields")> _
    Public Const WM_FONTCHANGE As Integer = &H1D
    <System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Performance", "CA1823:AvoidUnusedPrivateFields")> _
    Public Const WM_HSCROLL As Integer = &H114
    <System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Performance", "CA1823:AvoidUnusedPrivateFields")> _
    Public Const TCM_HITTEST As Integer = &H130D
    <System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Performance", "CA1823:AvoidUnusedPrivateFields")> _
    Public Const WM_PAINT As Integer = &HF
    <System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Performance", "CA1823:AvoidUnusedPrivateFields")> _
    Public Const WS_EX_LAYOUTRTL As Integer = &H400000
    <System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Performance", "CA1823:AvoidUnusedPrivateFields")> _
    Public Const WS_EX_NOINHERITLAYOUT As Integer = &H100000

#End Region

#Region "Content Alignment"

    '<System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Performance", "CA1823:AvoidUnusedPrivateFields")> _
    Public Shared ReadOnly AnyRightAlign As ContentAlignment = ContentAlignment.BottomRight Or ContentAlignment.MiddleRight Or ContentAlignment.TopRight
    '<System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Performance", "CA1823:AvoidUnusedPrivateFields")> _
    Public Shared ReadOnly AnyLeftAlign As ContentAlignment = ContentAlignment.BottomLeft Or ContentAlignment.MiddleLeft Or ContentAlignment.TopLeft
    '<System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Performance", "CA1823:AvoidUnusedPrivateFields")> _
    Public Shared ReadOnly AnyTopAlign As ContentAlignment = ContentAlignment.TopRight Or ContentAlignment.TopCenter Or ContentAlignment.TopLeft
    '<System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Performance", "CA1823:AvoidUnusedPrivateFields")> _
    Public Shared ReadOnly AnyBottomAlign As ContentAlignment = ContentAlignment.BottomRight Or ContentAlignment.BottomCenter Or ContentAlignment.BottomLeft
    '<System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Performance", "CA1823:AvoidUnusedPrivateFields")> _
    Public Shared ReadOnly AnyMiddleAlign As ContentAlignment = ContentAlignment.MiddleRight Or ContentAlignment.MiddleCenter Or ContentAlignment.MiddleLeft
    '<System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Performance", "CA1823:AvoidUnusedPrivateFields")> _
    Public Shared ReadOnly AnyCenterAlign As ContentAlignment = ContentAlignment.BottomCenter Or ContentAlignment.MiddleCenter Or ContentAlignment.TopCenter

#End Region

#Region "User32.dll"

    Public Shared Function SendMessage(ByVal hWnd As IntPtr, ByVal msg As Integer, ByVal wParam As IntPtr, ByVal lParam As IntPtr) As IntPtr
        Dim _Control As Control = Control.FromHandle(hWnd)
        If _Control Is Nothing Then
            Return IntPtr.Zero
        End If
        Dim message As New Message()
        message.HWnd = hWnd
        message.LParam = lParam
        message.WParam = wParam
        message.Msg = msg
        Dim wproc As MethodInfo = _Control.[GetType]().GetMethod("WndProc", BindingFlags.NonPublic Or BindingFlags.InvokeMethod Or BindingFlags.FlattenHierarchy Or BindingFlags.IgnoreCase Or BindingFlags.Instance)
        Dim args As Object() = New Object() {message}
        wproc.Invoke(_Control, args)
        Return CType(args(0), Message).Result
    End Function

#End Region

#Region "Misc Functions"

    Public Shared Function LoWord(ByVal dWord As IntPtr) As Integer
        Return dWord.ToInt32() And &HFFFF
    End Function

    Public Shared Function HiWord(ByVal dWord As IntPtr) As Integer
        If (dWord.ToInt32() And &H80000000UI) = &H80000000UI Then
            Return (dWord.ToInt32() >> 16)
        Else
            Return (dWord.ToInt32() >> 16) And &HFFFF
        End If
    End Function

    <System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Security", "CA2106:SecureAsserts")> _
    <System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Security", "CA2122:DoNotIndirectlyExposeMethodsWithLinkDemands")> _
    Public Shared Function ToIntPtr(ByVal [structure] As Object) As IntPtr
        Dim lparam As IntPtr = IntPtr.Zero
        lparam = Marshal.AllocCoTaskMem(Marshal.SizeOf([structure]))
        Marshal.StructureToPtr([structure], lparam, False)
        Return lparam
    End Function


#End Region

#Region "Windows Structures and Enums"

    <Flags()> _
    Public Enum TCHITTESTFLAGS
        TCHT_NOWHERE = 1
        TCHT_ONITEMICON = 2
        TCHT_ONITEMLABEL = 4
        TCHT_ONITEM = TCHT_ONITEMICON Or TCHT_ONITEMLABEL
    End Enum



    <StructLayout(LayoutKind.Sequential)> _
    Public Structure TCHITTESTINFO

        Public Sub New(ByVal location As Point)
            pt = location
            flags = TCHITTESTFLAGS.TCHT_ONITEM
        End Sub

        Public pt As Point
        Public flags As TCHITTESTFLAGS
    End Structure

    <StructLayout(LayoutKind.Sequential, Pack:=4)> _
    Public Structure PAINTSTRUCT
        Public hdc As IntPtr
        Public fErase As Integer
        Public rcPaint As RECT
        Public fRestore As Integer
        Public fIncUpdate As Integer
        <MarshalAs(UnmanagedType.ByValArray, SizeConst:=32)> _
        Public rgbReserved As Byte()
    End Structure

    <StructLayout(LayoutKind.Sequential)> _
    Public Structure RECT
        Public left As Integer
        Public top As Integer
        Public right As Integer
        Public bottom As Integer

        Public Sub New(ByVal left As Integer, ByVal top As Integer, ByVal right As Integer, ByVal bottom As Integer)
            Me.left = left
            Me.top = top
            Me.right = right
            Me.bottom = bottom
        End Sub

        Public Sub New(ByVal r As Rectangle)
            Me.left = r.Left
            Me.top = r.Top
            Me.right = r.Right
            Me.bottom = r.Bottom
        End Sub

        Public Shared Function FromXYWH(ByVal x As Integer, ByVal y As Integer, ByVal width As Integer, ByVal height As Integer) As RECT
            Return New RECT(x, y, x + width, y + height)
        End Function

        Public Shared Function FromIntPtr(ByVal ptr As IntPtr) As RECT
            Dim rect As RECT = CType(Marshal.PtrToStructure(ptr, GetType(RECT)), RECT)
            Return rect
        End Function

        Public ReadOnly Property Size() As Size
            Get
                Return New Size(Me.right - Me.left, Me.bottom - Me.top)
            End Get
        End Property
    End Structure


#End Region

End Class

Friend NotInheritable Class ThemedColors

#Region "    Variables and Constants "

    Private Const NormalColor As String = "NormalColor"
    Private Const HomeStead As String = "HomeStead"
    Private Const Metallic As String = "Metallic"
    Private Const NoTheme As String = "NoTheme"

    Private Shared _toolBorder As Color()
#End Region

#Region "    Properties "

    <System.Diagnostics.CodeAnalysis.SuppressMessageAttribute("Microsoft.Performance", "CA1811:AvoidUncalledPrivateCode")> _
    Public Shared ReadOnly Property CurrentThemeIndex() As ColorScheme
        Get
            Return ThemedColors.GetCurrentThemeIndex()
        End Get
    End Property

    <System.Diagnostics.CodeAnalysis.SuppressMessageAttribute("Microsoft.Performance", "CA1811:AvoidUncalledPrivateCode")> _
    Public Shared ReadOnly Property ToolBorder() As Color
        Get
            Return ThemedColors._toolBorder(CInt(ThemedColors.CurrentThemeIndex))
        End Get
    End Property

#End Region

#Region "    Constructors "

    <System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Performance", "CA1810:InitializeReferenceTypeStaticFieldsInline")> _
    Shared Sub New()
        ThemedColors._toolBorder = New Color() {Color.FromArgb(127, 157, 185), Color.FromArgb(164, 185, 127), Color.FromArgb(165, 172, 178), Color.FromArgb(132, 130, 132)}
    End Sub

    Private Sub New()
    End Sub

#End Region

    Private Shared Function GetCurrentThemeIndex() As ColorScheme
        Dim theme As ColorScheme = ColorScheme.NoTheme
        If VisualStyles.VisualStyleInformation.IsSupportedByOS AndAlso VisualStyles.VisualStyleInformation.IsEnabledByUser AndAlso Application.RenderWithVisualStyles Then
            Select Case VisualStyles.VisualStyleInformation.ColorScheme
                Case NormalColor
                    theme = ColorScheme.NormalColor
                    Exit Select
                Case HomeStead
                    theme = ColorScheme.HomeStead
                    Exit Select
                Case Metallic
                    theme = ColorScheme.Metallic
                    Exit Select
                Case Else
                    theme = ColorScheme.NoTheme
                    Exit Select
            End Select
        End If

        Return theme
    End Function

    Public Enum ColorScheme
        NormalColor = 0
        HomeStead = 1
        Metallic = 2
        NoTheme = 3
    End Enum

End Class

