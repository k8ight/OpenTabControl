#Region "   Imports"

Imports System.Drawing
Imports System.Drawing.Drawing2D
Imports System.Windows.Forms
Imports System.ComponentModel

#End Region

<System.ComponentModel.ToolboxItem(False)> _
Public Class OpenStyleOfficerovider
    Inherits TabStyleRoundedProvider
   
    Public Sub New(ByVal tabControl As OpenTabControl)
        MyBase.New(tabControl)
        Me._Radius = 3
        Me._ShowTabCloser = True
        Me._CloserColorActive = Color.Black
        Me._CloserColor = Color.FromArgb(117, 99, 61)
        Me._TextColor = Color.White
        Me._TextColorDisabled = Color.WhiteSmoke
        Me._BorderColor = Color.Transparent
        Me._BorderColorHot = Color.FromArgb(155, 167, 183)
        Me.Padding = New Point(6, 5)
    End Sub

    Protected Overrides Function GetTabBackgroundBrush(ByVal index As Integer) As Brush
        Dim fillBrush As LinearGradientBrush = Nothing
        Dim dark As Color = Color.Transparent
        Dim light As Color = Color.Transparent
        If Me._TabControl.SelectedIndex = index Then
            dark = Color.FromArgb(255, 214, 78)
            light = SystemColors.Window
        ElseIf Not Me._TabControl.TabPages(index).Enabled Then
            light = dark
        ElseIf Me.HotTrack AndAlso index = Me._TabControl.ActiveIndex Then
            dark = Color.FromArgb(108, 116, 118)
            light = dark
        End If
        Dim tabBounds As Rectangle = Me.GetTabRect(index)
        tabBounds.Inflate(3, 3)
        tabBounds.X -= 1
        tabBounds.Y -= 1
        Select Case Me._TabControl.Alignment
            Case TabAlignment.Top
                fillBrush = New LinearGradientBrush(tabBounds, light, dark, LinearGradientMode.Vertical)
                Exit Select
            Case TabAlignment.Bottom
                fillBrush = New LinearGradientBrush(tabBounds, dark, light, LinearGradientMode.Vertical)
                Exit Select
            Case TabAlignment.Left
                fillBrush = New LinearGradientBrush(tabBounds, light, dark, LinearGradientMode.Horizontal)
                Exit Select
            Case TabAlignment.Right
                fillBrush = New LinearGradientBrush(tabBounds, dark, light, LinearGradientMode.Horizontal)
                Exit Select
        End Select
        fillBrush.Blend = GetBackgroundBlend()
        Return fillBrush
    End Function

    Private Overloads Shared Function GetBackgroundBlend() As Blend
        Dim relativeIntensities As Single() = New Single() {0.0F, 0.5F, 1.0F, 1.0F}
        Dim relativePositions As Single() = New Single() {0.0F, 0.5F, 0.51F, 1.0F}
        Dim blend As New Blend()
        blend.Factors = relativeIntensities
        blend.Positions = relativePositions
        Return blend
    End Function

    Public Overrides Function GetPageBackgroundBrush(ByVal index As Integer) As Brush
        Dim light As Color = Color.Transparent
        If Me._TabControl.SelectedIndex = index Then
            light = Color.FromArgb(255, 214, 78)
        ElseIf Not Me._TabControl.TabPages(index).Enabled Then
            light = Color.Transparent
        ElseIf Me._HotTrack AndAlso index = Me._TabControl.ActiveIndex Then
            light = Color.Transparent
        End If
        Return New SolidBrush(light)
    End Function

    Protected Overrides Sub DrawTabCloseButton(ByVal index As Integer, ByVal graphics As Graphics)
        If Me._ShowTabCloser Then
            Dim closerRect As Rectangle = Me._TabControl.GetTabCloserRect(index)
            graphics.SmoothingMode = SmoothingMode.AntiAlias
            If closerRect.Contains(Me._TabControl.MousePosition) Then
                Using closerPath As GraphicsPath = GetCloserButtonPath(closerRect)
                    graphics.FillPath(Brushes.White, closerPath)
                    Using closerPen As New Pen(Color.FromArgb(255, 214, 78))
                        graphics.DrawPath(closerPen, closerPath)
                    End Using
                End Using
                Using closerPath As GraphicsPath = GetCloserPath(closerRect)
                    Using closerPen As New Pen(Me._CloserColorActive)
                        closerPen.Width = 2
                        graphics.DrawPath(closerPen, closerPath)
                    End Using
                End Using
            Else
                If index = Me._TabControl.SelectedIndex Then
                    Using closerPath As GraphicsPath = GetCloserPath(closerRect)
                        Using closerPen As New Pen(Me._CloserColor)
                            closerPen.Width = 2
                            graphics.DrawPath(closerPen, closerPath)
                        End Using
                    End Using
                ElseIf index = Me._TabControl.ActiveIndex Then
                    Using closerPath As GraphicsPath = GetCloserPath(closerRect)
                        Using closerPen As New Pen(Color.FromArgb(155, 167, 183))
                            closerPen.Width = 2
                            graphics.DrawPath(closerPen, closerPath)
                        End Using
                    End Using
                End If

            End If
        End If
    End Sub

    Private Shared Function GetCloserButtonPath(ByVal closerRect As Rectangle) As GraphicsPath
        Dim closerPath As New GraphicsPath()
        closerPath.AddLine(closerRect.X - 1, closerRect.Y - 2, closerRect.Right + 1, closerRect.Y - 2)
        closerPath.AddLine(closerRect.Right + 2, closerRect.Y - 1, closerRect.Right + 2, closerRect.Bottom + 1)
        closerPath.AddLine(closerRect.Right + 1, closerRect.Bottom + 2, closerRect.X - 1, closerRect.Bottom + 2)
        closerPath.AddLine(closerRect.X - 2, closerRect.Bottom + 1, closerRect.X - 2, closerRect.Y - 1)
        closerPath.CloseFigure()
        Return closerPath
    End Function

End Class

<System.ComponentModel.ToolboxItem(False)> _
Public Class TabStyleRoundedProvider
    Inherits OpenStyleProvider
    Public Sub New(ByVal tabControl As OpenTabControl)
        MyBase.New(tabControl)
        Me._Radius = 10
        Me.Padding = New Point(6, 3)
    End Sub

    Public Overrides Sub AddTabBorder(ByVal path As System.Drawing.Drawing2D.GraphicsPath, ByVal tabBounds As System.Drawing.Rectangle)
        Select Case Me._TabControl.Alignment
            Case TabAlignment.Top
                path.AddLine(tabBounds.X, tabBounds.Bottom, tabBounds.X, tabBounds.Y + Me._Radius)
                path.AddArc(tabBounds.X, tabBounds.Y, Me._Radius * 2, Me._Radius * 2, 180, 90)
                path.AddLine(tabBounds.X + Me._Radius, tabBounds.Y, tabBounds.Right - Me._Radius, tabBounds.Y)
                path.AddArc(tabBounds.Right - Me._Radius * 2, tabBounds.Y, Me._Radius * 2, Me._Radius * 2, 270, 90)
                path.AddLine(tabBounds.Right, tabBounds.Y + Me._Radius, tabBounds.Right, tabBounds.Bottom)
                Exit Select
            Case TabAlignment.Bottom
                path.AddLine(tabBounds.Right, tabBounds.Y, tabBounds.Right, tabBounds.Bottom - Me._Radius)
                path.AddArc(tabBounds.Right - Me._Radius * 2, tabBounds.Bottom - Me._Radius * 2, Me._Radius * 2, Me._Radius * 2, 0, 90)
                path.AddLine(tabBounds.Right - Me._Radius, tabBounds.Bottom, tabBounds.X + Me._Radius, tabBounds.Bottom)
                path.AddArc(tabBounds.X, tabBounds.Bottom - Me._Radius * 2, Me._Radius * 2, Me._Radius * 2, 90, 90)
                path.AddLine(tabBounds.X, tabBounds.Bottom - Me._Radius, tabBounds.X, tabBounds.Y)
                Exit Select
            Case TabAlignment.Left
                path.AddLine(tabBounds.Right, tabBounds.Bottom, tabBounds.X + Me._Radius, tabBounds.Bottom)
                path.AddArc(tabBounds.X, tabBounds.Bottom - Me._Radius * 2, Me._Radius * 2, Me._Radius * 2, 90, 90)
                path.AddLine(tabBounds.X, tabBounds.Bottom - Me._Radius, tabBounds.X, tabBounds.Y + Me._Radius)
                path.AddArc(tabBounds.X, tabBounds.Y, Me._Radius * 2, Me._Radius * 2, 180, 90)
                path.AddLine(tabBounds.X + Me._Radius, tabBounds.Y, tabBounds.Right, tabBounds.Y)
                Exit Select
            Case TabAlignment.Right
                path.AddLine(tabBounds.X, tabBounds.Y, tabBounds.Right - Me._Radius, tabBounds.Y)
                path.AddArc(tabBounds.Right - Me._Radius * 2, tabBounds.Y, Me._Radius * 2, Me._Radius * 2, 270, 90)
                path.AddLine(tabBounds.Right, tabBounds.Y + Me._Radius, tabBounds.Right, tabBounds.Bottom - Me._Radius)
                path.AddArc(tabBounds.Right - Me._Radius * 2, tabBounds.Bottom - Me._Radius * 2, Me._Radius * 2, Me._Radius * 2, 0, 90)
                path.AddLine(tabBounds.Right - Me._Radius, tabBounds.Bottom, tabBounds.X, tabBounds.Bottom)
                Exit Select
        End Select
    End Sub
End Class

<ToolboxItemAttribute(False)> _
Public MustInherit Class OpenStyleProvider
    Inherits Component
#Region "Constructor"
   
    Protected Sub New(ByVal tabControl As OpenTabControl)
        Me._TabControl = tabControl
        Me._BorderColor = Color.Empty
        Me._BorderColorSelected = Color.Empty
        Me._FocusColor = Color.Orange
        If Me._TabControl.RightToLeftLayout Then
            Me._ImageAlign = ContentAlignment.MiddleRight
        Else
            Me._ImageAlign = ContentAlignment.MiddleLeft
        End If
        Me.HotTrack = True
        Me.Padding = New Point(6, 3)
    End Sub

#End Region

#Region "Factory Methods"
   
    Public Shared Function CreateProvider(ByVal tabControl As OpenTabControl) As OpenStyleProvider
        Dim provider As OpenStyleProvider
        provider = New OpenStyleOfficerovider(tabControl)
        provider._Style = tabControl.TabStyle
        Return provider
    End Function

#End Region

#Region "Protected variables"

    <System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Design", "CA1051:DoNotDeclareVisibleInstanceFields")> _
    Protected _TabControl As OpenTabControl
    <System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Design", "CA1051:DoNotDeclareVisibleInstanceFields")> _
    Protected _Padding As Point
    <System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Design", "CA1051:DoNotDeclareVisibleInstanceFields")> _
    Protected _HotTrack As Boolean
    <System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Design", "CA1051:DoNotDeclareVisibleInstanceFields")> _
    Protected _Style As OpenTabStyle = OpenTabStyle.Normal
    <System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Design", "CA1051:DoNotDeclareVisibleInstanceFields")> _
    Protected _ImageAlign As ContentAlignment
    <System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Design", "CA1051:DoNotDeclareVisibleInstanceFields")> _
    Protected _Radius As Integer = 1
    <System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Design", "CA1051:DoNotDeclareVisibleInstanceFields")> _
    Protected _Overlap As Integer
    <System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Design", "CA1051:DoNotDeclareVisibleInstanceFields")> _
    Protected _FocusTrack As Boolean
    <System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Design", "CA1051:DoNotDeclareVisibleInstanceFields")> _
    Protected _Opacity As Single = 1
    <System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Design", "CA1051:DoNotDeclareVisibleInstanceFields")> _
    Protected _ShowTabCloser As Boolean
    <System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Design", "CA1051:DoNotDeclareVisibleInstanceFields")> _
    Protected _BorderColorSelected As Color = Color.Empty
    <System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Design", "CA1051:DoNotDeclareVisibleInstanceFields")> _
    Protected _BorderColor As Color = Color.Empty
    <System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Design", "CA1051:DoNotDeclareVisibleInstanceFields")> _
    Protected _BorderColorHot As Color = Color.Empty
    <System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Design", "CA1051:DoNotDeclareVisibleInstanceFields")> _
    Protected _CloserColorActive As Color = Color.Black
    <System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Design", "CA1051:DoNotDeclareVisibleInstanceFields")> _
    Protected _CloserColor As Color = Color.DarkGray
    <System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Design", "CA1051:DoNotDeclareVisibleInstanceFields")> _
    Protected _FocusColor As Color = Color.Empty
    <System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Design", "CA1051:DoNotDeclareVisibleInstanceFields")> _
    Protected _TextColor As Color = Color.Empty
    <System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Design", "CA1051:DoNotDeclareVisibleInstanceFields")> _
    Protected _TextColorSelected As Color = Color.Empty
    <System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Design", "CA1051:DoNotDeclareVisibleInstanceFields")> _
    Protected _TextColorDisabled As Color = Color.Empty

#End Region

#Region "overridable Methods"

    Public MustOverride Sub AddTabBorder(ByVal path As GraphicsPath, ByVal tabBounds As Rectangle)

    Public Overridable Function GetTabRect(ByVal index As Integer) As Rectangle
        If index < 0 Then
            Return New Rectangle()
        End If
        Dim tabBounds As Rectangle = Me._TabControl.GetTabRect(index)
        If Me._TabControl.RightToLeftLayout Then
            tabBounds.X = Me._TabControl.Width - tabBounds.Right
        End If
        Dim firstTabinRow As Boolean = Me._TabControl.IsFirstTabInRow(index)
        Select Case Me._TabControl.Alignment
            Case TabAlignment.Top
                tabBounds.Height += 2
                Exit Select
            Case TabAlignment.Bottom
                tabBounds.Height += 2
                tabBounds.Y -= 2
                Exit Select
            Case TabAlignment.Left
                tabBounds.Width += 2
                Exit Select
            Case TabAlignment.Right
                tabBounds.X -= 2
                tabBounds.Width += 2
                Exit Select
        End Select
        If (Not firstTabinRow OrElse Me._TabControl.RightToLeftLayout) AndAlso Me._Overlap > 0 Then
            If Me._TabControl.Alignment <= TabAlignment.Bottom Then
                tabBounds.X -= Me._Overlap
                tabBounds.Width += Me._Overlap
            Else
                tabBounds.Y -= Me._Overlap
                tabBounds.Height += Me._Overlap
            End If
        End If
        Me.EnsureFirstTabIsInView(tabBounds, index)
        Return tabBounds
    End Function

    <System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Design", "CA1045:DoNotPassTypesByReference", MessageId:="0#")> _
    Protected Overridable Sub EnsureFirstTabIsInView(ByRef tabBounds As Rectangle, ByVal index As Integer)
        Dim firstTabinRow As Boolean = Me._TabControl.IsFirstTabInRow(index)
        If firstTabinRow Then
            If Me._TabControl.Alignment <= TabAlignment.Bottom Then
                If Me._TabControl.RightToLeftLayout Then
                    If tabBounds.Left < Me._TabControl.Right Then
                        Dim tabPageRight As Integer = Me._TabControl.GetPageBounds(index).Right
                        If tabBounds.Right > tabPageRight Then
                            tabBounds.Width -= (tabBounds.Right - tabPageRight)
                        End If
                    End If
                Else
                    If tabBounds.Right > 0 Then
                        Dim tabPageX As Integer = Me._TabControl.GetPageBounds(index).X
                        If tabBounds.X < tabPageX Then
                            tabBounds.Width -= (tabPageX - tabBounds.X)
                            tabBounds.X = tabPageX
                        End If
                    End If
                End If
            Else
                If Me._TabControl.RightToLeftLayout Then
                    If tabBounds.Top < Me._TabControl.Bottom Then
                        Dim tabPageBottom As Integer = Me._TabControl.GetPageBounds(index).Bottom
                        If tabBounds.Bottom > tabPageBottom Then
                            tabBounds.Height -= (tabBounds.Bottom - tabPageBottom)
                        End If
                    End If
                Else
                    If tabBounds.Bottom > 0 Then
                        Dim tabPageY As Integer = Me._TabControl.GetPageBounds(index).Location.Y
                        If tabBounds.Y < tabPageY Then
                            tabBounds.Height -= (tabPageY - tabBounds.Y)
                            tabBounds.Y = tabPageY
                        End If
                    End If
                End If
            End If
        End If
    End Sub

    Protected Overridable Function GetTabBackgroundBrush(ByVal index As Integer) As Brush
        Dim fillBrush As LinearGradientBrush = Nothing
        Dim dark As Color = Color.FromArgb(207, 207, 207)
        Dim light As Color = Color.FromArgb(242, 242, 242)

        If Me._TabControl.SelectedIndex = index Then
            dark = SystemColors.ControlLight
            light = SystemColors.Window
        ElseIf Not Me._TabControl.TabPages(index).Enabled Then
            light = dark
        ElseIf Me._HotTrack AndAlso index = Me._TabControl.ActiveIndex Then
            light = Color.FromArgb(234, 246, 253)
            dark = Color.FromArgb(167, 217, 245)
        End If
        Dim tabBounds As Rectangle = Me.GetTabRect(index)
        tabBounds.Inflate(3, 3)
        tabBounds.X -= 1
        tabBounds.Y -= 1
        Select Case Me._TabControl.Alignment
            Case TabAlignment.Top
                If Me._TabControl.SelectedIndex = index Then
                    dark = light
                End If
                fillBrush = New LinearGradientBrush(tabBounds, light, dark, LinearGradientMode.Vertical)
                Exit Select
            Case TabAlignment.Bottom
                fillBrush = New LinearGradientBrush(tabBounds, light, dark, LinearGradientMode.Vertical)
                Exit Select
            Case TabAlignment.Left
                fillBrush = New LinearGradientBrush(tabBounds, dark, light, LinearGradientMode.Horizontal)
                Exit Select
            Case TabAlignment.Right
                fillBrush = New LinearGradientBrush(tabBounds, light, dark, LinearGradientMode.Horizontal)
                Exit Select
        End Select
        fillBrush.Blend = Me.GetBackgroundBlend()
        Return fillBrush
    End Function

#End Region

#Region "Base Properties"

    <Browsable(False), DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)> _
    Public Property DisplayStyle() As OpenTabStyle
        Get
            Return Me._Style
        End Get
        Set(ByVal value As OpenTabStyle)
            Me._Style = value
        End Set
    End Property

    <Category("Appearance")> _
    Public Property ImageAlign() As ContentAlignment
        Get
            Return Me._ImageAlign
        End Get
        Set(ByVal value As ContentAlignment)
            Me._ImageAlign = value
            Me._TabControl.Invalidate()
        End Set
    End Property

    <Category("Appearance")> _
    Public Property Padding() As Point
        Get
            Return Me._Padding
        End Get
        Set(ByVal value As Point)
            Me._Padding = value
            If Me._ShowTabCloser Then
                If value.X + CInt(Me._Radius \ 2) < -6 Then
                    DirectCast(Me._TabControl, TabControl).Padding = New Point(0, value.Y)
                Else
                    DirectCast(Me._TabControl, TabControl).Padding = New Point(value.X + CInt(Me._Radius \ 2) + 6, value.Y)
                End If
            Else
                If value.X + CInt(Me._Radius \ 2) < 1 Then
                    DirectCast(Me._TabControl, TabControl).Padding = New Point(0, value.Y)
                Else
                    DirectCast(Me._TabControl, TabControl).Padding = New Point(value.X + CInt(Me._Radius \ 2) - 1, value.Y)
                End If
            End If
        End Set
    End Property

    <Category("Appearance"), DefaultValue(1), Browsable(True)> _
    Public Property Radius() As Integer
        Get
            Return Me._Radius
        End Get
        Set(ByVal value As Integer)
            If value < 1 Then
                Throw New ArgumentException("The radius must be greater than 1", "value")
            End If
            Me._Radius = value
            Me.Padding = Me._Padding
        End Set
    End Property

    <Category("Appearance")> _
    Public Property Overlap() As Integer
        Get
            Return Me._Overlap
        End Get
        Set(ByVal value As Integer)
            If value < 0 Then
                Throw New ArgumentException("The tabs cannot have a negative overlap", "value")
            End If

            Me._Overlap = value
        End Set
    End Property

    <Category("Appearance")> _
    Public Property FocusTrack() As Boolean
        Get
            Return Me._FocusTrack
        End Get
        Set(ByVal value As Boolean)
            Me._FocusTrack = value
            Me._TabControl.Invalidate()
        End Set
    End Property

    <Category("Appearance")> _
    Public Property HotTrack() As Boolean
        Get
            Return Me._HotTrack
        End Get
        Set(ByVal value As Boolean)
            Me._HotTrack = value
            DirectCast(Me._TabControl, TabControl).HotTrack = value
        End Set
    End Property

    <Category("Appearance")> _
    Public Property TabCloseButton() As Boolean
        Get
            Return Me._ShowTabCloser
        End Get
        Set(ByVal value As Boolean)
            Me._ShowTabCloser = value
            Me.Padding = Me._Padding
        End Set
    End Property

    <Category("Appearance")> _
    Public Property Opacity() As Single
        Get
            Return Me._Opacity
        End Get
        Set(ByVal value As Single)
            If value < 0 Then
                Throw New ArgumentException("The opacity must be between 0 and 1", "value")
            End If
            If value > 1 Then
                Throw New ArgumentException("The opacity must be between 0 and 1", "value")
            End If
            Me._Opacity = value
            Me._TabControl.Invalidate()
        End Set
    End Property

    <Category("Appearance"), DefaultValue(GetType(Color), "")> _
    Public Property BorderColorSelected() As Color
        Get
            If Me._BorderColorSelected.IsEmpty Then
                Return ThemedColors.ToolBorder
            Else
                Return Me._BorderColorSelected
            End If
        End Get
        Set(ByVal value As Color)
            If value.Equals(ThemedColors.ToolBorder) Then
                Me._BorderColorSelected = Color.Empty
            Else
                Me._BorderColorSelected = value
            End If
            Me._TabControl.Invalidate()
        End Set
    End Property

    <Category("Appearance"), DefaultValue(GetType(Color), "")> _
    Public Property BorderColorHot() As Color
        Get
            If Me._BorderColorHot.IsEmpty Then
                Return SystemColors.ControlDark
            Else
                Return Me._BorderColorHot
            End If
        End Get
        Set(ByVal value As Color)
            If value.Equals(SystemColors.ControlDark) Then
                Me._BorderColorHot = Color.Empty
            Else
                Me._BorderColorHot = value
            End If
            Me._TabControl.Invalidate()
        End Set
    End Property

    <Category("Appearance"), DefaultValue(GetType(Color), "")> _
    Public Property BorderColor() As Color
        Get
            If Me._BorderColor.IsEmpty Then
                Return SystemColors.ControlDark
            Else
                Return Me._BorderColor
            End If
        End Get
        Set(ByVal value As Color)
            If value.Equals(SystemColors.ControlDark) Then
                Me._BorderColor = Color.Empty
            Else
                Me._BorderColor = value
            End If
            Me._TabControl.Invalidate()
        End Set
    End Property

    <Category("Appearance"), DefaultValue(GetType(Color), "Black")> _
    Public Property TextColor() As Color
        Get
            If Me._TextColor.IsEmpty Then
                Return SystemColors.ControlText
            Else
                Return Me._TextColor
            End If
        End Get
        Set(ByVal value As Color)
            If value.Equals(SystemColors.ControlText) Then
                Me._TextColor = Color.Empty
            Else
                Me._TextColor = value
            End If
            Me._TabControl.Invalidate()
        End Set
    End Property

    <Category("Appearance"), DefaultValue(GetType(Color), "")> _
    Public Property TextColorSelected() As Color
        Get
            If Me._TextColorSelected.IsEmpty Then
                Return SystemColors.ControlText
            Else
                Return Me._TextColorSelected
            End If
        End Get
        Set(ByVal value As Color)
            If value.Equals(SystemColors.ControlText) Then
                Me._TextColorSelected = Color.Empty
            Else
                Me._TextColorSelected = value
            End If
            Me._TabControl.Invalidate()
        End Set
    End Property

    <Category("Appearance"), DefaultValue(GetType(Color), "")> _
    Public Property TextColorDisabled() As Color
        Get
            If Me._TextColor.IsEmpty Then
                Return SystemColors.ControlDark
            Else
                Return Me._TextColorDisabled
            End If
        End Get
        Set(ByVal value As Color)
            If value.Equals(SystemColors.ControlDark) Then
                Me._TextColorDisabled = Color.Empty
            Else
                Me._TextColorDisabled = value
            End If
            Me._TabControl.Invalidate()
        End Set
    End Property

    <Category("Appearance"), DefaultValue(GetType(Color), "Orange")> _
    Public Property FocusColor() As Color
        Get
            Return Me._FocusColor
        End Get
        Set(ByVal value As Color)
            Me._FocusColor = value
            Me._TabControl.Invalidate()
        End Set
    End Property

    Public Property CloserColorActive() As Color
        Get
            Return Me._CloserColorActive
        End Get
        Set(ByVal value As Color)
            Me._CloserColorActive = value
            Me._TabControl.Invalidate()
        End Set
    End Property

    <Category("Appearance"), DefaultValue(GetType(Color), "DarkGrey")> _
    Public Property CloserColor() As Color
        Get
            Return Me._CloserColor
        End Get
        Set(ByVal value As Color)
            Me._CloserColor = value
            Me._TabControl.Invalidate()
        End Set
    End Property

#End Region

#Region "Painting"

    Public Sub PaintTab(ByVal index As Integer, ByVal graphics As Graphics)
        Using tabpath As GraphicsPath = Me.GetTabBorder(index)
            Using fillBrush As Brush = Me.GetTabBackgroundBrush(index)
                graphics.FillPath(fillBrush, tabpath)
                If Me._TabControl.Focused Then
                    Me.DrawTabFocusIndicator(tabpath, index, graphics)
                End If
                Me.DrawTabCloseButton(index, graphics)
            End Using
        End Using
    End Sub

    Protected Overridable Sub DrawTabCloseButton(ByVal index As Integer, ByVal graphics As Graphics)
        If Me._ShowTabCloser Then
            Dim closerRect As Rectangle = Me._TabControl.GetTabCloserRect(index)
            graphics.SmoothingMode = SmoothingMode.AntiAlias
            Using closerPath As GraphicsPath = OpenStyleProvider.GetCloserPath(closerRect)
                If closerRect.Contains(Me._TabControl.MousePosition) Then
                    Using closerPen As New Pen(Me._CloserColorActive)
                        graphics.DrawPath(closerPen, closerPath)
                    End Using
                Else
                    Using closerPen As New Pen(Me._CloserColor)
                        graphics.DrawPath(closerPen, closerPath)
                    End Using

                End If
            End Using
        End If
    End Sub

    Protected Shared Function GetCloserPath(ByVal closerRect As Rectangle) As GraphicsPath
        Dim closerPath As New GraphicsPath()
        closerPath.AddLine(closerRect.X, closerRect.Y, closerRect.Right, closerRect.Bottom)
        closerPath.CloseFigure()
        closerPath.AddLine(closerRect.Right, closerRect.Y, closerRect.X, closerRect.Bottom)
        closerPath.CloseFigure()

        Return closerPath
    End Function

    Private Sub DrawTabFocusIndicator(ByVal tabpath As GraphicsPath, ByVal index As Integer, ByVal graphics As Graphics)
        If Me._FocusTrack AndAlso Me._TabControl.Focused AndAlso index = Me._TabControl.SelectedIndex Then
            Dim focusBrush As Brush = Nothing
            Dim pathRect As RectangleF = tabpath.GetBounds()
            Dim focusRect As Rectangle = Rectangle.Empty
            Select Case Me._TabControl.Alignment
                Case TabAlignment.Top
                    focusRect = New Rectangle(CInt(Math.Truncate(pathRect.X)), CInt(Math.Truncate(pathRect.Y)), CInt(Math.Truncate(pathRect.Width)), 4)
                    focusBrush = New LinearGradientBrush(focusRect, Me._FocusColor, SystemColors.Window, LinearGradientMode.Vertical)
                    Exit Select
                Case TabAlignment.Bottom
                    focusRect = New Rectangle(CInt(Math.Truncate(pathRect.X)), CInt(Math.Truncate(pathRect.Bottom)) - 4, CInt(Math.Truncate(pathRect.Width)), 4)
                    focusBrush = New LinearGradientBrush(focusRect, SystemColors.ControlLight, Me._FocusColor, LinearGradientMode.Vertical)
                    Exit Select
                Case TabAlignment.Left
                    focusRect = New Rectangle(CInt(Math.Truncate(pathRect.X)), CInt(Math.Truncate(pathRect.Y)), 4, CInt(Math.Truncate(pathRect.Height)))
                    focusBrush = New LinearGradientBrush(focusRect, Me._FocusColor, SystemColors.ControlLight, LinearGradientMode.Horizontal)
                    Exit Select
                Case TabAlignment.Right
                    focusRect = New Rectangle(CInt(Math.Truncate(pathRect.Right)) - 4, CInt(Math.Truncate(pathRect.Y)), 4, CInt(Math.Truncate(pathRect.Height)))
                    focusBrush = New LinearGradientBrush(focusRect, SystemColors.ControlLight, Me._FocusColor, LinearGradientMode.Horizontal)
                    Exit Select
            End Select

            '	Ensure the focus stip does not go outside the tab
            Dim focusRegion As New Region(focusRect)
            focusRegion.Intersect(tabpath)
            graphics.FillRegion(focusBrush, focusRegion)
            focusRegion.Dispose()
            focusBrush.Dispose()
        End If
    End Sub

#End Region

#Region "Background brushes"

    Private Function GetBackgroundBlend() As Blend
        Dim relativeIntensities As Single() = New Single() {0.0F, 0.7F, 1.0F}
        Dim relativePositions As Single() = New Single() {0.0F, 0.6F, 1.0F}

        '	Glass look to top aligned tabs
        If Me._TabControl.Alignment = TabAlignment.Top Then
            relativeIntensities = New Single() {0.0F, 0.5F, 1.0F, 1.0F}
            relativePositions = New Single() {0.0F, 0.5F, 0.51F, 1.0F}
        End If

        Dim blend As New Blend()
        blend.Factors = relativeIntensities
        blend.Positions = relativePositions

        Return blend
    End Function

    Public Overridable Function GetPageBackgroundBrush(ByVal index As Integer) As Brush

        '	Capture the colours dependant on selection state of the tab
        Dim light As Color = Color.FromArgb(242, 242, 242)
        If Me._TabControl.Alignment = TabAlignment.Top Then
            light = Color.FromArgb(207, 207, 207)
        End If

        If Me._TabControl.SelectedIndex = index Then
            light = SystemColors.Window
        ElseIf Not Me._TabControl.TabPages(index).Enabled Then
            light = Color.FromArgb(207, 207, 207)
        ElseIf Me._HotTrack AndAlso index = Me._TabControl.ActiveIndex Then
            '	Enable hot tracking
            light = Color.FromArgb(234, 246, 253)
        End If

        Return New SolidBrush(light)
    End Function

#End Region

#Region "Tab border and rect"

    Public Function GetTabBorder(ByVal index As Integer) As GraphicsPath

        Dim path As New GraphicsPath()
        Dim tabBounds As Rectangle = Me.GetTabRect(index)

        Me.AddTabBorder(path, tabBounds)

        path.CloseFigure()
        Return path
    End Function

#End Region

End Class