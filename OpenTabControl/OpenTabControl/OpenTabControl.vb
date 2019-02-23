

#Region "   Imports"

Imports System
Imports System.Collections.Generic
Imports System.ComponentModel
Imports System.Drawing
Imports System.Drawing.Drawing2D
Imports System.Drawing.Imaging
Imports System.Security.Permissions
Imports System.Runtime.InteropServices
Imports System.Windows.Forms

#End Region

#Region "   OpenTabControl Class"
''' <summary>
'''OpenTabControl Class
''' </summary>
''' <remarks>Version 1.0</remarks>
Public Class OpenTabControl
    Inherits TabControl
#Region "Constructor And Distructor"
   
    Public Sub New()
        Me.SetStyle(ControlStyles.UserPaint Or ControlStyles.AllPaintingInWmPaint Or ControlStyles.Opaque Or ControlStyles.ResizeRedraw, True)
        Me._BackBuffer = New Bitmap(Me.Width, Me.Height)
        Me._BackBufferGraphics = Graphics.FromImage(Me._BackBuffer)
        Me._TabBuffer = New Bitmap(Me.Width, Me.Height)
        Me._TabBufferGraphics = Graphics.FromImage(Me._TabBuffer)
        Me.TabStyle = OpenTabStyle.MSOffice2007
    End Sub
 
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        MyBase.Dispose(disposing)
        If disposing Then
            If Me._BackImage IsNot Nothing Then
                Me._BackImage.Dispose()
            End If
            If Me._BackBufferGraphics IsNot Nothing Then
                Me._BackBufferGraphics.Dispose()
            End If
            If Me._BackBuffer IsNot Nothing Then
                Me._BackBuffer.Dispose()
            End If
            If Me._TabBufferGraphics IsNot Nothing Then
                Me._TabBufferGraphics.Dispose()
            End If
            If Me._TabBuffer IsNot Nothing Then
                Me._TabBuffer.Dispose()
            End If

            If Me._StyleProvider IsNot Nothing Then
                Me._StyleProvider.Dispose()
            End If
        End If
    End Sub

#End Region

#Region "Private variables"

    Private _BackImage As Bitmap
    Private _BackBuffer As Bitmap
    Private _BackBufferGraphics As Graphics
    Private _TabBuffer As Bitmap
    Private _TabBufferGraphics As Graphics
    Private _BackgroundColor As Color
    Private _OldValue As Integer
    Private _DragStartPosition As Point = Point.Empty
    Private _Style As OpenTabStyle
    Private _StyleProvider As OpenStyleProvider
    Private _TabPages As List(Of TabPage)

#End Region

#Region "Public properties"

 
    <Category("MBTabControl Style"), Description("Get and Set the TabStyle for MBTabControl"), Browsable(False)> _
    Public Property TabStyle() As OpenTabStyle
        Get
            Return Me._Style
        End Get
        Set(ByVal value As OpenTabStyle)
            If Me._Style <> value Then
                Me._Style = value
                Me._StyleProvider = OpenStyleProvider.CreateProvider(Me)
                Me.Invalidate()
            End If
        End Set
    End Property
 
    <Category("Appearance"), RefreshProperties(RefreshProperties.All), Browsable(True)> _
    Public Shadows Property Multiline() As Boolean
        Get
            Return MyBase.Multiline
        End Get
        Set(ByVal value As Boolean)
            MyBase.Multiline = value
            Me.Invalidate()
        End Set
    End Property
 
    <Browsable(False), EditorBrowsable(EditorBrowsableState.Never)> _
    Public Shadows Property TabPadding() As Point
        Get
            Return Me._StyleProvider.Padding
        End Get
        Set(ByVal value As Point)
            Me._StyleProvider.Padding = value
        End Set
    End Property
 
    Public Overrides Property RightToLeftLayout() As Boolean
        Get
            Return MyBase.RightToLeftLayout
        End Get
        Set(ByVal value As Boolean)
            MyBase.RightToLeftLayout = value
            Me.UpdateStyles()
        End Set
    End Property
   
    <Category("Appearance"), Description("Get or Set Tab Alignment for MBTabControl"), Browsable(True)> _
    Public Shadows Property Alignment() As TabAlignment
        Get
            Return MyBase.Alignment
        End Get
        Set(ByVal value As TabAlignment)
            MyBase.Alignment = value
            Select Case value
                Case TabAlignment.Top, TabAlignment.Bottom
                    Me.Multiline = False
                    Exit Select
                Case TabAlignment.Left, TabAlignment.Right
                    Me.Multiline = True
                    Exit Select
            End Select
        End Set
    End Property

    <Browsable(False), EditorBrowsable(EditorBrowsableState.Never)> _
    <System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA1801:ReviewUnusedParameters", MessageId:="value")> _
    Public Shadows Property Appearance() As TabAppearance
        Get
            Return MyBase.Appearance
        End Get
        Set(ByVal value As TabAppearance)
            MyBase.Appearance = TabAppearance.Normal
        End Set
    End Property
 
    Public Overrides ReadOnly Property DisplayRectangle() As Rectangle
        Get
            If Me._Style = OpenTabStyle.None Then
                Return New Rectangle(0, 0, Width, Height)
            Else
                Dim tabStripHeight As Integer = 0
                Dim itemHeight As Integer = 0
                If Me.Alignment <= TabAlignment.Bottom Then
                    itemHeight = Me.ItemSize.Height
                Else
                    itemHeight = Me.ItemSize.Width
                End If

                tabStripHeight = 5 + (itemHeight * Me.RowCount)

                Dim rect As New Rectangle(4, tabStripHeight, Width - 8, Height - tabStripHeight - 4)
                Select Case Me.Alignment
                    Case TabAlignment.Top
                        rect = New Rectangle(4, tabStripHeight, Width - 8, Height - tabStripHeight - 4)
                        Exit Select
                    Case TabAlignment.Bottom
                        rect = New Rectangle(4, 4, Width - 8, Height - tabStripHeight - 4)
                        Exit Select
                    Case TabAlignment.Left
                        rect = New Rectangle(tabStripHeight, 4, Width - tabStripHeight - 4, Height - 8)
                        Exit Select
                    Case TabAlignment.Right
                        rect = New Rectangle(4, 4, Width - tabStripHeight - 4, Height - 8)
                        Exit Select
                End Select
                Return rect
            End If
        End Get
    End Property
  
    <Browsable(False)> _
    Public ReadOnly Property ActiveIndex() As Integer
        Get
            Dim hitTestInfo As New NativeMethods.TCHITTESTINFO(Me.PointToClient(Control.MousePosition))
            Dim index As Integer = NativeMethods.SendMessage(Me.Handle, NativeMethods.TCM_HITTEST, IntPtr.Zero, NativeMethods.ToIntPtr(hitTestInfo)).ToInt32()
            If index = -1 Then
                Return -1
            Else
                If Me.TabPages(index).Enabled Then
                    Return index
                Else
                    Return -1
                End If
            End If
        End Get
    End Property
    <Browsable(False)> _
        Public ReadOnly Property ActiveTab() As TabPage
        Get
            Dim activeIndex As Integer = Me.ActiveIndex
            If activeIndex > -1 Then
                Return Me.TabPages(activeIndex)
            Else
                Return Nothing
            End If
        End Get
    End Property

    Public Shadows ReadOnly Property MousePosition() As Point
        Get
            Dim loc As Point = Me.PointToClient(Control.MousePosition)
            If Me.RightToLeftLayout Then
                loc.X = (Me.Width - loc.X)
            End If
            Return loc
        End Get
    End Property
 
    Public Property TabBorderColor() As System.Drawing.Color
        Get
            Return Me._StyleProvider.BorderColor
        End Get
        Set(ByVal value As System.Drawing.Color)
            Me._StyleProvider.BorderColor = value
            Me.Invalidate()
        End Set
    End Property

    Public Property TabHotBorderColor() As Color
        Get
            Return Me._StyleProvider.BorderColorHot
        End Get
        Set(ByVal value As Color)
            Me._StyleProvider.BorderColorHot = value
            Me.Invalidate()
        End Set
    End Property

    Public Property SelectedTabBorderColor() As Color
        Get
            Return Me._StyleProvider.BorderColorSelected
        End Get
        Set(ByVal value As Color)
            Me._StyleProvider.BorderColorSelected = value
            Me.Invalidate()
        End Set
    End Property
  
    Public Property FocusColor() As Color
        Get
            Return Me._StyleProvider.FocusColor
        End Get
        Set(ByVal value As Color)
            Me._StyleProvider.FocusColor = value
            Me.Invalidate()
        End Set
    End Property
 
    Public Property CloseButtonColor() As Color
        Get
            Return Me._StyleProvider.CloserColor
        End Get
        Set(ByVal value As Color)
            Me._StyleProvider.CloserColor = value
            Me.Invalidate()
        End Set
    End Property
  
    Public Property ActivateCloseButtonColor() As Color
        Get
            Return Me._StyleProvider.CloserColorActive
        End Get
        Set(ByVal value As Color)
            Me._StyleProvider.CloserColorActive = value
            Me.Invalidate()
        End Set
    End Property

    Public Property FocuseTrack() As Boolean
        Get
            Return Me._StyleProvider.FocusTrack
        End Get
        Set(ByVal value As Boolean)
            Me._StyleProvider.FocusTrack = value
            Me.Invalidate()
        End Set
    End Property

    Public Property TabHotTrack() As Boolean
        Get
            Return Me._StyleProvider.HotTrack
        End Get
        Set(ByVal value As Boolean)
            Me._StyleProvider.HotTrack = value
            Me.Invalidate()
        End Set
    End Property

    Public Property TabImageAlignment() As ContentAlignment
        Get
            Return Me._StyleProvider.ImageAlign
        End Get
        Set(ByVal value As ContentAlignment)
            Me._StyleProvider.ImageAlign = value
            Me.Invalidate()
        End Set
    End Property
 
    Public Property Radius() As Int32
        Get
            Return Me._StyleProvider.Radius
        End Get
        Set(ByVal value As Int32)
            Me._StyleProvider.Radius = value
            Me.Invalidate()
        End Set
    End Property
  
    Public Property TabCloseButton() As Boolean
        Get
            Return Me._StyleProvider.TabCloseButton
        End Get
        Set(ByVal value As Boolean)
            Me._StyleProvider.TabCloseButton = value
            Me.Invalidate()
        End Set
    End Property

    Public Property TabTextColor() As Color
        Get
            Return Me._StyleProvider.TextColorSelected
        End Get
        Set(ByVal value As Color)
            Me._StyleProvider.TextColorSelected = value
            Me.Invalidate()
        End Set
    End Property

#End Region

#Region "Public Methods"
 
    ''' <param name="page">Page As TabPage</param>
    Public Sub HideTab(ByVal page As TabPage)
        If page IsNot Nothing AndAlso Me.TabPages.Contains(page) Then
            Me.BackupTabPages()
            Me.TabPages.Remove(page)
        End If
    End Sub

    ''' <param name="index">Index As Integer</param>
    Public Sub HideTab(ByVal index As Integer)
        If Me.IsValidTabIndex(index) Then
            Me.HideTab(Me._TabPages(index))
        End If
    End Sub

    ''' <param name="key">Key As String</param>
    Public Sub HideTab(ByVal key As String)
        If Me.TabPages.ContainsKey(key) Then
            Me.HideTab(Me.TabPages(key))
        End If
    End Sub

    ''' <param name="page">Page As TabPage</param>
    Public Sub ShowTab(ByVal page As TabPage)
        If page IsNot Nothing Then
            If Me._TabPages IsNot Nothing Then
                If Not Me.TabPages.Contains(page) AndAlso Me._TabPages.Contains(page) Then
                    Dim pageIndex As Integer = Me._TabPages.IndexOf(page)
                    If pageIndex > 0 Then
                        Dim start As Integer = pageIndex - 1
                        For index As Integer = start To 0 Step -1
                            If Me.TabPages.Contains(Me._TabPages(index)) Then
                                pageIndex = Me.TabPages.IndexOf(Me._TabPages(index)) + 1
                                Exit For
                            End If
                        Next
                    End If
                    If (pageIndex >= 0) AndAlso (pageIndex < Me.TabPages.Count) Then
                        Me.TabPages.Insert(pageIndex, page)
                    Else
                        Me.TabPages.Add(page)
                    End If
                End If
            Else
                If Not Me.TabPages.Contains(page) Then
                    Me.TabPages.Add(page)
                End If
            End If
        End If
    End Sub

    ''' <param name="index">Index As Integer</param>
    Public Sub ShowTab(ByVal index As Integer)
        If Me.IsValidTabIndex(index) Then
            Me.ShowTab(Me._TabPages(index))
        End If
    End Sub

    ''' <param name="key">Key As Strin</param>
    Public Sub ShowTab(ByVal key As String)
        If Me._TabPages IsNot Nothing Then
            Dim tab As TabPage = Me._TabPages.Find(Function(page As TabPage) page.Name.Equals(key, StringComparison.OrdinalIgnoreCase))
            Me.ShowTab(tab)
        End If
    End Sub
 
    ''' <param name="index">Index As Integer</param>
    ''' <returns>Return Boolean Value</returns>
    Private Function IsValidTabIndex(ByVal index As Integer) As Boolean
        Me.BackupTabPages()
        Return ((index >= 0) AndAlso (index < Me._TabPages.Count))
    End Function
  
    Private Sub BackupTabPages()
        If Me._TabPages Is Nothing Then
            Me._TabPages = New List(Of TabPage)()
            For Each page As TabPage In Me.TabPages
                Me._TabPages.Add(page)
            Next
        End If
    End Sub

#End Region

#Region "Protected Methods"
    
    ''' <param name="e">e As EventArgs</param>
    Protected Overrides Sub OnFontChanged(ByVal e As EventArgs)
        Dim hFont As IntPtr = Me.Font.ToHfont()
        NativeMethods.SendMessage(Me.Handle, NativeMethods.WM_SETFONT, hFont, CType(-1, IntPtr))
        NativeMethods.SendMessage(Me.Handle, NativeMethods.WM_FONTCHANGE, IntPtr.Zero, IntPtr.Zero)
        Me.UpdateStyles()
        If Me.Visible Then
            Me.Invalidate()
        End If
    End Sub

    ''' <param name="e">e As EventArgs</param>
    Protected Overrides Sub OnResize(ByVal e As EventArgs)
        If Me.Width > 0 AndAlso Me.Height > 0 Then
            If Me._BackImage IsNot Nothing Then
                Me._BackImage.Dispose()
                Me._BackImage = Nothing
            End If
            If Me._BackBufferGraphics IsNot Nothing Then
                Me._BackBufferGraphics.Dispose()
            End If
            If Me._BackBuffer IsNot Nothing Then
                Me._BackBuffer.Dispose()
            End If
            Me._BackBuffer = New Bitmap(Me.Width, Me.Height)
            Me._BackBufferGraphics = Graphics.FromImage(Me._BackBuffer)
            If Me._TabBufferGraphics IsNot Nothing Then
                Me._TabBufferGraphics.Dispose()
            End If
            If Me._TabBuffer IsNot Nothing Then
                Me._TabBuffer.Dispose()
            End If
            Me._TabBuffer = New Bitmap(Me.Width, Me.Height)
            Me._TabBufferGraphics = Graphics.FromImage(Me._TabBuffer)
            If Me._BackImage IsNot Nothing Then
                Me._BackImage.Dispose()
                Me._BackImage = Nothing
            End If
        End If
        MyBase.OnResize(e)
    End Sub

    ''' <param name="e">e As EventArgs</param>
    Protected Overrides Sub OnParentBackColorChanged(ByVal e As EventArgs)
        If Me._BackImage IsNot Nothing Then
            Me._BackImage.Dispose()
            Me._BackImage = Nothing
        End If
        MyBase.OnParentBackColorChanged(e)
    End Sub
  
    ''' <param name="e">e As EventArgs</param>
    Protected Overrides Sub OnParentBackgroundImageChanged(ByVal e As EventArgs)
        If Me._BackImage IsNot Nothing Then
            Me._BackImage.Dispose()
            Me._BackImage = Nothing
        End If
        MyBase.OnParentBackgroundImageChanged(e)
    End Sub
 
    ''' <param name="graphics">graphics As Graphics</param>
    ''' <param name="clipRect">ClientRectangle As Rectangle</param>
    Protected Sub PaintTransparentBackground(ByVal graphics As Graphics, ByVal clipRect As Rectangle)
        If (Me.Parent IsNot Nothing) Then
            clipRect.Offset(Me.Location)
            Dim state As GraphicsState = graphics.Save()
            graphics.TranslateTransform(CSng(-Me.Location.X), CSng(-Me.Location.Y))
            graphics.SmoothingMode = SmoothingMode.HighSpeed
            Dim e As New PaintEventArgs(graphics, clipRect)
            Try
                Me.InvokePaintBackground(Me.Parent, e)
                Me.InvokePaint(Me.Parent, e)
            Finally
                graphics.Restore(state)
                clipRect.Offset(-Me.Location.X, -Me.Location.Y)
            End Try
        End If
    End Sub
 
    ''' <param name="e">e As EventArgs</param>
    Protected Overrides Sub OnPaint(ByVal e As PaintEventArgs)
        If e.ClipRectangle.Equals(Me.ClientRectangle) Then
            Me.CustomPaint(e.Graphics)
        Else
            Me.Invalidate()
        End If
    End Sub
  
    Protected Overrides Sub OnCreateControl()
        MyBase.OnCreateControl()
        Me.OnFontChanged(EventArgs.Empty)
    End Sub
  
    Protected Overrides Sub Finalize()
        MyBase.Finalize()
    End Sub
   
    ''' <param name="e">e As EventArgs</param>
    Protected Overrides Sub OnParentChanged(ByVal e As EventArgs)
        MyBase.OnParentChanged(e)
        If Me.Parent IsNot Nothing Then
            AddHandler Me.Parent.Resize, AddressOf Me.OnParentResize
        End If
    End Sub
   
    ''' <param name="e">e As TabControlCancelEventArgs</param>
    Protected Overrides Sub OnSelecting(ByVal e As TabControlCancelEventArgs)
        MyBase.OnSelecting(e)
        If e.Action = TabControlAction.Selecting AndAlso e.TabPage IsNot Nothing AndAlso Not e.TabPage.Enabled Then
            e.Cancel = True
        End If
    End Sub
    
    ''' <param name="e">e As EventArgs</param>
    Protected Overrides Sub OnMove(ByVal e As EventArgs)
        If Me.Width > 0 AndAlso Me.Height > 0 Then
            If Me._BackImage IsNot Nothing Then
                Me._BackImage.Dispose()
                Me._BackImage = Nothing
            End If
        End If
        MyBase.OnMove(e)
        Me.Invalidate()
    End Sub

    ''' <param name="e">e As EventArgs</param>
    Protected Overrides Sub OnControlAdded(ByVal e As ControlEventArgs)
        MyBase.OnControlAdded(e)
        If Me.Visible Then
            Me.Invalidate()
        End If
    End Sub
 
    ''' <param name="e">e As EventArgs</param>
    Protected Overrides Sub OnControlRemoved(ByVal e As ControlEventArgs)
        MyBase.OnControlRemoved(e)
        If Me.Visible Then
            Me.Invalidate()
        End If
    End Sub

    <UIPermission(SecurityAction.LinkDemand, Window:=UIPermissionWindow.AllWindows)> _
    Protected Overrides Function ProcessMnemonic(ByVal charCode As Char) As Boolean
        For Each page As TabPage In Me.TabPages
            If IsMnemonic(charCode, page.Text) Then
                Me.SelectedTab = page
                Return True
            End If
        Next
        Return MyBase.ProcessMnemonic(charCode)
    End Function

    ''' <param name="e">e As EventArgs</param>
    Protected Overrides Sub OnSelectedIndexChanged(ByVal e As EventArgs)
        MyBase.OnSelectedIndexChanged(e)
    End Sub

    <SecurityPermission(SecurityAction.LinkDemand, Flags:=SecurityPermissionFlag.UnmanagedCode)> _
    <System.Diagnostics.DebuggerStepThrough()> _
    Protected Overrides Sub WndProc(ByRef m As Message)
        Select Case m.Msg
            Case NativeMethods.WM_HSCROLL
                MyBase.WndProc(m)
                Me.OnHScroll(New ScrollEventArgs(CType(NativeMethods.LoWord(m.WParam), ScrollEventType), _OldValue, NativeMethods.HiWord(m.WParam), ScrollOrientation.HorizontalScroll))
                Exit Select
            Case Else
                MyBase.WndProc(m)
                Exit Select
        End Select
    End Sub
  
    ''' <param name="e">e As EventArgs</param>
    Protected Overrides Sub OnMouseClick(ByVal e As MouseEventArgs)
        Dim index As Integer = Me.ActiveIndex
        If index > -1 AndAlso Me.TabImageClickEvent IsNot Nothing AndAlso (Me.TabPages(index).ImageIndex > -1 OrElse Not String.IsNullOrEmpty(Me.TabPages(index).ImageKey)) AndAlso Me.GetTabImageRect(index).Contains(Me.MousePosition) Then
            Me.OnTabImageClick(New TabControlEventArgs(Me.TabPages(index), index, TabControlAction.Selected))
            MyBase.OnMouseClick(e)
        ElseIf Not Me.DesignMode AndAlso index > -1 AndAlso Me._StyleProvider.TabCloseButton AndAlso Me.GetTabCloserRect(index).Contains(Me.MousePosition) Then
            Dim tab As TabPage = Me.ActiveTab
            Dim args As New TabControlCancelEventArgs(tab, index, False, TabControlAction.Deselecting)
            Me.OnTabClosing(args)
            If Not args.Cancel Then
                Me.TabPages.Remove(tab)
                tab.Dispose()
            End If
        Else
            MyBase.OnMouseClick(e)
        End If
    End Sub
   
    ''' <param name="e">e As EventArgs</param>
    Protected Overridable Sub OnTabImageClick(ByVal e As TabControlEventArgs)
        RaiseEvent TabImageClick(Me, e)
    End Sub
   
    ''' <param name="e">e As EventArgs</param>
    Protected Overridable Sub OnTabClosing(ByVal e As TabControlCancelEventArgs)
        RaiseEvent TabClosing(Me, e)
    End Sub
  
    ''' <param name="e">e As EventArgs</param>
    Protected Overridable Sub OnHScroll(ByVal e As ScrollEventArgs)
        Me.Invalidate()
        RaiseEvent HScroll(Me, e)
        If e.Type = ScrollEventType.EndScroll Then
            Me._OldValue = e.NewValue
        End If
    End Sub
   
    ''' <param name="e">e As EventArgs</param>
    Protected Overrides Sub OnMouseMove(ByVal e As MouseEventArgs)
        MyBase.OnMouseMove(e)
        If Me._StyleProvider.TabCloseButton Then
            Dim tabRect As Rectangle = Me._StyleProvider.GetTabRect(Me.ActiveIndex)
            If tabRect.Contains(Me.MousePosition) Then
                Me.Invalidate()
            End If
        End If
    End Sub

#End Region

#Region "Private Methods"

    Private Sub AddPageBorder(ByVal path As GraphicsPath, ByVal pageBounds As Rectangle, ByVal tabBounds As Rectangle)
        Select Case Me.Alignment
            Case TabAlignment.Top
                path.AddLine(tabBounds.Right, pageBounds.Y, pageBounds.Right, pageBounds.Y)
                path.AddLine(pageBounds.Right, pageBounds.Y, pageBounds.Right, pageBounds.Bottom)
                path.AddLine(pageBounds.Right, pageBounds.Bottom, pageBounds.X, pageBounds.Bottom)
                path.AddLine(pageBounds.X, pageBounds.Bottom, pageBounds.X, pageBounds.Y)
                path.AddLine(pageBounds.X, pageBounds.Y, tabBounds.X, pageBounds.Y)
                Exit Select
            Case TabAlignment.Bottom
                path.AddLine(tabBounds.X, pageBounds.Bottom, pageBounds.X, pageBounds.Bottom)
                path.AddLine(pageBounds.X, pageBounds.Bottom, pageBounds.X, pageBounds.Y)
                path.AddLine(pageBounds.X, pageBounds.Y, pageBounds.Right, pageBounds.Y)
                path.AddLine(pageBounds.Right, pageBounds.Y, pageBounds.Right, pageBounds.Bottom)
                path.AddLine(pageBounds.Right, pageBounds.Bottom, tabBounds.Right, pageBounds.Bottom)
                Exit Select
            Case TabAlignment.Left
                path.AddLine(pageBounds.X, tabBounds.Y, pageBounds.X, pageBounds.Y)
                path.AddLine(pageBounds.X, pageBounds.Y, pageBounds.Right, pageBounds.Y)
                path.AddLine(pageBounds.Right, pageBounds.Y, pageBounds.Right, pageBounds.Bottom)
                path.AddLine(pageBounds.Right, pageBounds.Bottom, pageBounds.X, pageBounds.Bottom)
                path.AddLine(pageBounds.X, pageBounds.Bottom, pageBounds.X, tabBounds.Bottom)
                Exit Select
            Case TabAlignment.Right
                path.AddLine(pageBounds.Right, tabBounds.Bottom, pageBounds.Right, pageBounds.Bottom)
                path.AddLine(pageBounds.Right, pageBounds.Bottom, pageBounds.X, pageBounds.Bottom)
                path.AddLine(pageBounds.X, pageBounds.Bottom, pageBounds.X, pageBounds.Y)
                path.AddLine(pageBounds.X, pageBounds.Y, pageBounds.Right, pageBounds.Y)
                path.AddLine(pageBounds.Right, pageBounds.Y, pageBounds.Right, tabBounds.Y)
                Exit Select
        End Select
    End Sub
 
    Private Sub DrawTabPage(ByVal index As Integer, ByVal graphics As Graphics)
        graphics.SmoothingMode = SmoothingMode.HighSpeed
        Using tabPageBorderPath As GraphicsPath = Me.GetTabPageBorder(index)
            Using fillBrush As Brush = Me._StyleProvider.GetPageBackgroundBrush(index)
                graphics.FillPath(fillBrush, tabPageBorderPath)
            End Using
            If Me._Style <> OpenTabStyle.None Then
                Me._StyleProvider.PaintTab(index, graphics)
                Me.DrawTabImage(index, graphics)
                Me.DrawTabText(index, graphics)
            End If
            Me.DrawTabBorder(tabPageBorderPath, index, graphics)
        End Using
    End Sub
  
    Private Sub DrawTabBorder(ByVal path As GraphicsPath, ByVal index As Integer, ByVal graphics As Graphics)
        graphics.SmoothingMode = SmoothingMode.HighQuality
        Dim borderColor As Color
        If index = Me.SelectedIndex Then
            borderColor = Me._StyleProvider.BorderColorSelected
        ElseIf Me._StyleProvider.HotTrack AndAlso index = Me.ActiveIndex Then
            borderColor = Me._StyleProvider.BorderColorHot
        Else
            borderColor = Me._StyleProvider.BorderColor
        End If

        Using borderPen As New Pen(borderColor)
            graphics.DrawPath(borderPen, path)
        End Using
    End Sub
  
    Private Sub DrawTabText(ByVal index As Integer, ByVal graphics As Graphics)
        graphics.SmoothingMode = SmoothingMode.HighQuality
        graphics.TextRenderingHint = System.Drawing.Text.TextRenderingHint.ClearTypeGridFit
        Dim tabBounds As Rectangle = Me.GetTabTextRect(index)
        If Me.SelectedIndex = index Then
            Using textBrush As Brush = New SolidBrush(Me._StyleProvider.TextColorSelected)
                graphics.DrawString(Me.TabPages(index).Text, Me.Font, textBrush, tabBounds, Me.GetStringFormat())
            End Using
        Else
            If Me.TabPages(index).Enabled Then

                Using textBrush As Brush = New SolidBrush(Me._StyleProvider.TextColorSelected)
                    graphics.DrawString(Me.TabPages(index).Text, Me.Font, textBrush, tabBounds, Me.GetStringFormat())
                End Using
            Else
                Using textBrush As Brush = New SolidBrush(Me._StyleProvider.TextColorDisabled)
                    graphics.DrawString(Me.TabPages(index).Text, Me.Font, textBrush, tabBounds, Me.GetStringFormat())
                End Using
            End If
        End If
    End Sub
   
    Private Sub DrawTabImage(ByVal index As Integer, ByVal graphics As Graphics)
        Dim tabImage As Image = Nothing
        If Me.TabPages(index).ImageIndex > -1 AndAlso Me.ImageList IsNot Nothing AndAlso Me.ImageList.Images.Count > Me.TabPages(index).ImageIndex Then
            tabImage = Me.ImageList.Images(Me.TabPages(index).ImageIndex)
        ElseIf (Not String.IsNullOrEmpty(Me.TabPages(index).ImageKey) AndAlso Not Me.TabPages(index).ImageKey.Equals("(none)", StringComparison.OrdinalIgnoreCase)) AndAlso Me.ImageList IsNot Nothing AndAlso Me.ImageList.Images.ContainsKey(Me.TabPages(index).ImageKey) Then
            tabImage = Me.ImageList.Images(Me.TabPages(index).ImageKey)
        End If
        If tabImage IsNot Nothing Then
            If Me.RightToLeftLayout Then
                tabImage.RotateFlip(RotateFlipType.RotateNoneFlipX)
            End If
            Dim imageRect As Rectangle = Me.GetTabImageRect(index)
            If Me.TabPages(index).Enabled Then
                graphics.DrawImage(tabImage, imageRect)
            Else
                ControlPaint.DrawImageDisabled(graphics, tabImage, imageRect.X, imageRect.Y, Color.Transparent)
            End If
        End If
    End Sub

    Private Sub CustomPaint(ByVal screenGraphics As Graphics)
        If Me.Width > 0 AndAlso Me.Height > 0 Then
            If Me._BackImage Is Nothing Then
                Me._BackImage = New Bitmap(Me.Width, Me.Height)
                Dim backGraphics As Graphics = Graphics.FromImage(Me._BackImage)
                backGraphics.Clear(Color.Transparent)
                Me.PaintTransparentBackground(backGraphics, Me.ClientRectangle)
            End If
            Me._BackBufferGraphics.Clear(Color.Gray)
            Me._BackBufferGraphics.DrawImageUnscaled(Me._BackImage, 0, 0)
            Me._TabBufferGraphics.Clear(Me.Parent.BackColor)
            If Me.TabCount > 0 Then
                If Me.Alignment <= TabAlignment.Bottom AndAlso Not Me.Multiline Then
                    Me._TabBufferGraphics.Clip = New Region(New RectangleF(Me.ClientRectangle.X + 3, Me.ClientRectangle.Y, Me.ClientRectangle.Width - 6, Me.ClientRectangle.Height))
                End If
                If Me.Multiline Then
                    For row As Integer = 0 To Me.RowCount - 1
                        For index As Integer = Me.TabCount - 1 To 0 Step -1
                            If index <> Me.SelectedIndex AndAlso (Me.RowCount = 1 OrElse Me.GetTabRow(index) = row) Then
                                Me.DrawTabPage(index, Me._TabBufferGraphics)
                            End If
                        Next
                    Next
                Else
                    For index As Integer = Me.TabCount - 1 To 0 Step -1
                        If index <> Me.SelectedIndex Then
                            Me.DrawTabPage(index, Me._TabBufferGraphics)
                        End If
                    Next
                End If
                If Me.SelectedIndex > -1 Then
                    Me.DrawTabPage(Me.SelectedIndex, Me._TabBufferGraphics)
                End If
            End If
            Me._TabBufferGraphics.Flush()
            Dim alphaMatrix As New ColorMatrix()
            alphaMatrix.Matrix00 = InlineAssignHelper(alphaMatrix.Matrix11, InlineAssignHelper(alphaMatrix.Matrix22, InlineAssignHelper(alphaMatrix.Matrix44, 1)))
            alphaMatrix.Matrix33 = Me._StyleProvider.Opacity
            Using alphaAttributes As New ImageAttributes()
                alphaAttributes.SetColorMatrix(alphaMatrix)
                Me._BackBufferGraphics.DrawImage(Me._TabBuffer, New Rectangle(0, 0, Me._TabBuffer.Width, Me._TabBuffer.Height), 0, 0, Me._TabBuffer.Width, Me._TabBuffer.Height, _
                 GraphicsUnit.Pixel, alphaAttributes)
            End Using
            Me._BackBufferGraphics.Flush()
            If Me.RightToLeftLayout Then
                screenGraphics.DrawImageUnscaled(Me._BackBuffer, -1, 0)
            Else
                screenGraphics.DrawImageUnscaled(Me._BackBuffer, 0, 0)
            End If
        End If
    End Sub

    Private Sub OnParentResize(ByVal sender As Object, ByVal e As EventArgs)
        If Me.Visible Then
            Me.Invalidate()
        End If
    End Sub

#End Region

#Region "Events"

  
    <Category("MBTabControl Action")> _
    Public Event HScroll As ScrollEventHandler
   
    <Category("MBTabControl Action")> _
    Public Event TabImageClick As EventHandler(Of TabControlEventArgs)
    
    <Category("MBTabControl Action")> _
    Public Event TabClosing As EventHandler(Of TabControlCancelEventArgs)

#End Region

#Region "Private Function"
   
    Private Function GetStringFormat() As StringFormat
        Dim format As StringFormat = Nothing
        Select Case Me.Alignment
            Case TabAlignment.Top, TabAlignment.Bottom
                format = New StringFormat()
                Exit Select
            Case TabAlignment.Left, TabAlignment.Right
                format = New StringFormat(StringFormatFlags.DirectionVertical)
                Exit Select
        End Select
        format.Alignment = StringAlignment.Center
        format.LineAlignment = StringAlignment.Center
        If Me.FindForm() IsNot Nothing AndAlso Me.FindForm().KeyPreview Then
            format.HotkeyPrefix = System.Drawing.Text.HotkeyPrefix.Show
        Else
            format.HotkeyPrefix = System.Drawing.Text.HotkeyPrefix.Hide
        End If
        If Me.RightToLeft = RightToLeft.Yes Then
            format.FormatFlags = format.FormatFlags Or StringFormatFlags.DirectionRightToLeft
        End If
        Return format
    End Function
 
    Private Function TextAlignment(ByVal textalign As ContentAlignment)
        Dim sf As StringFormat = New StringFormat()
        Select Case textalign
            Case ContentAlignment.TopLeft
                sf.LineAlignment = StringAlignment.Near
            Case ContentAlignment.TopCenter
                sf.LineAlignment = StringAlignment.Near
            Case ContentAlignment.TopRight
                sf.LineAlignment = StringAlignment.Near
            Case ContentAlignment.MiddleLeft
                sf.LineAlignment = StringAlignment.Center
            Case ContentAlignment.MiddleCenter
                sf.LineAlignment = StringAlignment.Center
            Case ContentAlignment.MiddleRight
                sf.LineAlignment = StringAlignment.Center
            Case ContentAlignment.BottomLeft
                sf.LineAlignment = StringAlignment.Far
            Case ContentAlignment.BottomCenter
                sf.LineAlignment = StringAlignment.Far
            Case ContentAlignment.BottomRight
                sf.LineAlignment = StringAlignment.Far
        End Select
        Select Case textalign
            Case ContentAlignment.TopLeft
                sf.Alignment = StringAlignment.Near
            Case ContentAlignment.MiddleLeft
                sf.Alignment = StringAlignment.Near
            Case ContentAlignment.BottomLeft
                sf.Alignment = StringAlignment.Near
            Case ContentAlignment.TopCenter
                sf.Alignment = StringAlignment.Center
            Case ContentAlignment.MiddleCenter
                sf.Alignment = StringAlignment.Center
            Case ContentAlignment.BottomCenter
                sf.Alignment = StringAlignment.Center
            Case ContentAlignment.TopRight
                sf.Alignment = StringAlignment.Far
            Case ContentAlignment.MiddleRight
                sf.Alignment = StringAlignment.Far
            Case ContentAlignment.BottomRight
                sf.Alignment = StringAlignment.Far
        End Select
        Return sf
    End Function
   
    Private Function GetTabPageBorder(ByVal index As Integer) As GraphicsPath
        Dim path As New GraphicsPath()
        Dim pageBounds As Rectangle = Me.GetPageBounds(index)
        Dim tabBounds As Rectangle = Me._StyleProvider.GetTabRect(index)
        Me._StyleProvider.AddTabBorder(path, tabBounds)
        Me.AddPageBorder(path, pageBounds, tabBounds)
        path.CloseFigure()
        Return path
    End Function

    Private Function GetTabTextRect(ByVal index As Integer) As Rectangle
        Dim textRect As New Rectangle()
        Using path As GraphicsPath = Me._StyleProvider.GetTabBorder(index)
            Dim tabBounds As RectangleF = path.GetBounds()
            textRect = New Rectangle(CInt(Math.Truncate(tabBounds.X)), CInt(Math.Truncate(tabBounds.Y)), CInt(Math.Truncate(tabBounds.Width)), CInt(Math.Truncate(tabBounds.Height)))
            Select Case Me.Alignment
                Case TabAlignment.Top
                    textRect.Y += 4
                    textRect.Height -= 6
                    Exit Select
                Case TabAlignment.Bottom
                    textRect.Y += 2
                    textRect.Height -= 6
                    Exit Select
                Case TabAlignment.Left
                    textRect.X += 4
                    textRect.Width -= 6
                    Exit Select
                Case TabAlignment.Right
                    textRect.X += 2
                    textRect.Width -= 6
                    Exit Select
            End Select
            If Me.ImageList IsNot Nothing AndAlso (Me.TabPages(index).ImageIndex > -1 OrElse (Not String.IsNullOrEmpty(Me.TabPages(index).ImageKey) AndAlso Not Me.TabPages(index).ImageKey.Equals("(none)", StringComparison.OrdinalIgnoreCase))) Then
                Dim imageRect As Rectangle = Me.GetTabImageRect(index)
                If (Me._StyleProvider.ImageAlign And NativeMethods.AnyLeftAlign) <> CType(0, ContentAlignment) Then
                    If Me.Alignment <= TabAlignment.Bottom Then
                        textRect.X = imageRect.Right + 4
                        textRect.Width -= (textRect.Right - CInt(Math.Truncate(tabBounds.Right)))
                    Else
                        textRect.Y = imageRect.Y + 4
                        textRect.Height -= (textRect.Bottom - CInt(Math.Truncate(tabBounds.Bottom)))
                    End If
                    If Me._StyleProvider.TabCloseButton Then
                        Dim closerRect As Rectangle = Me.GetTabCloserRect(index)
                        If Me.Alignment <= TabAlignment.Bottom Then
                            If Me.RightToLeftLayout Then
                                textRect.Width -= (closerRect.Right + 4 - textRect.X)
                                textRect.X = closerRect.Right + 4
                            Else
                                textRect.Width -= (CInt(Math.Truncate(tabBounds.Right)) - closerRect.X + 4)
                            End If
                        Else
                            If Me.RightToLeftLayout Then
                                textRect.Height -= (closerRect.Bottom + 4 - textRect.Y)
                                textRect.Y = closerRect.Bottom + 4
                            Else
                                textRect.Height -= (CInt(Math.Truncate(tabBounds.Bottom)) - closerRect.Y + 4)
                            End If
                        End If
                    End If
                ElseIf (Me._StyleProvider.ImageAlign And NativeMethods.AnyCenterAlign) <> CType(0, ContentAlignment) Then
                    If Me._StyleProvider.TabCloseButton Then
                        Dim closerRect As Rectangle = Me.GetTabCloserRect(index)
                        If Me.Alignment <= TabAlignment.Bottom Then
                            If Me.RightToLeftLayout Then
                                textRect.Width -= (closerRect.Right + 4 - textRect.X)
                                textRect.X = closerRect.Right + 4
                            Else
                                textRect.Width -= (CInt(Math.Truncate(tabBounds.Right)) - closerRect.X + 4)
                            End If
                        Else
                            If Me.RightToLeftLayout Then
                                textRect.Height -= (closerRect.Bottom + 4 - textRect.Y)
                                textRect.Y = closerRect.Bottom + 4
                            Else
                                textRect.Height -= (CInt(Math.Truncate(tabBounds.Bottom)) - closerRect.Y + 4)
                            End If
                        End If
                    End If
                Else
                    If Me.Alignment <= TabAlignment.Bottom Then
                        textRect.Width -= (CInt(Math.Truncate(tabBounds.Right)) - imageRect.X + 4)
                    Else
                        textRect.Height -= (CInt(Math.Truncate(tabBounds.Bottom)) - imageRect.Y + 4)
                    End If
                    If Me._StyleProvider.TabCloseButton Then
                        Dim closerRect As Rectangle = Me.GetTabCloserRect(index)
                        If Me.Alignment <= TabAlignment.Bottom Then
                            If Me.RightToLeftLayout Then
                                textRect.Width -= (closerRect.Right + 4 - textRect.X)
                                textRect.X = closerRect.Right + 4
                            Else
                                textRect.Width -= (CInt(Math.Truncate(tabBounds.Right)) - closerRect.X + 4)
                            End If
                        Else
                            If Me.RightToLeftLayout Then
                                textRect.Height -= (closerRect.Bottom + 4 - textRect.Y)
                                textRect.Y = closerRect.Bottom + 4
                            Else
                                textRect.Height -= (CInt(Math.Truncate(tabBounds.Bottom)) - closerRect.Y + 4)
                            End If
                        End If
                    End If
                End If
            Else
                If Me._StyleProvider.TabCloseButton Then
                    Dim closerRect As Rectangle = Me.GetTabCloserRect(index)
                    If Me.Alignment <= TabAlignment.Bottom Then
                        If Me.RightToLeftLayout Then
                            textRect.Width -= (closerRect.Right + 4 - textRect.X)
                            textRect.X = closerRect.Right + 4
                        Else
                            textRect.Width -= (CInt(Math.Truncate(tabBounds.Right)) - closerRect.X + 4)
                        End If
                    Else
                        If Me.RightToLeftLayout Then
                            textRect.Height -= (closerRect.Bottom + 4 - textRect.Y)
                            textRect.Y = closerRect.Bottom + 4
                        Else
                            textRect.Height -= (CInt(Math.Truncate(tabBounds.Bottom)) - closerRect.Y + 4)
                        End If
                    End If
                End If
            End If
            If Me.Alignment <= TabAlignment.Bottom Then
                While Not path.IsVisible(textRect.Right, textRect.Y) AndAlso textRect.Width > 0
                    textRect.Width -= 1
                End While
                While Not path.IsVisible(textRect.X, textRect.Y) AndAlso textRect.Width > 0
                    textRect.X += 1
                    textRect.Width -= 1
                End While
            Else
                While Not path.IsVisible(textRect.X, textRect.Bottom) AndAlso textRect.Height > 0
                    textRect.Height -= 1
                End While
                While Not path.IsVisible(textRect.X, textRect.Y) AndAlso textRect.Height > 0
                    textRect.Y += 1
                    textRect.Height -= 1
                End While
            End If
        End Using
        Return textRect
    End Function

    Private Function GetTabImageRect(ByVal index As Integer) As Rectangle
        Using tabBorderPath As GraphicsPath = Me._StyleProvider.GetTabBorder(index)
            Return Me.GetTabImageRect(tabBorderPath)
        End Using
    End Function

    Private Function GetTabImageRect(ByVal tabBorderPath As GraphicsPath) As Rectangle
        Dim imageRect As New Rectangle()
        Dim rect As RectangleF = tabBorderPath.GetBounds()
        Select Case Me.Alignment
            Case TabAlignment.Top
                rect.Y += 4
                rect.Height -= 6
                Exit Select
            Case TabAlignment.Bottom
                rect.Y += 2
                rect.Height -= 6
                Exit Select
            Case TabAlignment.Left
                rect.X += 4
                rect.Width -= 6
                Exit Select
            Case TabAlignment.Right
                rect.X += 2
                rect.Width -= 6
                Exit Select
        End Select
        If Me.Alignment <= TabAlignment.Bottom Then
            If (Me._StyleProvider.ImageAlign And NativeMethods.AnyLeftAlign) <> CType(0, ContentAlignment) Then
                imageRect = New Rectangle(CInt(Math.Truncate(rect.X)), CInt(Math.Truncate(rect.Y)) + CInt(Math.Truncate(Math.Floor(CDbl(CInt(Math.Truncate(rect.Height)) - 16) / 2))), 16, 16)
                While Not tabBorderPath.IsVisible(imageRect.X, imageRect.Y)
                    imageRect.X += 1
                End While
                imageRect.X += 4
            ElseIf (Me._StyleProvider.ImageAlign And NativeMethods.AnyCenterAlign) <> CType(0, ContentAlignment) Then
                imageRect = New Rectangle(CInt(Math.Truncate(rect.X)) + CInt(Math.Truncate(Math.Floor(CDbl((CInt(Math.Truncate(rect.Right)) - CInt(Math.Truncate(rect.X)) - CInt(Math.Truncate(rect.Height)) + 2) \ 2)))), CInt(Math.Truncate(rect.Y)) + CInt(Math.Truncate(Math.Floor(CDbl(CInt(Math.Truncate(rect.Height)) - 16) / 2))), 16, 16)
            Else
                imageRect = New Rectangle(CInt(Math.Truncate(rect.Right)), CInt(Math.Truncate(rect.Y)) + CInt(Math.Truncate(Math.Floor(CDbl(CInt(Math.Truncate(rect.Height)) - 16) / 2))), 16, 16)
                While Not tabBorderPath.IsVisible(imageRect.Right, imageRect.Y)
                    imageRect.X -= 1
                End While
                imageRect.X -= 4
                If Me._StyleProvider.TabCloseButton AndAlso Not Me.RightToLeftLayout Then
                    imageRect.X -= 10
                End If
            End If
        Else
            If (Me._StyleProvider.ImageAlign And NativeMethods.AnyLeftAlign) <> CType(0, ContentAlignment) Then
                imageRect = New Rectangle(CInt(Math.Truncate(rect.X)) + CInt(Math.Truncate(Math.Floor(CDbl(CInt(Math.Truncate(rect.Width)) - 16) / 2))), CInt(Math.Truncate(rect.Y)), 16, 16)
                While Not tabBorderPath.IsVisible(imageRect.X, imageRect.Y)
                    imageRect.Y += 1
                End While
                imageRect.Y += 4
            ElseIf (Me._StyleProvider.ImageAlign And NativeMethods.AnyCenterAlign) <> CType(0, ContentAlignment) Then
                imageRect = New Rectangle(CInt(Math.Truncate(rect.X)) + CInt(Math.Truncate(Math.Floor(CDbl(CInt(Math.Truncate(rect.Width)) - 16) / 2))), CInt(Math.Truncate(rect.Y)) + CInt(Math.Truncate(Math.Floor(CDbl((CInt(Math.Truncate(rect.Bottom)) - CInt(Math.Truncate(rect.Y)) - CInt(Math.Truncate(rect.Width)) + 2) \ 2)))), 16, 16)
            Else
                imageRect = New Rectangle(CInt(Math.Truncate(rect.X)) + CInt(Math.Truncate(Math.Floor(CDbl(CInt(Math.Truncate(rect.Width)) - 16) / 2))), CInt(Math.Truncate(rect.Bottom)), 16, 16)
                While Not tabBorderPath.IsVisible(imageRect.X, imageRect.Bottom)
                    imageRect.Y -= 1
                End While
                imageRect.Y -= 4
                If Me._StyleProvider.TabCloseButton AndAlso Not Me.RightToLeftLayout Then
                    imageRect.Y -= 10
                End If
            End If
        End If
        Return imageRect
    End Function

#End Region

#Region "Public Functions"

    Public Function GetPageBounds(ByVal index As Integer) As Rectangle
        Dim pageBounds As Rectangle = Me.TabPages(index).Bounds
        pageBounds.Width += 1
        pageBounds.Height += 1
        pageBounds.X -= 1
        pageBounds.Y -= 1
        If pageBounds.Bottom > Me.Height - 4 Then
            pageBounds.Height -= (pageBounds.Bottom - Me.Height + 4)
        End If
        Return pageBounds
    End Function

    Public Function GetTabRow(ByVal index As Integer) As Integer
        Dim rect As Rectangle = Me.GetTabRect(index)
        Dim row As Integer = -1
        Select Case Me.Alignment
            Case TabAlignment.Top
                row = (rect.Y - 2) \ rect.Height
                Exit Select
            Case TabAlignment.Bottom
                row = ((Me.Height - rect.Y - 2) \ rect.Height) - 1
                Exit Select
            Case TabAlignment.Left
                row = (rect.X - 2) \ rect.Width
                Exit Select
            Case TabAlignment.Right
                row = ((Me.Width - rect.X - 2) \ rect.Width) - 1
                Exit Select
        End Select
        Return row
    End Function

    Public Function GetTabPosition(ByVal index As Integer) As Point
        If Not Me.Multiline Then
            Return New Point(0, index)
        End If
        If Me.RowCount = 1 Then
            Return New Point(0, index)
        End If
        Dim row As Integer = Me.GetTabRow(index)
        Dim rect As Rectangle = Me.GetTabRect(index)
        Dim column As Integer = -1
        For testIndex As Integer = 0 To Me.TabCount - 1
            Dim testRect As Rectangle = Me.GetTabRect(testIndex)
            If Me.Alignment <= TabAlignment.Bottom Then
                If testRect.Y = rect.Y Then
                    column += 1
                End If
            Else
                If testRect.X = rect.X Then
                    column += 1
                End If
            End If
            If testRect.Location.Equals(rect.Location) Then
                Return New Point(row, column)
            End If
        Next
        Return New Point(0, 0)
    End Function

    Public Function IsFirstTabInRow(ByVal index As Integer) As Boolean
        If index < 0 Then
            Return False
        End If
        Dim firstTabinRow As Boolean = (index = 0)
        If Not firstTabinRow Then
            If Me.Alignment <= TabAlignment.Bottom Then
                If Me.GetTabRect(index).X = 2 Then
                    firstTabinRow = True
                End If
            Else
                If Me.GetTabRect(index).Y = 2 Then
                    firstTabinRow = True
                End If
            End If
        End If
        Return firstTabinRow
    End Function

    Public Function GetTabCloserRect(ByVal index As Integer) As Rectangle
        Dim closerRect As New Rectangle()
        Using path As GraphicsPath = Me._StyleProvider.GetTabBorder(index)
            Dim rect As RectangleF = path.GetBounds()
            Select Case Me.Alignment
                Case TabAlignment.Top
                    rect.Y += 4
                    rect.Height -= 6
                    Exit Select
                Case TabAlignment.Bottom
                    rect.Y += 2
                    rect.Height -= 6
                    Exit Select
                Case TabAlignment.Left
                    rect.X += 4
                    rect.Width -= 6
                    Exit Select
                Case TabAlignment.Right
                    rect.X += 2
                    rect.Width -= 6
                    Exit Select
            End Select
            If Me.Alignment <= TabAlignment.Bottom Then
                If Me.RightToLeftLayout Then
                    closerRect = New Rectangle(CInt(Math.Truncate(rect.Left)), CInt(Math.Truncate(rect.Y)) + CInt(Math.Truncate(Math.Floor(CDbl(CInt(Math.Truncate(rect.Height)) - 6) / 2))), 6, 6)
                    While Not path.IsVisible(closerRect.Left, closerRect.Y) AndAlso closerRect.Right < Me.Width
                        closerRect.X += 1
                    End While
                    closerRect.X += 4
                Else
                    closerRect = New Rectangle(CInt(Math.Truncate(rect.Right)), CInt(Math.Truncate(rect.Y)) + CInt(Math.Truncate(Math.Floor(CDbl(CInt(Math.Truncate(rect.Height)) - 6) / 2))), 6, 6)
                    While Not path.IsVisible(closerRect.Right, closerRect.Y) AndAlso closerRect.Right > -6
                        closerRect.X -= 1
                    End While
                    closerRect.X -= 4
                End If
            Else
                If Me.RightToLeftLayout Then
                    closerRect = New Rectangle(CInt(Math.Truncate(rect.X)) + CInt(Math.Truncate(Math.Floor(CDbl(CInt(Math.Truncate(rect.Width)) - 6) / 2))), CInt(Math.Truncate(rect.Top)), 6, 6)
                    While Not path.IsVisible(closerRect.X, closerRect.Top) AndAlso closerRect.Bottom < Me.Height
                        closerRect.Y += 1
                    End While
                    closerRect.Y += 4
                Else
                    closerRect = New Rectangle(CInt(Math.Truncate(rect.X)) + CInt(Math.Truncate(Math.Floor(CDbl(CInt(Math.Truncate(rect.Width)) - 6) / 2))), CInt(Math.Truncate(rect.Bottom)), 6, 6)
                    While Not path.IsVisible(closerRect.X, closerRect.Bottom) AndAlso closerRect.Top > -6
                        closerRect.Y -= 1
                    End While
                    closerRect.Y -= 4
                End If
            End If
        End Using
        Return closerRect
    End Function

    Private Shared Function InlineAssignHelper(Of T)(ByRef target As T, ByVal value As T) As T
        target = value
        Return value
    End Function

#End Region

End Class

#End Region

#Region "   Enumerations"

Public Enum OpenTabStyle
    None
    Normal
    MSOffice2007
End Enum
#End Region