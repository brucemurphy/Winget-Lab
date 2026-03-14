Imports System.IO
Imports System.Net.Http
Imports System.Diagnostics
Imports System.Text.RegularExpressions
Imports System.Threading.Tasks
Imports System.Windows
Imports System.Windows.Controls
Imports System.Windows.Input
Imports System.Windows.Media
Imports System.Windows.Media.Effects
Imports System.Windows.Media.Imaging
Imports System.Windows.Threading
Imports System.Xml.Linq
Imports Microsoft.Win32
Imports System.Runtime.InteropServices

Partial Public Class MainWindow
    Inherits Window
    Private Shared ReadOnly httpClient As New HttpClient()
    Private _headerValue As String = ""

    ' Activity indicator
    Private _wingetBlinkTimer As DispatcherTimer
    Private _wingetActive As Boolean = False
    Private ReadOnly _rand As New Random()

    ' Package data structure
    Private Class PackageInfo
        Public Property Name As String
        Public Property Id As String
        Public Property Version As String
        Public Property Match As String
        Public Property Source As String
        Public Property FullLine As String

        Public Overrides Function ToString() As String
            Return Name
        End Function

        Public Function GetFormattedDisplay() As String
            Dim namePart = If(String.IsNullOrWhiteSpace(Name), "", Name.PadRight(40))
            Dim idPart = If(String.IsNullOrWhiteSpace(Id), "", Id.PadRight(40))
            Dim versionPart = If(String.IsNullOrWhiteSpace(Version), "", Version.PadRight(16))
            Dim matchPart = If(String.IsNullOrWhiteSpace(Match), "", Match.PadRight(14))
            Dim sourcePart = If(String.IsNullOrWhiteSpace(Source), "", Source)

            Return $"{namePart} {idPart} {versionPart} {matchPart} {sourcePart}".TrimEnd()
        End Function
    End Class

    Private _selectedPackages As New List(Of PackageInfo)()
    Private _allSearchResults As New List(Of PackageInfo)() ' Store all unfiltered results
    Private Const MsStoreSourceIdentifier As String = "msstore"
    Private _moreDetailMenuItem As MenuItem

    Private Async Sub MainWindow_Loaded(sender As Object, e As RoutedEventArgs) Handles Me.Loaded
        EnableDarkTitleBar()
        ApplySystemTheme()
        InitWingetIndicators()
        Await SetBingWallpaperAsync()
        SetupInitialVisibility()
        WireMenuEvents()
        InitializeSourceFilter()
        SetupSearchResultsContextMenu()
        SetInitialFocus()
    End Sub

    Private Sub SetInitialFocus()
        Dispatcher.BeginInvoke(New Action(Sub()
                                              SearchTextBox.Focus()
                                              Keyboard.Focus(SearchTextBox)
                                          End Sub), DispatcherPriority.ContextIdle)
    End Sub

    Private Sub WireMenuEvents()
        AddHandler NewMenuItem.Click, AddressOf NewMenuItem_Click
        AddHandler OpenMenuItem.Click, AddressOf OpenMenuItem_Click
        AddHandler GetCurrentListMenuItem.Click, AddressOf GetCurrentListMenuItem_Click
        AddHandler CaptureStartMenuMenuItem.Click, AddressOf CaptureStartMenuMenuItem_Click
    End Sub

    Private Sub SetupInitialVisibility()
        SearchResultsListBox.Visibility = Visibility.Collapsed
        SelectedNamesListBox.Visibility = Visibility.Collapsed
        RemoveButton.Visibility = Visibility.Collapsed
        FilenameTextBox.IsEnabled = False
        BuildButton.IsEnabled = False
        DeleteLinksCheckBox.IsEnabled = False
        CaptureStartMenuCheckBox.IsEnabled = False
        SourceFilterPanel.Visibility = Visibility.Collapsed
        SelectedSearchTextBox.IsEnabled = False
        SelectedSearchTextBox.Text = ""
        Dim storeButton = GetStoreLinkButton()
        If storeButton IsNot Nothing Then
            storeButton.Visibility = Visibility.Collapsed
            storeButton.Tag = Nothing
        End If
    End Sub

    ' ==================== THEME METHODS ====================

    Private Sub ApplySystemTheme()
        ' Detect if Windows is using dark mode
        Dim isDarkMode = IsWindowsInDarkMode()

        If isDarkMode Then
            ' Apply dark theme to main window only
            Me.Background = New SolidColorBrush(Color.FromRgb(&H20, &H20, &H20))
        Else
            ' Apply light theme to main window only (transparent to show wallpaper)
            Me.Background = Brushes.Transparent
        End If
    End Sub

    Private Function IsWindowsInDarkMode() As Boolean
        Try
            Using key = Registry.CurrentUser.OpenSubKey("Software\Microsoft\Windows\CurrentVersion\Themes\Personalize")
                If key IsNot Nothing Then
                    Dim value = key.GetValue("AppsUseLightTheme")
                    If value IsNot Nothing Then
                        ' 0 = Dark Mode, 1 = Light Mode
                        Return CInt(value) = 0
                    End If
                End If
            End Using
        Catch ex As Exception
            Debug.WriteLine($"Error detecting theme: {ex.Message}")
        End Try
        Return False ' Default to light mode if detection fails
    End Function

    ' ==================== WINGET INDICATOR METHODS ====================

    Private Sub InitWingetIndicators()
        If _wingetBlinkTimer Is Nothing Then
            _wingetBlinkTimer = New DispatcherTimer(TimeSpan.FromMilliseconds(120),
                                                    DispatcherPriority.Background,
                                                    AddressOf WingetBlinkTick,
                                                    Dispatcher.CurrentDispatcher)
        End If
        UpdateWingetLightOff()
    End Sub

    Private Sub WingetBlinkTick(sender As Object, e As EventArgs)
        If Not _wingetActive Then Return
        Dim sample = _rand.NextDouble()
        If sample < 0.55 Then
            WingetActivityLight.Fill = New SolidColorBrush(Color.FromRgb(0, 220, 120))
            WingetActivityLight.Effect = New DropShadowEffect With {
                .Color = Colors.Lime,
                .BlurRadius = 8,
                .ShadowDepth = 0,
                .Opacity = 0.75
            }
        ElseIf sample < 0.8 Then
            UpdateWingetLightDim()
        Else
            UpdateWingetLightOff()
        End If
    End Sub

    Private Sub StartWingetIndicator()
        _wingetActive = True
        If _wingetBlinkTimer IsNot Nothing AndAlso Not _wingetBlinkTimer.IsEnabled Then
            _wingetBlinkTimer.Start()
        End If
    End Sub

    Private Sub StopWingetIndicator()
        _wingetActive = False
        _wingetBlinkTimer?.Stop()
        UpdateWingetLightOff()
    End Sub

    Private Sub UpdateWingetLightOff()
        If WingetActivityLight IsNot Nothing Then
            WingetActivityLight.Fill = New SolidColorBrush(Color.FromRgb(50, 50, 50))
            WingetActivityLight.Effect = Nothing
        End If
    End Sub

    Private Sub UpdateWingetLightDim()
        If WingetActivityLight IsNot Nothing Then
            WingetActivityLight.Fill = New SolidColorBrush(Color.FromRgb(25, 80, 55))
            WingetActivityLight.Effect = Nothing
        End If
    End Sub

    ' ==================== PROGRESS METHODS ====================

    Private Sub UpdateProgress(mainText As String, detailText As String)
        If ProgressText IsNot Nothing Then
            ProgressText.Text = mainText
            ProgressText.Visibility = Visibility.Visible
        End If
        If ProgressTextDetail IsNot Nothing Then
            ProgressTextDetail.Text = detailText
            ProgressTextDetail.Visibility = Visibility.Visible
        End If
    End Sub

    Private Sub UpdateProgressSuccess(mainText As String, detailText As String)
        If ProgressText IsNot Nothing Then
            ProgressText.Text = mainText
            ProgressText.Foreground = New SolidColorBrush(Colors.LimeGreen)
            ProgressText.Visibility = Visibility.Visible
        End If
        If ProgressTextDetail IsNot Nothing Then
            ProgressTextDetail.Text = detailText
            ProgressTextDetail.Visibility = Visibility.Visible
        End If
    End Sub

    Private Sub UpdateProgressError(mainText As String, detailText As String)
        If ProgressText IsNot Nothing Then
            ProgressText.Text = mainText
            ProgressText.Foreground = New SolidColorBrush(Colors.Red)
            ProgressText.Visibility = Visibility.Visible
        End If
        If ProgressTextDetail IsNot Nothing Then
            ProgressTextDetail.Text = detailText
            ProgressTextDetail.Visibility = Visibility.Visible
        End If
    End Sub

    Private Async Sub ResetProgress()
        Await Task.Delay(2000)
        If ProgressText IsNot Nothing Then
            ProgressText.Text = ""
            ProgressText.Foreground = New SolidColorBrush(Color.FromRgb(&HE2, &HE2, &HE2))
            ProgressText.Visibility = Visibility.Collapsed
        End If
        If ProgressTextDetail IsNot Nothing Then
            ProgressTextDetail.Text = ""
            ProgressTextDetail.Visibility = Visibility.Collapsed
        End If
    End Sub

    ' ==================== FILTER FUNCTIONALITY ====================

    Private Sub InitializeSourceFilter()
        ' Wire up button click to toggle popup
        AddHandler SourceFilterButton.Click, Sub(s, e)
                                                 SourceFilterPopup.IsOpen = Not SourceFilterPopup.IsOpen
                                             End Sub
    End Sub

    ' ==================== SEARCH RESULTS CONTEXT MENU ====================

    Private Sub SetupSearchResultsContextMenu()
        If SearchResultsListBox Is Nothing Then Return

        _moreDetailMenuItem = New MenuItem With {
            .Header = "More detail",
            .IsEnabled = False
        }
        AddHandler _moreDetailMenuItem.Click, AddressOf SearchResultMoreDetail_Click

        Dim contextMenu As New ContextMenu()
        contextMenu.Items.Add(_moreDetailMenuItem)
        SearchResultsListBox.ContextMenu = contextMenu
    End Sub

    Private Sub SearchResultsListBox_ContextMenuOpening(sender As Object, e As ContextMenuEventArgs) Handles SearchResultsListBox.ContextMenuOpening
        If _moreDetailMenuItem Is Nothing Then Return

        Dim package = GetPackageFromListBoxItem(SearchResultsListBox.SelectedItem)
        Dim enabled = PackageSupportsStoreDetail(package)

        _moreDetailMenuItem.IsEnabled = enabled
        _moreDetailMenuItem.Tag = If(enabled, package, Nothing)
    End Sub

    Private Sub SearchResultsListBox_PreviewMouseRightButtonDown(sender As Object, e As MouseButtonEventArgs) Handles SearchResultsListBox.PreviewMouseRightButtonDown
        Dim origin = TryCast(e.OriginalSource, DependencyObject)
        If origin Is Nothing Then Return

        Dim item = FindAncestor(Of ListBoxItem)(origin)
        If item IsNot Nothing Then
            item.IsSelected = True
        Else
            SearchResultsListBox.SelectedIndex = -1
        End If
    End Sub

    Private Sub SelectedNamesListBox_PreviewMouseRightButtonDown(sender As Object, e As MouseButtonEventArgs) Handles SelectedNamesListBox.PreviewMouseRightButtonDown
        Dim origin = TryCast(e.OriginalSource, DependencyObject)
        If origin Is Nothing Then Return

        Dim item = FindAncestor(Of ListBoxItem)(origin)
        If item IsNot Nothing Then
            item.IsSelected = True
        Else
            SelectedNamesListBox.SelectedIndex = -1
        End If
    End Sub

    Private Shared Function PackageSupportsStoreDetail(package As PackageInfo) As Boolean
        Return package IsNot Nothing AndAlso
               Not String.IsNullOrWhiteSpace(package.Id) AndAlso
               Not String.IsNullOrWhiteSpace(package.Source) AndAlso
               package.Source.IndexOf(MsStoreSourceIdentifier, StringComparison.OrdinalIgnoreCase) >= 0
    End Function

    Private Sub SearchResultMoreDetail_Click(sender As Object, e As RoutedEventArgs)
        Dim menuItem = TryCast(sender, MenuItem)
        Dim package = TryCast(menuItem?.Tag, PackageInfo)
        LaunchStorePage(package)
    End Sub

    Private Sub StoreLinkButton_Click(sender As Object, e As RoutedEventArgs)
        Dim button = TryCast(sender, Button)
        Dim package = TryCast(button?.Tag, PackageInfo)
        LaunchStorePage(package)
    End Sub

    Private Sub LaunchStorePage(package As PackageInfo)
        If package Is Nothing OrElse String.IsNullOrWhiteSpace(package.Id) Then Return

        Dim targetUrl = $"https://apps.microsoft.com/detail/{package.Id}"

        Try
            Process.Start(New ProcessStartInfo(targetUrl) With {.UseShellExecute = True})
        Catch ex As Exception
            UpdateProgressError("Launch Failed", "Unable to open Microsoft Store page")
            Debug.WriteLine($"Store launch error: {ex.Message}")
            ResetProgress()
        End Try
    End Sub

    Private Shared Function FindAncestor(Of T As DependencyObject)(current As DependencyObject) As T
        While current IsNot Nothing
            Dim candidate = TryCast(current, T)
            If candidate IsNot Nothing Then
                Return candidate
            End If

            Dim visual = TryCast(current, Visual)
            If visual IsNot Nothing Then
                current = VisualTreeHelper.GetParent(visual)
                Continue While
            End If

            Dim contentElement = TryCast(current, FrameworkContentElement)
            If contentElement IsNot Nothing Then
                current = TryCast(contentElement.Parent, DependencyObject)
                Continue While
            End If

            current = Nothing
        End While

        Return Nothing
    End Function

    Private Function GetSelectedPackageFromSelectedList() As PackageInfo
        Dim package = GetPackageFromListBoxItem(SelectedNamesListBox.SelectedItem)
        If package IsNot Nothing Then
            Return package
        End If

        Dim selectedIndex = SelectedNamesListBox.SelectedIndex
        If selectedIndex >= 0 AndAlso selectedIndex < _selectedPackages.Count Then
            Return _selectedPackages(selectedIndex)
        End If

        Return Nothing
    End Function

    Private Sub AddPackageToSelectedDisplay(package As PackageInfo)
        Dim listItem As New ListBoxItem With {
            .Content = package.GetFormattedDisplay(),
            .Tag = package
        }
        SelectedNamesListBox.Items.Add(listItem)
        SelectedIDsListBox.Items.Add(package.Id)
    End Sub

    Private Function GetSelectedListFilter() As String
        If SelectedSearchTextBox Is Nothing Then
            Return Nothing
        End If

        Dim current = SelectedSearchTextBox.Text
        If String.IsNullOrWhiteSpace(current) Then
            Return Nothing
        End If

        Return current.Trim()
    End Function

    Private Sub RefreshSelectedPackagesList(Optional filterTerm As String = Nothing, Optional showFeedback As Boolean = False)
        If SelectedNamesListBox Is Nothing OrElse SelectedIDsListBox Is Nothing Then
            Return
        End If

        SelectedNamesListBox.Items.Clear()
        SelectedIDsListBox.Items.Clear()

        Dim comparison = StringComparison.OrdinalIgnoreCase
        Dim hasFilter = Not String.IsNullOrWhiteSpace(filterTerm)
        Dim matchCount = 0

        For Each package In _selectedPackages
            If Not hasFilter OrElse PackageMatchesSearch(package, filterTerm, comparison) Then
                AddPackageToSelectedDisplay(package)
                matchCount += 1
            End If
        Next

        If showFeedback AndAlso hasFilter Then
            If matchCount > 0 Then
                UpdateProgressSuccess("Filtered", $"{matchCount} matching package(s)")
            Else
                UpdateProgressError("Not Found", $"No selected package matching '{filterTerm}'")
            End If
            ResetProgress()
        End If
    End Sub

    Private Sub SelectedSearchTextBox_KeyDown(sender As Object, e As Input.KeyEventArgs) Handles SelectedSearchTextBox.KeyDown
        If e.Key = Input.Key.Enter Then
            e.Handled = True
            SearchSelectedPackages(SelectedSearchTextBox.Text)
        ElseIf e.Key = Input.Key.Escape Then
            SelectedSearchTextBox.Clear()
            SelectedNamesListBox.SelectedIndex = -1
            e.Handled = True
        End If
    End Sub

    Private Sub SelectedSearchTextBox_TextChanged(sender As Object, e As TextChangedEventArgs) Handles SelectedSearchTextBox.TextChanged
        If SelectedSearchTextBox Is Nothing OrElse Not SelectedSearchTextBox.IsEnabled Then
            Return
        End If

        RefreshSelectedPackagesList(GetSelectedListFilter())
    End Sub

    Private Sub SearchSelectedPackages(searchTerm As String)
        Dim query = searchTerm?.Trim()
        If String.IsNullOrEmpty(query) Then
            UpdateProgressError("Search Required", "Enter text to search selected packages")
            ResetProgress()
            Return
        End If

        RefreshSelectedPackagesList(query, showFeedback:=True)

        If SelectedNamesListBox.Items.Count > 0 Then
            SelectedNamesListBox.SelectedIndex = 0
            SelectedNamesListBox.ScrollIntoView(SelectedNamesListBox.SelectedItem)
        End If
    End Sub

    Private Shared Function PackageMatchesSearch(package As PackageInfo, term As String, comparison As StringComparison) As Boolean
        Return (Not String.IsNullOrWhiteSpace(package.Name) AndAlso package.Name.IndexOf(term, comparison) >= 0) OrElse
               (Not String.IsNullOrWhiteSpace(package.Id) AndAlso package.Id.IndexOf(term, comparison) >= 0) OrElse
               (Not String.IsNullOrWhiteSpace(package.Match) AndAlso package.Match.IndexOf(term, comparison) >= 0) OrElse
               (Not String.IsNullOrWhiteSpace(package.Source) AndAlso package.Source.IndexOf(term, comparison) >= 0)
    End Function

    Private Sub PopulateSourceFilter()
        ' Get unique source values from search results
        Dim uniqueSources = _allSearchResults.
            Where(Function(p) Not String.IsNullOrWhiteSpace(p.Source)).
            Select(Function(p) p.Source).
            Distinct().
            OrderBy(Function(s) s).
            ToList()

        ' Check if there are packages with blank/null sources
        Dim hasBlankSources = _allSearchResults.Any(Function(p) String.IsNullOrWhiteSpace(p.Source))

        ' Clear existing menu items
        SourceFilterMenu.Children.Clear()

        ' Add "All Sources" option
        Dim allSourcesItem As New MenuItem With {
            .Header = "All Sources",
            .Style = TryCast(FindResource("FilterMenuItemStyle"), Style),
            .IsCheckable = True,
            .IsChecked = True
        }
        AddHandler allSourcesItem.Click, AddressOf FilterMenuItem_Click
        SourceFilterMenu.Children.Add(allSourcesItem)

        ' Add separator
        If uniqueSources.Count > 0 OrElse hasBlankSources Then
            SourceFilterMenu.Children.Add(New Separator With {.Background = New SolidColorBrush(Color.FromRgb(&H3F, &H3F, &H3F))})
        End If

        ' Add "(No Source)" option if there are packages without sources
        If hasBlankSources Then
            Dim noSourceItem As New MenuItem With {
                .Header = "(No Source)",
                .Style = TryCast(FindResource("FilterMenuItemStyle"), Style),
                .IsCheckable = True
            }
            AddHandler noSourceItem.Click, AddressOf FilterMenuItem_Click
            SourceFilterMenu.Children.Add(noSourceItem)
        End If

        ' Add each unique source
        For Each source In uniqueSources
            Dim menuItem As New MenuItem With {
                .Header = source,
                .Style = TryCast(FindResource("FilterMenuItemStyle"), Style),
                .IsCheckable = True
            }
            AddHandler menuItem.Click, AddressOf FilterMenuItem_Click
            SourceFilterMenu.Children.Add(menuItem)
        Next

        ' Show the filter panel if we have results
        If _allSearchResults.Count > 0 Then
            SourceFilterPanel.Visibility = Visibility.Visible
        Else
            SourceFilterPanel.Visibility = Visibility.Collapsed
        End If
    End Sub

    ' New method to handle filter menu item clicks
    Private Sub FilterMenuItem_Click(sender As Object, e As RoutedEventArgs)
        Dim clickedItem = TryCast(sender, MenuItem)
        If clickedItem Is Nothing Then Return

        ' Uncheck all other items
        For Each child In SourceFilterMenu.Children
            Dim menuItem = TryCast(child, MenuItem)
            If menuItem IsNot Nothing AndAlso menuItem IsNot clickedItem Then
                menuItem.IsChecked = False
            End If
        Next

        ' Ensure clicked item is checked
        clickedItem.IsChecked = True

        ' Update button text
        Dim selectedText = clickedItem.Header.ToString()
        Dim filterText = SourceFilterButton.Template.FindName("FilterText", SourceFilterButton)
        If TypeOf filterText Is TextBlock Then
            DirectCast(filterText, TextBlock).Text = If(selectedText = "All Sources", "Filter Source", selectedText)
        End If

        ' Close popup
        SourceFilterPopup.IsOpen = False

        ' Apply the filter
        ApplySourceFilter(selectedText)
    End Sub

    ' Update ApplySourceFilter to accept selected filter parameter
    Private Sub ApplySourceFilter(Optional selectedFilter As String = Nothing)
        If _allSearchResults.Count = 0 Then Return

        ' If no filter specified, get it from checked menu item
        If String.IsNullOrEmpty(selectedFilter) Then
            For Each child In SourceFilterMenu.Children
                Dim menuItem = TryCast(child, MenuItem)
                If menuItem IsNot Nothing AndAlso menuItem.IsChecked Then
                    selectedFilter = menuItem.Header.ToString()
                    Exit For
                End If
            Next
        End If

        If String.IsNullOrEmpty(selectedFilter) Then selectedFilter = "All Sources"

        SearchResultsListBox.Items.Clear()

        Dim filteredResults As List(Of PackageInfo)

        If selectedFilter = "All Sources" Then
            ' Show all results
            filteredResults = _allSearchResults
        ElseIf selectedFilter = "(No Source)" Then
            ' Filter to show only packages with no source
            filteredResults = _allSearchResults.Where(Function(p) String.IsNullOrWhiteSpace(p.Source)).ToList()
        Else
            ' Filter to show only packages from selected source (check for null Source)
            filteredResults = _allSearchResults.Where(Function(p) p.Source IsNot Nothing AndAlso p.Source.Equals(selectedFilter, StringComparison.OrdinalIgnoreCase)).ToList()
        End If

        ' Populate the listbox with filtered results
        For Each package In filteredResults
            Dim displayLine = If(String.IsNullOrWhiteSpace(package.FullLine), package.GetFormattedDisplay(), package.FullLine)
            Dim listItem As New ListBoxItem With {
                .Content = displayLine,
                .Tag = package
            }
            SearchResultsListBox.Items.Add(listItem)
        Next

        ' Update progress message
        Dim totalCount = _allSearchResults.Count
        Dim displayCount = filteredResults.Count

        If selectedFilter = "All Sources" Then
            UpdateProgressSuccess("Showing All", $"{displayCount} package(s)")
        Else
            UpdateProgressSuccess("Filtered", $"{displayCount} of {totalCount} package(s) - {selectedFilter}")
        End If

        UpdateStoreLinkButton()
        ResetProgress()
    End Sub

    ' ==================== SEARCH FUNCTIONALITY ====================

    Private Async Sub SearchButton_Click(sender As Object, e As RoutedEventArgs) Handles SearchButton.Click
        Await PerformSearchAsync()
    End Sub

    Private Async Sub SearchTextBox_KeyDown(sender As Object, e As Input.KeyEventArgs) Handles SearchTextBox.KeyDown
        If e.Key = Input.Key.Enter Then
            Await PerformSearchAsync()
            e.Handled = True
        ElseIf e.Key = Input.Key.Escape Then
            SearchTextBox.Text = ""
            SearchResultsListBox.Items.Clear()
            _allSearchResults.Clear()
            SearchResultsListBox.Visibility = Visibility.Collapsed
            SourceFilterPanel.Visibility = Visibility.Collapsed
            e.Handled = True
        End If
    End Sub

    Private Sub SearchTextBox_GotFocus(sender As Object, e As RoutedEventArgs)
        If SearchTextBox.Text = "Search packages..." Then
            SearchTextBox.Text = ""
        End If
    End Sub

    Private Async Function PerformSearchAsync() As Task
        If String.IsNullOrWhiteSpace(SearchTextBox.Text) OrElse SearchTextBox.Text = "Search packages..." Then
            Return
        End If

        UpdateProgress("Searching...", $"Looking for '{SearchTextBox.Text}'")
        StartWingetIndicator()

        SearchResultsListBox.Items.Clear()
        _allSearchResults.Clear()

        Try
            _allSearchResults = Await RunWingetSearchAsync(SearchTextBox.Text)

            If _allSearchResults.Count = 0 Then
                UpdateProgressError("No Results", "No packages found matching criteria")
                SearchResultsListBox.Visibility = Visibility.Collapsed
                SourceFilterPanel.Visibility = Visibility.Collapsed
            Else
                ' Populate the source filter dropdown
                PopulateSourceFilter()

                ' Apply the current filter
                ApplySourceFilter()

                SearchResultsListBox.Visibility = Visibility.Visible
            End If
        Catch ex As Exception
            UpdateProgressError("Search Failed", ex.Message)
        Finally
            StopWingetIndicator()
            If _allSearchResults.Count = 0 Then
                ResetProgress()
            End If
        End Try
    End Function

    Private Async Function RunWingetSearchAsync(searchTerm As String) As Task(Of List(Of PackageInfo))
        Return Await Task.Run(Function() As List(Of PackageInfo)
                                  Dim packages As New List(Of PackageInfo)

                                  Try
                                      ' Create worker files
                                      Dim workerFile = "WorkerFile.cmd"
                                      Dim outputFile = "WorkerOutput.txt"

                                      ' Clean up old files
                                      If File.Exists(workerFile) Then File.Delete(workerFile)
                                      If File.Exists(outputFile) Then File.Delete(outputFile)

                                      ' Create worker batch file
                                      Using writer As New StreamWriter(workerFile)
                                          writer.WriteLine("@Echo Off")
                                          writer.WriteLine("CLS")
                                          writer.WriteLine($"WINGET Search ""{searchTerm}"" > {outputFile}")
                                      End Using

                                      ' Execute winget search
                                      Dim psi As New ProcessStartInfo With {
                                          .FileName = "cmd.exe",
                                          .Arguments = $"/c {workerFile}",
                                          .UseShellExecute = False,
                                          .CreateNoWindow = True,
                                          .WindowStyle = ProcessWindowStyle.Hidden
                                      }

                                      Using proc = Process.Start(psi)
                                          proc.WaitForExit()
                                      End Using

                                      ' Parse results
                                      If File.Exists(outputFile) Then
                                          Dim lines = File.ReadAllLines(outputFile).ToList()

                                          ' Find header line
                                          Dim headerIndex = -1
                                          For i = 0 To lines.Count - 1
                                              If lines(i).Contains("---") Then
                                                  headerIndex = i - 1
                                                  Exit For
                                              End If
                                          Next

                                          If headerIndex >= 0 AndAlso headerIndex < lines.Count Then
                                              _headerValue = lines(headerIndex)

                                              ' Parse data lines (skip header and separator)
                                              For i = headerIndex + 2 To lines.Count - 1
                                                  If Not String.IsNullOrWhiteSpace(lines(i)) Then
                                                      packages.Add(ParsePackageLine(lines(i)))
                                                  End If
                                              Next
                                          End If
                                      End If

                                      ' Clean up
                                      If File.Exists(workerFile) Then File.Delete(workerFile)
                                      If File.Exists(outputFile) Then File.Delete(outputFile)

                                  Catch ex As Exception
                                      Debug.WriteLine($"Search error: {ex.Message}")
                                  End Try

                                  Return packages
                              End Function)
    End Function

    Private Function ParsePackageLine(line As String) As PackageInfo
        Dim package As New PackageInfo With {.FullLine = line}
        Try
            If Not String.IsNullOrEmpty(_headerValue) Then
                Dim nameIndex = _headerValue.IndexOf("Name")
                Dim idIndex = _headerValue.IndexOf("Id")
                Dim versionIndex = _headerValue.IndexOf("Version")
                Dim matchIndex = _headerValue.IndexOf("Match")
                Dim sourceIndex = _headerValue.IndexOf("Source")

                ' Extract fields with better boundary checking and trimming
                If nameIndex >= 0 AndAlso idIndex > nameIndex Then
                    Dim endPos = Math.Min(idIndex, line.Length)
                    If nameIndex < line.Length Then
                        package.Name = line.Substring(nameIndex, endPos - nameIndex).Trim()
                    End If
                End If

                If idIndex >= 0 AndAlso versionIndex > idIndex Then
                    Dim endPos = Math.Min(versionIndex, line.Length)
                    If idIndex < line.Length Then
                        package.Id = line.Substring(idIndex, endPos - idIndex).Trim()
                    End If
                End If

                If versionIndex >= 0 AndAlso matchIndex > versionIndex Then
                    Dim endPos = Math.Min(matchIndex, line.Length)
                    If versionIndex < line.Length Then
                        package.Version = line.Substring(versionIndex, endPos - versionIndex).Trim()
                    End If
                End If

                If matchIndex >= 0 AndAlso sourceIndex > matchIndex Then
                    Dim endPos = Math.Min(sourceIndex, line.Length)
                    If matchIndex < line.Length Then
                        package.Match = line.Substring(matchIndex, endPos - matchIndex).Trim()
                    End If
                End If

                If sourceIndex >= 0 AndAlso line.Length > sourceIndex Then
                    Dim sourceValue = line.Substring(sourceIndex).Trim()
                    ' Only set Source if the value is not empty or whitespace
                    If Not String.IsNullOrWhiteSpace(sourceValue) Then
                        package.Source = sourceValue
                    End If
                End If
            End If
        Catch ex As Exception
            Debug.WriteLine($"Parse error: {ex.Message} - Line: {line}")
        End Try
        Return package
    End Function

    Private Function GetPackageFromListBoxItem(item As Object) As PackageInfo
        If item Is Nothing Then Return Nothing

        If TypeOf item Is ListBoxItem Then
            Return TryCast(DirectCast(item, ListBoxItem).Tag, PackageInfo)
        End If

        Dim itemValue = TryCast(item, String)
        If Not String.IsNullOrEmpty(itemValue) Then
            Return ParsePackageLine(itemValue)
        End If

        Return Nothing
    End Function

    ' ==================== ADD/REMOVE FUNCTIONALITY ====================

    Private Sub SearchResultsListBox_MouseDoubleClick(sender As Object, e As Input.MouseButtonEventArgs) Handles SearchResultsListBox.MouseDoubleClick
        AddSelectedPackage()
    End Sub

    Private Sub SearchResultsListBox_KeyDown(sender As Object, e As Input.KeyEventArgs) Handles SearchResultsListBox.KeyDown
        If e.Key = Input.Key.Enter Then
            AddSelectedPackage()
            e.Handled = True
        End If
    End Sub

    Private Sub AddSelectedPackage()
        If SearchResultsListBox.SelectedItem Is Nothing Then Return

        Dim package = GetPackageFromListBoxItem(SearchResultsListBox.SelectedItem)

        If package IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(package.Id) Then
            ' Check for duplicate package name
            Dim existingPackage = _selectedPackages.FirstOrDefault(Function(p) p.Name.Equals(package.Name, StringComparison.OrdinalIgnoreCase))

            If existingPackage IsNot Nothing Then
                ' Duplicate found - warn user
                Dim result = MessageBox.Show(
                    $"A package with the name '{package.Name}' is already in the selected list.{Environment.NewLine}{Environment.NewLine}" &
                    $"Existing: {existingPackage.Id}{Environment.NewLine}" &
                    $"New: {package.Id}{Environment.NewLine}{Environment.NewLine}" &
                    "Do you want to add it anyway?",
                    "Duplicate Package Name",
                    MessageBoxButton.YesNo,
                    MessageBoxImage.Warning)

                If result = MessageBoxResult.No Then
                    UpdateProgressError("Cancelled", "Duplicate package not added")
                    ResetProgress()
                    Return
                End If
            End If

            ' Add the package
            _selectedPackages.Add(package)
            RefreshSelectedPackagesList(GetSelectedListFilter())

            ' Enable build controls
            SelectedNamesListBox.Visibility = Visibility.Visible
            RemoveButton.Visibility = Visibility.Visible
            FilenameTextBox.IsEnabled = True
            BuildButton.IsEnabled = True
            DeleteLinksCheckBox.IsEnabled = True
            CaptureStartMenuCheckBox.IsEnabled = True
            SelectedSearchTextBox.IsEnabled = True

            UpdateProgressSuccess("Added", $"{package.Name}")
            ResetProgress()
        End If
    End Sub

    Private Sub RemoveButton_Click(sender As Object, e As RoutedEventArgs) Handles RemoveButton.Click
        RemoveSelectedPackage()
    End Sub

    Private Sub SelectedNamesListBox_KeyDown(sender As Object, e As Input.KeyEventArgs) Handles SelectedNamesListBox.KeyDown
        If e.Key = Input.Key.Delete Then
            RemoveSelectedPackage()
            e.Handled = True
        End If
    End Sub

    Private Sub RemoveSelectedPackage()
        Dim selectedItem = TryCast(SelectedNamesListBox.SelectedItem, ListBoxItem)
        If selectedItem Is Nothing Then Return

        Dim package = TryCast(selectedItem.Tag, PackageInfo)
        If package Is Nothing Then Return

        If _selectedPackages.Remove(package) Then
            RefreshSelectedPackagesList(GetSelectedListFilter())
        End If

        If _selectedPackages.Count = 0 Then
            SelectedNamesListBox.Visibility = Visibility.Collapsed
            RemoveButton.Visibility = Visibility.Collapsed
            FilenameTextBox.IsEnabled = False
            BuildButton.IsEnabled = False
            DeleteLinksCheckBox.IsEnabled = False
            CaptureStartMenuCheckBox.IsEnabled = False
            SelectedSearchTextBox.IsEnabled = False
            SelectedSearchTextBox.Text = ""
        End If
    End Sub

    ' ==================== BUILD FUNCTIONALITY ====================

    Private Async Sub BuildButton_Click(sender As Object, e As RoutedEventArgs) Handles BuildButton.Click
        If String.IsNullOrWhiteSpace(FilenameTextBox.Text) Then
            MessageBox.Show("Please enter a filename for the script.", "No Filename",
                          MessageBoxButton.OK, MessageBoxImage.Warning)
            Return
        End If

        UpdateProgress("Building...", "Creating installation script")
        StartWingetIndicator()

        Try
            Await BuildScriptAsync()
            UpdateProgressSuccess("Complete", $"{FilenameTextBox.Text}.cmd created")

            ' Open folder
            Process.Start("explorer.exe", Directory.GetCurrentDirectory())

            ' Clear form
            ClearAll()
        Catch ex As Exception
            UpdateProgressError("Build Failed", ex.Message)
        Finally
            StopWingetIndicator()
            ResetProgress()
        End Try
    End Sub

    Private Async Function BuildScriptAsync() As Task
        ' Capture all UI values on the UI thread before starting the background task
        Dim filename = FilenameTextBox.Text
        Dim deleteLinks = DeleteLinksCheckBox.IsChecked
        Dim captureStartMenu = CaptureStartMenuCheckBox.IsChecked
        Dim selectedPackages = _selectedPackages.ToList()

        Await Task.Run(Sub()
                           Dim scriptFile = $"{filename}.cmd"
                           Dim logFile = $"{filename}_install.log"
                           Dim configFile = $"{filename}.json"

                           ' Delete existing files
                           For Each file In {scriptFile, logFile, configFile}
                               If System.IO.File.Exists(file) Then System.IO.File.Delete(file)
                           Next

                           ' Build enhanced script with better structure
                           Using writer As New StreamWriter(scriptFile)
                               ' Header
                               writer.WriteLine("@ECHO OFF")
                               writer.WriteLine("SETLOCAL EnableDelayedExpansion")
                               writer.WriteLine("TITLE Winget Installation Script - {0}", filename)
                               writer.WriteLine("COLOR 0F")
                               writer.WriteLine("CLS")
                               writer.WriteLine()

                               ' Initialize variables
                               writer.WriteLine("REM ========================================")
                               writer.WriteLine("REM Initialize Installation Environment")
                               writer.WriteLine("REM ========================================")
                               writer.WriteLine("SET ""SCRIPT_NAME={0}""", filename)
                               writer.WriteLine("SET ""LOG_FILE={0}""", logFile)
                               writer.WriteLine("SET ""INSTALL_COUNT=0""")
                               writer.WriteLine("SET ""SUCCESS_COUNT=0""")
                               writer.WriteLine("SET ""FAILED_COUNT=0""")
                               writer.WriteLine("SET ""SKIPPED_COUNT=0""")
                               writer.WriteLine("SET ""TOTAL_PACKAGES={0}""", selectedPackages.Count)
                               writer.WriteLine("SET ""START_TIME=%TIME%""")
                               writer.WriteLine()

                               ' Create log header
                               writer.WriteLine("ECHO ======================================== > ""%LOG_FILE%""")
                               writer.WriteLine("ECHO Winget Installation Log >> ""%LOG_FILE%""")
                               writer.WriteLine("ECHO Script: %SCRIPT_NAME% >> ""%LOG_FILE%""")
                               writer.WriteLine("ECHO Date: %DATE% >> ""%LOG_FILE%""")
                               writer.WriteLine("ECHO Start Time: %START_TIME% >> ""%LOG_FILE%""")
                               writer.WriteLine("ECHO Total Packages: %TOTAL_PACKAGES% >> ""%LOG_FILE%""")
                               writer.WriteLine("ECHO ======================================== >> ""%LOG_FILE%""")
                               writer.WriteLine("ECHO. >> ""%LOG_FILE%""")
                               writer.WriteLine()

                               ' Welcome message
                               writer.WriteLine("ECHO.")
                               writer.WriteLine("ECHO ========================================")
                               writer.WriteLine("ECHO     Winget Package Installation")
                               writer.WriteLine("ECHO ========================================")
                               writer.WriteLine("ECHO.")
                               writer.WriteLine("ECHO This script will install {0} package(s)", selectedPackages.Count)
                               writer.WriteLine("ECHO Log file: %LOG_FILE%")
                               writer.WriteLine("ECHO.")
                               writer.WriteLine("PAUSE")
                               writer.WriteLine("CLS")
                               writer.WriteLine()

                               ' Install each package
                               For i = 0 To selectedPackages.Count - 1
                                   Dim package = selectedPackages(i)
                                   Dim pkgNum = i + 1

                                   writer.WriteLine("REM ========================================")
                                   writer.WriteLine("REM Package {0}/{1}: {2}", pkgNum, selectedPackages.Count, package.Name)
                                   writer.WriteLine("REM ========================================")
                                   writer.WriteLine("COLOR 0F")
                                   writer.WriteLine("CLS")
                                   writer.WriteLine("ECHO.")
                                   writer.WriteLine("ECHO ========================================")
                                   writer.WriteLine("ECHO Progress: {0}/{1} packages", pkgNum, selectedPackages.Count)
                                   writer.WriteLine("ECHO ========================================")
                                   writer.WriteLine("ECHO.")
                                   writer.WriteLine("ECHO Package: {0}", package.Name)
                                   writer.WriteLine("ECHO ID: {0}", package.Id)
                                   If Not String.IsNullOrEmpty(package.Version) Then
                                       writer.WriteLine("ECHO Version: {0}", package.Version)
                                   Else
                                       writer.WriteLine("ECHO Version: Unknown")
                                   End If
                                   writer.WriteLine("ECHO.")
                                   writer.WriteLine()

                                   ' Prompt for installation (with variable reset)
                                   writer.WriteLine("SET ""INSTALL_CHOICE=""")
                                   writer.WriteLine("SET /P INSTALL_CHOICE=""Install this package? [Y/N] (Default=Y): """)
                                   writer.WriteLine("IF /I ""%INSTALL_CHOICE%""=="""" SET ""INSTALL_CHOICE=Y""")
                                   writer.WriteLine()

                                   writer.WriteLine("IF /I ""%INSTALL_CHOICE%""==""Y"" (")
                                   writer.WriteLine("    SET /A INSTALL_COUNT+=1")
                                   writer.WriteLine("    ECHO [%TIME%] Installing: {0} >> ""%LOG_FILE%""", package.Name)
                                   writer.WriteLine("    ECHO Installing {0}...", package.Name)
                                   writer.WriteLine("    ECHO.")
                                   writer.WriteLine()
                                   writer.WriteLine("    REM Run the command")
                                   writer.WriteLine("    WINGET install --id ""{0}"" --silent --accept-package-agreements --accept-source-agreements", package.Id)
                                   writer.WriteLine()
                                   writer.WriteLine("    REM Capture the errorlevel into a variable (delayed expansion safe)")
                                   writer.WriteLine("    SET RESULT_CODE=!ERRORLEVEL!")
                                   writer.WriteLine()
                                   writer.WriteLine("    REM Debugging output")
                                   writer.WriteLine("    ECHO [DEBUG] RESULT_CODE=!RESULT_CODE!")
                                   writer.WriteLine("    ECHO [%TIME%] DEBUG: RESULT_CODE=!RESULT_CODE! >> ""%LOG_FILE%""")
                                   writer.WriteLine()
                                   writer.WriteLine("    REM Branch on the stored value")
                                   writer.WriteLine("    IF !RESULT_CODE! EQU 0 (")
                                   writer.WriteLine("        COLOR 2F")
                                   writer.WriteLine("        ECHO [SUCCESS] {0} installed successfully!", package.Name)
                                   writer.WriteLine("        ECHO [%TIME%] SUCCESS: {0} >> ""%LOG_FILE%""", package.Name)
                                   writer.WriteLine("        SET /A SUCCESS_COUNT+=1")
                                   writer.WriteLine("        GOTO InstallEnd{0}", pkgNum)
                                   writer.WriteLine("    )")
                                   writer.WriteLine()
                                   writer.WriteLine("    IF NOT !RESULT_CODE! EQU 0 (")
                                   writer.WriteLine("        COLOR 4F")
                                   writer.WriteLine("        ECHO [ERROR] Failed to install {0}", package.Name)
                                   writer.WriteLine("        ECHO Error Code: !RESULT_CODE!")
                                   writer.WriteLine("        ECHO [%TIME%] FAILED: {0} (Error: !RESULT_CODE!) >> ""%LOG_FILE%""", package.Name)
                                   writer.WriteLine("        SET /A FAILED_COUNT+=1")
                                   writer.WriteLine("        GOTO InstallEnd{0}", pkgNum)
                                   writer.WriteLine("    )")
                                   writer.WriteLine(") ELSE (")
                                   writer.WriteLine("    ECHO Skipping {0}...", package.Name)
                                   writer.WriteLine("    ECHO [%TIME%] SKIPPED: {0} >> ""%LOG_FILE%""", package.Name)
                                   writer.WriteLine("    SET /A SKIPPED_COUNT+=1")
                                   writer.WriteLine("    GOTO InstallEnd{0}", pkgNum)
                                   writer.WriteLine(")")
                                   writer.WriteLine(":InstallEnd{0}", pkgNum)
                                   writer.WriteLine("TIMEOUT /T 2 /NOBREAK >NUL")
                                   writer.WriteLine()
                               Next

                               ' Post-installation tasks
                               If deleteLinks = True Then
                                   writer.WriteLine("REM ========================================")
                                   writer.WriteLine("REM Cleanup Desktop Shortcuts")
                                   writer.WriteLine("REM ========================================")
                                   writer.WriteLine("COLOR 0E")
                                   writer.WriteLine("CLS")
                                   writer.WriteLine("ECHO.")
                                   writer.WriteLine("ECHO Cleaning up desktop shortcuts...")
                                   writer.WriteLine("ECHO [%TIME%] Cleaning desktop shortcuts >> ""%LOG_FILE%""")
                                   writer.WriteLine("DEL /F /Q ""C:\Users\Public\Desktop\*.lnk"" 2>NUL")
                                   writer.WriteLine("DEL /F /Q ""%USERPROFILE%\Desktop\*.lnk"" 2>NUL")
                                   writer.WriteLine("ECHO Desktop cleanup completed.")
                                   writer.WriteLine("TIMEOUT /T 2 /NOBREAK >NUL")
                                   writer.WriteLine()
                               End If

                               If captureStartMenu = True Then
                                   writer.WriteLine("REM ========================================")
                                   writer.WriteLine("REM Apply Start Menu Layout")
                                   writer.WriteLine("REM ========================================")
                                   writer.WriteLine("COLOR 0E")
                                   writer.WriteLine("CLS")
                                   writer.WriteLine("ECHO.")
                                   writer.WriteLine("ECHO Applying Start Menu layout...")
                                   writer.WriteLine("ECHO [%TIME%] Applying Start Menu layout >> ""%LOG_FILE%""")
                                   Application.Current.Dispatcher.Invoke(Sub() CaptureStartMenuLayout())
                                   writer.WriteLine("IF EXIST ""StartMenuLayout.bin"" (")
                                   writer.WriteLine("    COPY /Y ""StartMenuLayout.bin"" ""%LOCALAPPDATA%\Packages\Microsoft.Windows.StartMenuExperienceHost_cw5n1h2txyewy\LocalState\start2.bin""")
                                   writer.WriteLine("    ECHO Start Menu layout applied.")
                                   writer.WriteLine(") ELSE (")
                                   writer.WriteLine("    ECHO Warning: StartMenuLayout.bin not found!")
                                   writer.WriteLine(")")
                                   writer.WriteLine("TIMEOUT /T 2 /NOBREAK >NUL")
                                   writer.WriteLine()
                               End If

                               ' Final summary
                               writer.WriteLine("REM ========================================")
                               writer.WriteLine("REM Installation Summary")
                               writer.WriteLine("REM ========================================")
                               writer.WriteLine("COLOR 0F")
                               writer.WriteLine("CLS")
                               writer.WriteLine("SET ""END_TIME=%TIME%""")
                               writer.WriteLine("ECHO.")
                               writer.WriteLine("ECHO ========================================")
                               writer.WriteLine("ECHO   Installation Summary")
                               writer.WriteLine("ECHO ========================================")
                               writer.WriteLine("ECHO.")
                               writer.WriteLine("ECHO Total Packages: %TOTAL_PACKAGES%")
                               writer.WriteLine("ECHO Attempted: %INSTALL_COUNT%")
                               writer.WriteLine("ECHO Successful: %SUCCESS_COUNT%")
                               writer.WriteLine("ECHO Failed: %FAILED_COUNT%")
                               writer.WriteLine("ECHO Skipped: %SKIPPED_COUNT%")
                               writer.WriteLine("ECHO.")
                               writer.WriteLine("ECHO Start Time: %START_TIME%")
                               writer.WriteLine("ECHO End Time: %END_TIME%")
                               writer.WriteLine("ECHO.")
                               writer.WriteLine("ECHO Log file: %LOG_FILE%")
                               writer.WriteLine("ECHO ========================================")
                               writer.WriteLine("ECHO.")
                               writer.WriteLine()

                               ' Write summary to log
                               writer.WriteLine("ECHO. >> ""%LOG_FILE%""")
                               writer.WriteLine("ECHO ======================================== >> ""%LOG_FILE%""")
                               writer.WriteLine("ECHO Installation Summary >> ""%LOG_FILE%""")
                               writer.WriteLine("ECHO ======================================== >> ""%LOG_FILE%""")
                               writer.WriteLine("ECHO Total Packages: %TOTAL_PACKAGES% >> ""%LOG_FILE%""")
                               writer.WriteLine("ECHO Attempted: %INSTALL_COUNT% >> ""%LOG_FILE%""")
                               writer.WriteLine("ECHO Successful: %SUCCESS_COUNT% >> ""%LOG_FILE%""")
                               writer.WriteLine("ECHO Failed: %FAILED_COUNT% >> ""%LOG_FILE%""")
                               writer.WriteLine("ECHO Skipped: %SKIPPED_COUNT% >> ""%LOG_FILE%""")
                               writer.WriteLine("ECHO End Time: %END_TIME% >> ""%LOG_FILE%""")
                               writer.WriteLine("ECHO ======================================== >> ""%LOG_FILE%""")
                               writer.WriteLine()

                               ' Conditional exit
                               writer.WriteLine("IF %FAILED_COUNT% GTR 0 (")
                               writer.WriteLine("    COLOR 4F")
                               writer.WriteLine("    ECHO WARNING: Some installations failed!")
                               writer.WriteLine("    ECHO Check %LOG_FILE% for details.")
                               writer.WriteLine(") ELSE (")
                               writer.WriteLine("    COLOR 2F")
                               writer.WriteLine("    ECHO All installations completed successfully!")
                               writer.WriteLine(")")
                               writer.WriteLine("ECHO.")
                               writer.WriteLine("PAUSE")
                               writer.WriteLine("ENDLOCAL")
                               writer.WriteLine("EXIT /B %FAILED_COUNT%")
                           End Using

                           ' Write JSON configuration file (replaces manifest and ids files)
                           Using writer As New StreamWriter(configFile)
                               writer.WriteLine("{")
                               writer.WriteLine("  ""scriptName"": ""{0}"",", filename)
                               writer.WriteLine("  ""generatedDate"": ""{0}"",", DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
                               writer.WriteLine("  ""deleteShortcuts"": {0},", If(deleteLinks, "true", "false"))
                               writer.WriteLine("  ""applyStartMenu"": {0},", If(captureStartMenu, "true", "false"))
                               writer.WriteLine("  ""packages"": [")
                               For i = 0 To selectedPackages.Count - 1
                                   Dim pkg = selectedPackages(i)
                                   writer.Write("    {{""name"": ""{0}"", ""id"": ""{1}""", pkg.Name.Replace("""", "\"""), pkg.Id.Replace("""", "\"""))
                                   If Not String.IsNullOrEmpty(pkg.Version) Then
                                       writer.Write(", ""version"": ""{0}""", pkg.Version.Replace("""", "\"""))
                                   End If
                                   If Not String.IsNullOrEmpty(pkg.Match) Then
                                       writer.Write(", ""match"": ""{0}""", pkg.Match.Replace("""", "\"""))
                                   End If
                                   If Not String.IsNullOrEmpty(pkg.Source) Then
                                       writer.Write(", ""source"": ""{0}""", pkg.Source.Replace("""", "\"""))
                                   End If
                                   writer.WriteLine(If(i < selectedPackages.Count - 1, "}},", "}}"))
                               Next
                               writer.WriteLine("  ]")
                               writer.WriteLine("}")
                           End Using
                       End Sub)
    End Function

    Private Sub CaptureStartMenuLayout()
        Try
            Dim localAppDataPath = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData)
            Dim startMenuPath = Path.Combine(localAppDataPath, "Packages\Microsoft.Windows.StartMenuExperienceHost_cw5n1h2txyewy\LocalState")
            Dim files = Directory.GetFiles(startMenuPath, "start*.bin")

            If files.Length > 0 Then
                File.Copy(files(0), "StartMenuLayout.bin", True)
            End If
        Catch ex As Exception
            Debug.WriteLine($"Failed to capture start menu: {ex.Message}")
        End Try
    End Sub

    ' ==================== MENU HANDLERS ====================

    Private Sub NewMenuItem_Click(sender As Object, e As RoutedEventArgs)
        ClearAll()
    End Sub

    Private Async Sub OpenMenuItem_Click(sender As Object, e As RoutedEventArgs)
        Dim openDialog As New OpenFileDialog With {
            .Title = "Select Configuration File",
            .Filter = "JSON Configuration Files|*.json|All Files|*.*",
            .CheckFileExists = True
        }

        If openDialog.ShowDialog() = True Then
            Try
                UpdateProgress("Loading...", "Opening configuration file")
                StartWingetIndicator()

                Await LoadConfigurationAsync(openDialog.FileName)

                UpdateProgressSuccess("Loaded", "Configuration loaded successfully")
            Catch ex As Exception
                UpdateProgressError("Load Failed", ex.Message)
            Finally
                StopWingetIndicator()
                ResetProgress()
            End Try
        End If
    End Sub

    Private Async Function LoadConfigurationAsync(configPath As String) As Task
        Await Task.Run(Sub()
                           _selectedPackages.Clear()

                           Dispatcher.Invoke(Sub()
                                                 SelectedNamesListBox.Items.Clear()
                                                 SelectedIDsListBox.Items.Clear()
                                             End Sub)

                           ' Load JSON configuration
                           Dim jsonContent = File.ReadAllText(configPath)

                           ' Parse JSON manually (simple parser for this structure)
                           Dim packagesStartIndex = jsonContent.IndexOf("""packages"":")
                           If packagesStartIndex > 0 Then
                               Dim packagesArrayStart = jsonContent.IndexOf("[", packagesStartIndex)
                               Dim packagesArrayEnd = jsonContent.IndexOf("]", packagesArrayStart)

                               If packagesArrayStart > 0 AndAlso packagesArrayEnd > packagesArrayStart Then
                                   Dim packagesJson = jsonContent.Substring(packagesArrayStart + 1, packagesArrayEnd - packagesArrayStart - 1)

                                   ' Split by package objects
                                   Dim packageMatches = Regex.Matches(packagesJson, "\{[^}]+\}")

                                   For Each match As Match In packageMatches
                                       Dim packageJson = match.Value

                                       ' Extract fields
                                       Dim nameMatch = Regex.Match(packageJson, """name"":\s*""([^""]+)""")
                                       Dim idMatch = Regex.Match(packageJson, """id"":\s*""([^""]+)""")
                                       Dim versionMatch = Regex.Match(packageJson, """version"":\s*""([^""]+)""")
                                       Dim matchFieldMatch = Regex.Match(packageJson, """match"":\s*""([^""]+)""")
                                       Dim sourceMatch = Regex.Match(packageJson, """source"":\s*""([^""]+)""")

                                       If nameMatch.Success AndAlso idMatch.Success Then
                                           Dim package As New PackageInfo With {
                                               .Name = nameMatch.Groups(1).Value.Replace("\""", """"),
                                               .Id = idMatch.Groups(1).Value.Replace("\""", """"),
                                               .Version = If(versionMatch.Success, versionMatch.Groups(1).Value.Replace("\""", """"), ""),
                                               .Match = If(matchFieldMatch.Success, matchFieldMatch.Groups(1).Value.Replace("\""", """"), ""),
                                               .Source = If(sourceMatch.Success, sourceMatch.Groups(1).Value.Replace("\""", """"), "")
                                           }
                                           _selectedPackages.Add(package)
                                       End If
                                   Next
                               End If
                           End If

                           Dispatcher.Invoke(Sub()
                                                 RefreshSelectedPackagesList(GetSelectedListFilter())
                                                 Dim baseFilename = Path.GetFileNameWithoutExtension(configPath)
                                                 FilenameTextBox.Text = baseFilename

                                                 SelectedNamesListBox.Visibility = Visibility.Visible
                                                 RemoveButton.Visibility = Visibility.Visible
                                                 FilenameTextBox.IsEnabled = True
                                                 BuildButton.IsEnabled = True
                                                 DeleteLinksCheckBox.IsEnabled = True
                                                 CaptureStartMenuCheckBox.IsEnabled = True
                                                 SelectedSearchTextBox.IsEnabled = _selectedPackages.Count > 0
                                             End Sub)
                       End Sub)
    End Function

    Private Async Sub GetCurrentListMenuItem_Click(sender As Object, e As RoutedEventArgs)
        UpdateProgress("Loading...", "Getting installed packages")
        StartWingetIndicator()

        SearchResultsListBox.Items.Clear()
        _allSearchResults.Clear()

        Try
            _allSearchResults = Await GetInstalledPackagesAsync()

            If _allSearchResults.Count = 0 Then
                UpdateProgressError("No Results", "No installed packages found")
                SearchResultsListBox.Visibility = Visibility.Collapsed
                SourceFilterPanel.Visibility = Visibility.Collapsed
            Else
                ' Populate the source filter dropdown
                PopulateSourceFilter()

                ' Apply the current filter
                ApplySourceFilter()

                SearchResultsListBox.Visibility = Visibility.Visible
            End If
        Catch ex As Exception
            UpdateProgressError("Failed", ex.Message)
        Finally
            StopWingetIndicator()
            If _allSearchResults.Count = 0 Then
                ResetProgress()
            End If
        End Try
    End Sub

    Private Async Function GetInstalledPackagesAsync() As Task(Of List(Of PackageInfo))
        Return Await RunWingetListAsync()
    End Function

    Private Async Function RunWingetListAsync() As Task(Of List(Of PackageInfo))
        Return Await Task.Run(Function() As List(Of PackageInfo)
                                  Dim packages As New List(Of PackageInfo)

                                  Try
                                      Dim workerFile = "WorkerFile.cmd"
                                      Dim outputFile = "WorkerOutput.txt"

                                      If File.Exists(workerFile) Then File.Delete(workerFile)
                                      If File.Exists(outputFile) Then File.Delete(outputFile)

                                      Using writer As New StreamWriter(workerFile)
                                          writer.WriteLine("@Echo Off")
                                          writer.WriteLine("CLS")
                                          writer.WriteLine($"WINGET List > {outputFile}")
                                      End Using

                                      Dim psi As New ProcessStartInfo With {
                                          .FileName = "cmd.exe",
                                          .Arguments = $"/c {workerFile}",
                                          .UseShellExecute = False,
                                          .CreateNoWindow = True,
                                          .WindowStyle = ProcessWindowStyle.Hidden
                                      }

                                      Using proc = Process.Start(psi)
                                          proc.WaitForExit()
                                      End Using

                                      If File.Exists(outputFile) Then
                                          Dim lines = File.ReadAllLines(outputFile).ToList()

                                          Dim headerIndex = -1
                                          For i = 0 To lines.Count - 1
                                              If lines(i).Contains("---") Then
                                                  headerIndex = i - 1
                                                  Exit For
                                              End If
                                          Next

                                          If headerIndex >= 0 AndAlso headerIndex < lines.Count Then
                                              _headerValue = lines(headerIndex)

                                              For i = headerIndex + 2 To lines.Count - 1
                                                  If Not String.IsNullOrWhiteSpace(lines(i)) Then
                                                      packages.Add(ParsePackageLine(lines(i)))
                                                  End If
                                              Next
                                          End If
                                      End If

                                      If File.Exists(workerFile) Then File.Delete(workerFile)
                                      If File.Exists(outputFile) Then File.Delete(outputFile)

                                  Catch ex As Exception
                                      Debug.WriteLine($"List error: {ex.Message}")
                                  End Try

                                  Return packages
                              End Function)
    End Function

    Private Sub CaptureStartMenuMenuItem_Click(sender As Object, e As RoutedEventArgs)
        Try
            CaptureStartMenuLayout()
            MessageBox.Show("Start Menu Layout captured successfully!", "Success",
                          MessageBoxButton.OK, MessageBoxImage.Information)
        Catch ex As Exception
            MessageBox.Show($"Failed to capture Start Menu Layout: {ex.Message}", "Error",
                          MessageBoxButton.OK, MessageBoxImage.Error)
        End Try
    End Sub

    Private Sub ClearAll()
        SearchTextBox.Text = "Search packages..."
        SearchResultsListBox.Items.Clear()
        _allSearchResults.Clear()
        SelectedNamesListBox.Items.Clear()
        SelectedIDsListBox.Items.Clear()
        _selectedPackages.Clear()
        FilenameTextBox.Text = ""
        SetupInitialVisibility()
    End Sub

    ' ==================== BING WALLPAPER ====================

    Private Async Function SetBingWallpaperAsync() As Task
        If Not Await IsInternetAvailableAsync() Then Return

        Const xmlUrl As String = "https://www.bing.com/HPImageArchive.aspx?format=xml&idx=0&n=1&mkt=en-US"
        Dim xmlContent As String

        Try
            Using resp = Await httpClient.GetAsync(xmlUrl)
                resp.EnsureSuccessStatusCode()
                xmlContent = Await resp.Content.ReadAsStringAsync()
            End Using
        Catch
            Return
        End Try

        Dim doc As XDocument
        Try
            doc = XDocument.Parse(xmlContent)
        Catch
            Return
        End Try

        Dim imageElement = doc.Root?.Element("image")
        If imageElement Is Nothing Then Return

        Dim relativeUrl = imageElement.Element("url")?.Value
        If String.IsNullOrWhiteSpace(relativeUrl) Then Return

        Dim fullImageUrl = If(relativeUrl.StartsWith("http", StringComparison.OrdinalIgnoreCase),
                              relativeUrl,
                              "https://www.bing.com" & relativeUrl)

        Dim imageBytes As Byte()
        Try
            imageBytes = Await httpClient.GetByteArrayAsync(fullImageUrl)
        Catch
            Return
        End Try

        Dim bmp As New BitmapImage()
        Try
            Using ms As New MemoryStream(imageBytes)
                bmp.BeginInit()
                bmp.CacheOption = BitmapCacheOption.OnLoad
                bmp.StreamSource = ms
                bmp.EndInit()
                bmp.Freeze()
            End Using
        Catch
            Return
        End Try

        Me.Background = New ImageBrush(bmp) With {
            .Stretch = Stretch.UniformToFill,
            .AlignmentX = AlignmentX.Center,
            .AlignmentY = AlignmentY.Center
        }
        If RootDock IsNot Nothing Then
            RootDock.Background = Brushes.Transparent
        End If

        Dim headline = imageElement.Element("headline")?.Value
        Dim copyright = imageElement.Element("copyright")?.Value

        HeadingTextBlock.Text = If(String.IsNullOrWhiteSpace(headline), "", headline)
        CopyrightTextBlock.Text = If(String.IsNullOrWhiteSpace(copyright), "", copyright)
    End Function

    Private Shared Async Function IsInternetAvailableAsync() As Task(Of Boolean)
        Try
            Using resp = Await httpClient.GetAsync("https://www.bing.com", HttpCompletionOption.ResponseHeadersRead)
                Return resp.IsSuccessStatusCode
            End Using
        Catch
            Return False
        End Try
    End Function

    <DllImport("user32.dll")>
    Private Shared Sub SetSystemDarkMode(darkMode As Boolean)
    End Sub

    Private Sub EnableDarkTitleBar()
        Try
            Dim hwnd = New System.Windows.Interop.WindowInteropHelper(Me).Handle
            Dim useImmersiveDarkMode As Integer = 20
            Dim enabled As Integer = 1
            DwmSetWindowAttribute(hwnd, useImmersiveDarkMode, enabled, Marshal.SizeOf(enabled))
        Catch ex As Exception
            Debug.WriteLine($"Failed to set dark title bar: {ex.Message}")
        End Try
    End Sub

    <DllImport("dwmapi.dll", PreserveSig:=True)>
    Private Shared Function DwmSetWindowAttribute(hwnd As IntPtr, attr As Integer, ByRef attrValue As Integer, attrSize As Integer) As Integer
    End Function

    Private Sub SearchResultsListBox_SelectionChanged(sender As Object, e As SelectionChangedEventArgs) Handles SearchResultsListBox.SelectionChanged
        UpdateStoreLinkButton()
    End Sub

    Private Sub UpdateStoreLinkButton()
        Dim storeButton = GetStoreLinkButton()
        If storeButton Is Nothing Then Return

        Dim package = GetPackageFromListBoxItem(SearchResultsListBox.SelectedItem)
        Dim enabled = PackageSupportsStoreDetail(package)

        storeButton.Visibility = If(enabled, Visibility.Visible, Visibility.Collapsed)
        storeButton.Tag = If(enabled, package, Nothing)
    End Sub

    Private Function GetStoreLinkButton() As Button
        Return TryCast(Me.FindName("StoreLinkButton"), Button)
    End Function
End Class