[CmdletBinding()]
param()

if ($env:PROCESSOR_ARCHITECTURE -ne "x86") { Throw "This script requires an x86 process" }
if (([Management.AutoMation.Runspaces.Runspace]::DefaultRunSpace).ApartmentState -eq "STA") {
Throw "This script requires an MTA apartment" }

function Invoke-ScriptEntryPoint {
    AsyncRun-UIThread $UIThreadScript
    # Attach to Skype
    $Global:Skype = Get-Skype -Starting { param($Handle, $FriendlyName)
        try {
            $UIProxy.UIThread_Events.GenerateEvent("Skype.OpenWindow", $null, $FriendlyName, $null) | Out-Null
            $Global:Skype.SendMessage([string]::Empty, $Handle)
        } catch [exception] {
            Write-Host ($_ | Out-String)
        }
    } -MessageReceived { param($Sender, $Msg)
        Write-Warning "On Skype.OnMessage [MTA Get-Skype Callback]"
        $UIProxy.UIThread_Events.GenerateEvent("Skype.OnMessage", $null, ($Msg, $Sender), $null) | Out-Null
    }
    # Wire up sending messages
    Register-EngineEvent "Skype.SendMessage" -Action {
        Write-Warning "Skype.SendMessage [MTA]"
        try {
            if ($Args[0] -eq "/neg") {
                Write-Warning "Negotiating"
                $Global:Skype.Negotiate()
            } elseif ($Args[0] -eq "/keys") {
                Write-Host ($Skype.Keys | Out-String)
            } else {
                $Global:Skype.SendMessage($Args[0], $env:SkypeFriend)
            }
        } catch [exception] {
            Write-Warning "Skype failed to send: $_"
        }
    } | Out-Null
    # Accept the "window closed" event
    Register-EngineEvent "Skype.SecureUIClosed" -Action { 
        Write-Warning "Skype.SecureUIClosed [MTA]"
        $Global:Running = $false
    } | Out-Null
    # Accept debug events from the background
    Register-EngineEvent "Skype.Debug" -Action { 
        Write-Warning $Args[0]
    } | Out-Null
    # Run until we are told to quit
    $Global:Running = $true
    while ($Running)  { sleep -Milliseconds 500 }
    # Cleanup
    Unregister-Event -SourceIdentifier "Skype.Debug"
    Unregister-Event -SourceIdentifier "Skype.SecureUIClosed"
    Unregister-Event -SourceIdentifier "Skype.SendMessage"
}

$UIThreadScript = { # [STA Thread for UI]
$Interop_TypeDef = @'
[DllImport("Kernel32.dll")]
public static extern IntPtr GetConsoleWindow();

[DllImport("user32.dll")]
public static extern bool SetForegroundWindow(IntPtr hWnd);

[DllImport("user32.dll")]
public static extern IntPtr FindWindow(string ClassName, string Title);

[DllImport("user32.dll")]
 public static extern bool BringWindowToTop(IntPtr hWnd);

[DllImport("user32.dll")]
public static extern bool GetWindowRect(IntPtr hWnd, out RECT lpRect);

[DllImport("user32.dll")]
public static extern int SetWindowLong(IntPtr window, int index, int value);

[DllImport("user32.dll")]
public static extern int GetWindowLong(IntPtr window, int index);

public const int GWL_EXSTYLE = -20;
public const int WS_EX_TOOLWINDOW = 0x00000080;
public const int WS_EX_APPWINDOW = 0x00040000;

public struct RECT {
    public int Left;  
    public int Top;   
    public int Right;   
    public int Bottom;
}
'@
try {
    function Send-Event($EventName, $Text = $null) {
        $UIProxy.SkypeThread_Events.GenerateEvent($EventName, $null, $Text, $null) | Out-Null
    }
    Send-Event "Skype.Debug" "UIThreadScript"
    function Show-Control { 
       param([Parameter(Mandatory=$true, ParameterSetName="VisualElement", ValueFromPipeline=$true, ValueFromPipelineByPropertyName=$true)]      
       [Windows.Media.Visual] $control,     
       [Parameter(Mandatory=$true, ParameterSetName="Xaml", ValueFromPipeline=$true, ValueFromPipelineByPropertyName=$true)]      
       [string] $xaml,     
       [Parameter(ValueFromPipelineByPropertyName=$true,Position=0)]      
       [Hashtable] $event, 
       [Hashtable] $windowProperties) 

       Begin { 
           $Global:window = New-Object Windows.Window 
           $window.SizeToContent = "WidthAndHeight" 

           if ($windowProperties) { 
               foreach ($kv in $windowProperties.GetEnumerator()) { 
                   $window."$($kv.Key)" = $kv.Value 
               } 
           } 
           $visibleElements = @() 
           $windowEvents = @() 
       } 

       Process {       
           switch ($psCmdlet.ParameterSetName) { 
           "Xaml" { 
               $f = [System.xml.xmlreader]::Create([System.IO.StringReader] $xaml) 
               $visibleElements+=([system.windows.markup.xamlreader]::load($f))       
           } 
           "VisualElement" { 
               $visibleElements+=$control 
           } 
           } 
           if ($event) { 
               $element = $visibleElements[-1]       
               foreach ($evt in $event.GetEnumerator()) { 
                   # If the event name is like *.*, it is an event on a named target, otherwise, it's on any of the events on the top level object 
                   if ($evt.Key.Contains(".")) { 
                       $targetName = $evt.Key.Split(".")[1].Trim() 
                       if ($evt.Key -like "Window.*") { 
                           $target = $window 
                       } else { 
                           $target = ($visibleElements[-1]).FindName(($evt.Key.Split(".")[0]))                   
                       }                       
                   } else { 
                       $target = $visibleElements[-1] 
                       $targetName = $evt.Key 
                   } 
                   $target | Get-Member -type Event | 
                     ? { $_.Name -eq $targetName } | 
                     % { 
                       $eventMethod = $target."add_$targetName" 
                       $eventMethod.Invoke($evt.Value) 
                     }               
               } 
           } 
        } 

        End { 
            if ($visibleElements.Count -gt 1) { 
                $wrapPanel = New-Object Windows.Controls.WrapPanel 
                $visibleElements | % { $null = $wrapPanel.Children.Add($_) } 
                $window.Content = $wrapPanel 
            } else { 
                if ($visibleElements) { 
                    $window.Content = $visibleElements[0] 
                }
            }
            $null = $window.Show() # Need to also run UI Thread!
        } 
    }

    Add-Type –assemblyName PresentationFramework, PresentationCore, WindowsBase
    #Add-Type –assemblyName System.Web # For HttpUtility HTML Encode/Decode

    Add-Type -MemberDefinition $Interop_TypeDef -Namespace Win32 -Name Interop

    Show-Control -Xaml `
@" 
<DockPanel LastChildFill="True"  xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation">
    <TextBox SpellCheck.IsEnabled="True"
             TextWrapping="Wrap"
             AcceptsReturn="True"
             MaxLines="4"
             VerticalScrollBarVisibility="Auto"
             HorizontalScrollBarVisibility="Disabled"
             DockPanel.Dock="Bottom" Name="TextInput" />
    <WebBrowser Name="Web" />
</DockPanel>
"@ -event @{ 
        # Text Box input handler
        "TextInput.PreviewKeyDown" = {
            param($sender, [Windows.Input.KeyEventArgs]$e)
            if ($e.Key -eq [Windows.Input.Key]::Return) {
                if ([Windows.Input.Keyboard]::Modifiers -eq [Windows.Input.ModifierKeys]::Shift -or
                    [Windows.Input.Keyboard]::Modifiers -eq [Windows.Input.ModifierKeys]::Control) {
                    # Newline
                } else {
                    # Clear the input box and send the message
                    $e.Handled = $true
                    [System.Windows.Controls.TextBox]$TextInput = $window.Content.FindName("TextInput") 
                    $Msg = $($TextInput.Text)
                    $TextInput.Text = ""

                    if ($Msg -eq "/close") { 
                        $window.Close()
                        return 
                    }

                    $web.Document.IHTMLDocument2_writeln("<span style='color:red;'>Me</span>: $Msg<br />")
                    $web.Document.body.scrollTop = 99999
                    # Signal for the MTA thread to send the message
                    Send-Event "Skype.SendMessage" $Msg
                }
            }
        }
        "Window.Closing" = { 
            Send-Event "Skype.Debug" "Window.Closing"
            $window.IsEnabled = $false
        }
        "Web.Navigated" = {
            # Set the initial window contents
            $web.Document.IHTMLDocument2_writeln("Secure Chat<br />")
            # Handle display of messages
            Register-EngineEvent "Skype.OnMessage" -Action { 
                Send-Event "Skype.Debug" "Skype.OnMessage [STA]"
                $Msg, $Sender = $Args[0], $Args[1]
                    
                $html = "<span style='color:blue;'>$Sender</span>: $Msg<br />"
                $web.Document.IHTMLDocument2_writeln($html)
                $web.Document.body.scrollTop = 99999
            } | Out-Null
        }
        "Window.SourceInitialized" = { 
            Send-Event "Skype.Debug" "Window.SourceInitialized"
            [System.Windows.Controls.WebBrowser]$Global:Web = $window.Content.FindName("Web") 
            # DPI support to keep the WPF window lined up with the target window
            $Global:DPIScale = [System.Windows.presentationSource]::FromVisual($window).CompositionTarget.TransformToDevice.M11

            Register-EngineEvent "Skype.CloseSecureUI" -Action { 
                $window.Close()
            } | Out-Null

            # Initialize Trident
            $web.Navigate([uri]"about:blank")
        } 
    } -windowProperties @{
        'Height'=200
        'Width'=400
        'Left'=-9999
        'Top'=-9999
        'SizeToContent'='Manual'
        'WindowStyle'='None'
        'ShowInTaskbar'=$false
    } 

    $hwnd = [IntPtr]::Zero
    Register-EngineEvent "Skype.OpenWindow" -Action {
        $Global:Hwnd = [Win32.Interop]::FindWindow("TConversationForm", $Args[0])
    } | Out-Null

    $rct = New-Object Win32.Interop+RECT
    $wih = New-Object System.Windows.Interop.WindowInteropHelper $window
    $windowStyle = [Win32.Interop]::GetWindowLong($wih.Handle, [Win32.Interop]::GWL_EXSTYLE);
    [Win32.Interop]::SetWindowLong($wih.Handle, [Win32.Interop]::GWL_EXSTYLE, $windowStyle -bor [Win32.Interop]::WS_EX_TOOLWINDOW);

    Send-Event "Skype.Debug" "Use '!' to show the secure window."
    # DPI
    $Global:DPIScale = 1
    #  PS needs to get control to do eventing, so the message pump needs to live in PS
    while($window.IsEnabled) {
        # Makes WPF controls accept input properly
        [Windows.Threading.Dispatcher]::CurrentDispatcher.Invoke({}, [Windows.Threading.DispatcherPriority]::Background)
        # Dispatch outstanding messages
        [Windows.Forms.Application]::DoEvents()

        if ($hwnd.ToInt32() -ne 0)
        {
            [Win32.Interop]::GetWindowRect($Hwnd, [ref]$rct)
            $window.Left = ($rct.Left * $DPIScale)
            $window.Top = ($rct.Bottom * $DPIScale)
            $window.Width = (($rct.Right - $rct.Left) * $DPIScale)
            [Win32.Interop]::BringWindowToTop($wih.Handle)
        }
        sleep -Milliseconds 33
    }
} catch [exception] {
    Send-Event "Skype.Debug" "UI Thread Error: $($Error | Out-String)"
}
# Signal that the UI is now closed
Send-Event "Skype.SecureUIClosed" $null
# TODO: LEAK: free this runspace
} # END [STA Thread for UI]

function AsyncRun-UIThread($Script)
{
    # Create a UI thread on a background STA runspace
    $Global:UIProxy = @{
        'SkypeThread_Events'=([Management.AutoMation.Runspaces.Runspace]::DefaultRunSpace).Events
        'UIThread_Events'=@()
        'Runspaces'=@{}
        'Windows'=@{}
    }
    $newRunspace = [RunSpaceFactory]::CreateRunspace()
    $newRunspace.ApartmentState = "STA"
    $newRunspace.Open()
    $newRunspace.SessionStateProxy.setVariable("UIProxy", $UIProxy)
    $newPowerShell = [PowerShell]::Create()
    $newPowerShell.Runspace = $newRunspace
    $newPowerShell.AddScript($Script).BeginInvoke() | Out-Null
    # Save off info so we can communicate
    $UIProxy.Runspaces.Add($newRunspace, $newPowerShell)
    $UIProxy.UIThread_Events = $newRunspace.Events
}

Invoke-ScriptEntryPoint