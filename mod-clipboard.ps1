# Module for local clipboard handling functions

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

Add-Type -TypeDefinition @"
using System;
using System.Runtime.InteropServices;
using System.Windows.Forms;

public class ClipboardMonitor : Form
{
    private const int WM_CLIPBOARDUPDATE = 0x031D;
    
    [DllImport("user32.dll", SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    public static extern bool AddClipboardFormatListener(IntPtr hwnd);
    
    [DllImport("user32.dll", SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    public static extern bool RemoveClipboardFormatListener(IntPtr hwnd);
    
    public event Action ClipboardChanged;
    
    public ClipboardMonitor()
    {
        this.WindowState = FormWindowState.Minimized;
        this.ShowInTaskbar = false;
        this.Visible = false;
    }
    
    protected override void OnHandleCreated(EventArgs e)
    {
        base.OnHandleCreated(e);
        AddClipboardFormatListener(this.Handle);
    }
    
    protected override void OnHandleDestroyed(EventArgs e)
    {
        RemoveClipboardFormatListener(this.Handle);
        base.OnHandleDestroyed(e);
    }
    
    protected override void WndProc(ref Message m)
    {
        if (m.Msg == WM_CLIPBOARDUPDATE)
        {        
            if (ClipboardChanged != null)
            {
                ClipboardChanged.Invoke();        
            }
        }   
        base.WndProc(ref m);
    }
}
"@ -ReferencedAssemblies System.Windows.Forms,System.Drawing

function Show-ClipboardNotification {
    param(
        [string]$Title,
        [string]$Message,
        [string]$Icon = "Info" # Parameter Icon is not used in current implementation
    )
    
    try {
        # Create notification object
        $notification = New-Object System.Windows.Forms.NotifyIcon
        $notification.Icon = [System.Drawing.SystemIcons]::Information # Uses Information icon by default
        $notification.BalloonTipIcon = [System.Windows.Forms.ToolTipIcon]::Info
        $notification.BalloonTipText = $Message
        $notification.BalloonTipTitle = $Title
        $notification.Visible = $true
        
        # Show balloon tip
        $notification.ShowBalloonTip(3000) # Show for 3 seconds
        
        # Clean up after delay
        Start-Sleep -Milliseconds 3500 # Wait a bit longer than the tip display time
        $notification.Dispose()
        
    }
    catch {
        # Fallback to console if notifications fail
        Write-Host "NOTIFICATION: $Title - $Message"
    }
}

# Export functions and classes if needed, though dot-sourcing makes them available.
# Export-ModuleMember -Function Show-ClipboardNotification
# Note: Add-Type defined classes are available globally once loaded.
