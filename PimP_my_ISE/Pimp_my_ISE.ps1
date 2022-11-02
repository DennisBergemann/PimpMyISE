#########################################
# PimP my ISE
# Author: Dennis Bergemann
# 
# Stellt Funktionen und Menüeinträge für
# die Powershell_ISE
#########################################


#Typedeklarationen

# Bekannte Typeacceleratoren erfragen und an den Accelerator [accelerators] hängen
# Abrufbar mit [acclerators]::get
$typeAcc=[psobject].Assembly.GetType('System.Management.Automation.TypeAccelerators')
$typeAcc::Add('accelerators',$typeAcc)





#Functions

## sucht nach dem Suchwort in der MSDN
function Search-MSDN
{
    [CmdletBinding()]
    param(
    [string] $such = "Powershell",
    [switch] $full
    )
    
    if($full)
    {
        $website = "https://social.msdn.microsoft.com/Search/de-DE?query=$($such)&pgArea=header&emptyWatermark=true&ac=2"
        $ie = New-Object -ComObject InternetExplorer.Application 
        $ie.Visible = $true
        $ie.Navigate2($website)
    }
    else
    {
        Get-Help $such -Online -ErrorAction SilentlyContinue
    } 
}

function Search-Selected
{
    [CmdletBinding()]
    param(    
    [switch] $full
    )
    # Gibt das Selectierte Wort im Editor zurück
    # ist keins ausgewählt, wird stattdessen das Wort unter dem Cursor zurückgegeben
    function Get-SelectedWord
    {
        $text = $psIse.CurrentFile.Editor.SelectedText
        if(!$text)
        {
            $curLine = $psIse.CurrentFile.Editor.CaretLine
            $curColumn = $psIse.CurrentFile.Editor.CaretColumn
            $curLineLength = $psIse.CurrentFile.Editor.CaretLineText.Length
            
            $curIndexLeft = $curColumn
            $curIndexRight = $curColumn
            ## to the left
            while($curIndexLeft -gt 1)
            {               
                $psIse.CurrentFile.Editor.Select($curLine, $curIndexLeft-1,$curLine,$curIndexLeft)
                $curChar = $psIse.CurrentFile.Editor.SelectedText
                if($curChar -match "[\s,.*(){}""]") {break}
                $curIndexLeft--       
            }
            ## to the right
            while($curIndexRight -lt $curLineLength+1)
            {                  
                $psIse.CurrentFile.Editor.Select($curLine, $curIndexRight,$curLine,$curIndexRight+1)
                $curChar = $psIse.CurrentFile.Editor.SelectedText
                if($curChar -match "[\s,.*(){}""]") {break}
                $curIndexRight++       
            }
            
            ## get selection        
            $psIse.CurrentFile.Editor.Select($curLine, $curIndexLeft,$curLine,$curIndexRight)
            $text = $psIse.CurrentFile.Editor.SelectedText
            
            $psIse.CurrentFile.Editor.Select($curLine,$curColumn,$curLine,$curColumn)
        }
        return $text
    }
    if($full)
    {
        Search-MSDN (Get-SelectedWord) -full
    }
    else
    {
        Search-MSDN (Get-SelectedWord)     
    }
}


# Auswahl auskommentieren

function Comment-Selected
{
    if(($sel_Text = $psIse.CurrentFile.Editor.SelectedText))
    {
        $text = $psISE.CurrentFile.Editor.Text
        $sel_text = $sel_Text.Insert(0,"<#")
        $sel_text = $sel_Text.Insert($sel_Text.Length,"#>")
        $psIse.CurrentFile.Editor.InsertText($sel_Text)
        
    }
}


# Verändert den Prompt in der ISE dahingehend, dass angezeigt wird ob der User Admin ist.
function prompt
{
    $wid =[System.Security.Principal.WindowsIdentity]::GetCurrent()
    $prp = New-Object System.Security.Principal.WindowsPrincipal($wid)
    $adm = [System.Security.Principal.WindowsBuiltInRole]::Administrator
    $isAdmin = $prp.IsInRole($adm)
    if($isAdmin)
    {
        Write-Host '[ADMIN]' -ForegroundColor Red -NoNewline
    }
    else
    {
        Write-Host '[NOADMIN]' -ForegroundColor Green -NoNewline
    }
    $host.UI.RawUI.WindowTitle = $wid.name
    'PS '+$(Get-Location)+'> '      
}

# Get-Constructor
filter Get-Constructor
{
    if($_)
    {
        $type = $_        
        foreach($constructor in $type.GetConstructors())
        {
            $params = $constructor.GetParameters()
            ($params | foreach{$_.ToString() } ) -join ', '       
        }
    }
}


## Format-Code
function Format-Code
{
    [CmdletBinding()]
    Param(
    [switch] $Alternate = $false
    )
    
    $text = $psIse.CurrentFile.Editor.Text
    
    # Splitten aber die Trennzeichen lassen
    $textlines = $text -split '(?<=\n)'
    
    # Führende Spaces / Tabs entfernen
    for($i=0;$i -lt $textlines.Count;$i++)
    {
        $textlines[$i] = $textlines[$i] -replace('^(\s+)([^\s])','$2')
    }
    
    # Tabs einfügen
    $tab = "";
    
    if(!$Alternate)
    {
        for($i=0;$i -lt $textlines.Count;$i++)
        {    
            if(!($textlines[$i] -match '^.*\{.*\}') -and !($textlines[$i] -match '^#.*'))
            {
                if($textlines[$i] -match '\{\s+$') { $textlines[$i] = $tab + $textlines[$i]; $tab += ' '*4 ; continue}    
                if($textlines[$i] -match '^\}.*$')  { if($tab.Length -gt 0) {$tab = $tab.Substring(0,$tab.Length-4)}} #Testweise ^ eingefügt
            }
            #keine "@ und @" einrücken da letzteres von Powershell nur am Zeilenanfang akzeptiert wird
            if(!($textlines[$i] -match '^"@') -and !($textlines[$i] -match '^@"') -and !($textlines[$i] -match '^''@') -and !($textlines[$i] -match '^@'''))
            {
                $textlines[$i] = $tab + $textlines[$i]
            }  
            
        }
    } 
    else
    {
        for($i=0;$i -lt $textlines.Count;$i++)
        {    
            if(!($textlines[$i] -match '^.*\{.*\}') -and !($textlines[$i] -match '^#.*'))
            {             
                if($textlines[$i] -match '^\{\s+$')  { $tab += ' '*4 ; $textlines[$i] = $tab + $textlines[$i];  continue}
                if($textlines[$i] -match '\{\s+$')   { $textlines[$i] = $tab + $textlines[$i]; $tab += ' '*4 ;  continue}             
                if($textlines[$i] -match '^\}.*$')    { $textlines[$i] = $tab + $textlines[$i]; if($tab.Length -gt 0) {$tab = $tab.Substring(0,$tab.Length-4)};continue} #Testweise ^
            }
            #keine "@ und @" einrücken da letzteres von Powershell nur am Zeilenanfang akzeptiert wird
            if(!($textlines[$i] -match '^"@') -and !($textlines[$i] -match '^@"') -and !($textlines[$i] -match '^''@') -and !($textlines[$i] -match '^@'''))
            {
                $textlines[$i] = $tab + $textlines[$i]
            }
        }
    }
    
    
    ## Blast it back...
    $caretline = $psIse.CurrentFile.Editor.CaretLine
    $text1 = $textlines -join ""
    $psIse.CurrentFile.Editor.Text = $text1
    $psIse.CurrentFile.Editor.SetCaretPosition(1,1)
    
    # Die Zeile in welcher sich der Cursor befindet, möglichst nicht am unteren Rand positionieren 
    if(($caretline+15) -le $psIse.CurrentFile.Editor.LineCount) 
    {
        $psIse.CurrentFile.Editor.SetCaretPosition($caretline+15,1)
    }
    else
    {
        $psIse.CurrentFile.Editor.SetCaretPosition($psIse.CurrentFile.Editor.LineCount,1) 
    }
    
    $psIse.CurrentFile.Editor.SetCaretPosition($caretline,1)
}



#Variablen des Scriptes welches im Editor geöffnet ist
function Get-ScriptVars
{
    $obj  = New-Object System.Collections.ArrayList
    $text = $psIse.CurrentFile.Editor.Text    
    $vars = ([regex]'(?i)\$[a-z0-9]+').Matches($text).value |Sort -Unique    
    foreach($var in $vars)
    {
        try
        {
            $value = Get-Variable $var.remove(0,1) -ValueOnly -ErrorAction SilentlyContinue
            $obj  += [PSCustomObject]@{Name = "$var";Inhalt = "$value";DataType = "$($value.GetType().Name)"}
        }
        catch
        {
            $obj  += [PSCustomObject]@{Name = "$var";Value = "unbekannt";DataType = "unbekannt"}
        }       
    }
    return $obj
}
# Loaded-Assembly-Browser
function Get-AssemblyBrowser
{
    [AppDomain]::CurrentDomain.GetAssemblies() |
    ForEach-Object { try {$_.GetExportedTypes() } catch {} } |
    Where-Object { $_.isPublic } |
    Where-Object { $_.isClass  } |
    Where-Object { $_.Name -notmatch '(Attribute|Handler|Args|Exception|Collection|Expression|Parser|Statement)$'} |
    
    Select-Object -Property Name,Fullname | sort -Property Name |
    Out-GridView   
}

# FunctionExplorer
# Author: Dennis Bergemann

function Show-FunctionExplorer
{
    #Eingebettete Funktion von aussen nicht sichtbar
    function Colorize-Code($textBox,$code)
    {            
        
        #keywords
        $type = New-Object "Collections.Generic.Dictionary[String, [Drawing.Color]]"
        $type["Attribute"]          = [Drawing.Color]::LightSteelBlue       
        $type["Command"]            = [Drawing.Color]::LightCyan
        $type["Comment"]            = [Drawing.Color]::PaleGreen
        $type["GroupStart"]         = [Drawing.Color]::WhiteSmoke
        $type["GroupEnd"]           = [Drawing.Color]::WhiteSmoke
        $type["Keyword"]            = [Drawing.Color]::LightCyan
        $type["Member"]             = [Drawing.Color]::WhiteSmoke
        $type["Operator"]           = [Drawing.Color]::LightGray
        $type["String"]             = [Drawing.Color]::PaleVioletRed
        $type["Type"]               = [Drawing.Color]::DarkSeaGreen
        $type["Variable"]           = [Drawing.Color]::OrangeRed
        $type["Number"]             = [Drawing.Color]::Bisque
        $type["CommandArgument"]    = [Drawing.Color]::Violet
        $type["CommandParameter"]   = [Drawing.Color]::Moccasin         
        $type["StatementSeparator"] = [Drawing.Color]::WhiteSmoke             
        #counter
        $i = 0
        #highlight code
        [Management.Automation.PSParser]::Tokenize($code, [ref]$null) |
        ForEach-Object {
            
            #Token auswählen
            $textBox.SelectionStart = $_.Start - $i
            $textBox.SelectionLength = $_.Length
            
            #Auswahl bei Mehrzeilentoken anpassen            
            $i += $_.Endline - $_.Startline
            
            #Auswahl färben
            if ($type[$_.Type.ToString()] -is [System.Drawing.Color])
            {          
                
                #Infos als Objekt liegen lassen
                #[PSCustomObject]  @{Content = $_.Content; Type = $_.Type}
                $textBox.SelectionColor = $type[$_.Type.ToString()]
            }
            
        }
        #Auswahl aufheben
        $textBox.DeselectAll()                  
        
    }
    
    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing    
    
    [System.Windows.Forms.Application]::EnableVisualStyles()
    
    
    
    # Form erstellen
    $form = New-Object System.Windows.Forms.Form
    $form.Text = "Function-Explorer"
    $form.Size = New-Object System.Drawing.Size(960,600) 
    $form.StartPosition = "CenterScreen"
    $form.SizeGripStyle = "Hide"
    $form.Icon = [System.Drawing.SystemIcons]::Information
    
    $tableLayoutPanel = New-Object System.Windows.Forms.TableLayoutPanel
    $textBox = New-Object System.Windows.Forms.RichTextBox
    $listBox = New-Object System.Windows.Forms.ListBox
    
    
    # Komponenten initialisieren
    # TableLayoutPanel
    $tableLayoutPanel.ColumnCount = 2
    $tableLayoutPanel.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 37.62828))) | Out-Null
    $tableLayoutPanel.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 62.37172))) | Out-Null
    $tableLayoutPanel.Controls.Add($textBox, 1, 0)
    $tableLayoutPanel.Controls.Add($listBox, 0, 0)
    $tableLayoutPanel.Dock = [System.Windows.Forms.Dockstyle]::Fill
    $tableLayoutPanel.Location = New-Object System.Drawing.Point(0, 0)
    $tableLayoutPanel.Name = "tableLayoutPanel"
    $tableLayoutPanel.RowCount = 1
    $tableLayoutPanel.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 50)))|Out-Null
    $tableLayoutPanel.Size = New-Object System.Drawing.Size(877, 458)
    $tableLayoutPanel.TabIndex = 0
    
    # TextBox
    $textBox.BackColor = [Drawing.Color]::FromArgb(42, 42, 42)
    $textBox.ForeColor = [Drawing.Color]::White
    $textBox.Font = New-Object Drawing.Font("Lucida Console", 9, [Drawing.FontStyle]::Regular)
    $textBox.Dock = [System.Windows.Forms.DockStyle]::Fill
    $textBox.Enabled = $true
    $textBox.Location = New-Object System.Drawing.Point(332, 3)
    $textBox.Multiline = $true
    $textBox.Name = "textBox"
    $textBox.ScrollBars = [System.Windows.Forms.ScrollBars]::Both
    $textBox.WordWrap = $false
    $textBox.Size = New-Object System.Drawing.Size(542, 452)
    $textBox.TabIndex = 1
    
    #listBox
    $listbox.BackColor = [System.Drawing.Color]::FromArgb(42,42,42)
    $listbox.ForeColor = [System.Drawing.Color]::White    
    $listBox.Dock = [System.Windows.Forms.DockStyle]::Fill
    $listBox.FormattingEnabled = $true    
    $listBox.Location = New-Object System.Drawing.Point(3, 3)
    $listBox.Name = "listBox"
    $listBox.Size = New-Object System.Drawing.Size(323, 452)
    $listBox.Font = New-Object Drawing.Font("Lucida Console", 10, [Drawing.FontStyle]::Regular)
    $listBox.TabIndex = 0
    $listBox.DrawMode = [System.Windows.Forms.DrawMode]::OwnerDrawFixed    
    
    # OwnerDraw
    # Auszuführender Code bei DrawItem
    $DrawItem = 
    {
        $_.DrawBackground();
        
        [System.Drawing.Graphics] $g = $_.Graphics;
        if (($_.State -and [System.Windows.Forms.DrawItemState]::Selected) -eq [System.Windows.Forms.DrawItemState]::Selected)
        {
            $brush = [System.Drawing.Brushes]::Orange    
        }
        else
        {
            $brush = New-Object System.Drawing.SolidBrush($_.BackColor)
        }    
        
        $g.FillRectangle($brush, $_.Bounds);
        $_.Graphics.DrawString($listBox.Items[$_.Index].ToString(), $_.Font, (New-Object System.Drawing.SolidBrush($_.ForeColor)), $_.Bounds.x,$_.Bounds.y) 
        $_.DrawFocusRectangle();     
    } 
    $SelIndexChanged =
    {
        $textBox.Text = $functions[$listBox.SelectedIndex].Extent.Text 
        $code = $functions[$listBox.SelectedIndex].Extent.Text
        Colorize-Code -textBox $textBox -code $code
    }
    #Add it    
    $listBox.Add_DrawItem($DrawItem)                         
    
    # File parsen und Functionen extrahieren
    $ast = [System.Management.Automation.Language.Parser]::ParseInput($psISE.CurrentFile.Editor.Text, [ref]$null, [ref]$null)    
    $functions = $ast.FindAll( {$args[0] -is [System.Management.Automation.Language.FunctionDefinitionAst]}, $true)
    
    if($functions)
    {                                    
        # Liste mit Funktionsnamen füllen
        $i=0    
        foreach($function in $functions)
        {
            $listBox.Items.Add("$($function.Name) ($($function.Parameters))") | Out-Null
        }
        
        # Handler an die Liste hängen
        $listBox.Add_SelectedIndexChanged($SelIndexChanged)
        
        # Erstes Item anzeigen       
        if($functions)
        { 
            $textBox.Text = $functions[0].Extent.Text
            $code = $functions[0].Extent.Text
            Colorize-Code -textBox $textBox -code $code
        }
        
        $form.Controls.Add($tableLayoutPanel)
        
        # Komponenten freigeben
        
        #Form anzeigen
        $form.Topmost = $True
        #$result = $form.ShowDialog()
        [System.Windows.Forms.Application]::Run($form)
        $listBox.Remove_DrawItem($DrawItem)
        $listBox.Remove_SelectedIndexChanged($SelIndexChanged)
        $form = $null
    }
    else
    {
        [System.Windows.Forms.MessageBox]::Show("Keine Funktionen vorhanden", "Warnung", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Asterisk)|Out-Null     
    }                                            
}

function Remove-ISEAlias
{
    $text =$psISE.CurrentFile.Editor.Text
    #Stringbuilder geht schneller
    $sb = New-Object System.Text.StringBuilder $text
    #Letzter Befehl zuerst (sonst verändern sich die Findepositionen)
    $befehle = [System.Management.Automation.PSParser]::Tokenize($text,[ref]$null) | Where-Object {$_.Type -eq 'Command' } | Sort-Object -Property Start -Descending |
    ForEach-Object {
        
        $befehl = $text.Substring($_.Start,$_.Length)        
        $befehlstyp = @(try {Get-Command $befehl -ErrorAction Ignore} catch{})[0]
        if ($befehlstyp -is [System.Management.Automation.AliasInfo])
        {
            $sb.Remove($_.Start, $_.Length)
            $sb.Insert($_.Start, $befehlstyp.ResolvedCommandName)            
        }
        
        # Aktualisiert zum Editor zurück
        $psISE.CurrentFile.Editor.Text = $sb.ToString()
    }    
}

## Snippet-Browser
## Author: Dennis Bergemann

function Show-SnippetBrowser
{
    #Eingebettete Funktion von aussen nicht sichtbar
    function Colorize-Code($textBox,$code)
    {            
        
        #keywords
        $type = New-Object "Collections.Generic.Dictionary[String, [Drawing.Color]]"
        $type["Attribute"]          = [Drawing.Color]::LightSteelBlue       
        $type["Command"]            = [Drawing.Color]::LightCyan
        $type["Comment"]            = [Drawing.Color]::PaleGreen
        $type["GroupStart"]         = [Drawing.Color]::WhiteSmoke
        $type["GroupEnd"]           = [Drawing.Color]::WhiteSmoke
        $type["Keyword"]            = [Drawing.Color]::LightCyan
        $type["Member"]             = [Drawing.Color]::WhiteSmoke
        $type["Operator"]           = [Drawing.Color]::LightGray
        $type["String"]             = [Drawing.Color]::PaleVioletRed
        $type["Type"]               = [Drawing.Color]::DarkSeaGreen
        $type["Variable"]           = [Drawing.Color]::OrangeRed
        $type["Number"]             = [Drawing.Color]::Bisque
        $type["CommandArgument"]    = [Drawing.Color]::Violet
        $type["CommandParameter"]   = [Drawing.Color]::Moccasin         
        $type["StatementSeparator"] = [Drawing.Color]::WhiteSmoke             
        #counter
        $i = 0
        #highlight code
        [Management.Automation.PSParser]::Tokenize($code, [ref]$null) |
        ForEach-Object {
            
            #Token auswählen
            $textBox.SelectionStart = $_.Start - $i
            $textBox.SelectionLength = $_.Length
            
            #Auswahl bei Mehrzeilentoken anpassen            
            $i += $_.Endline - $_.Startline
            
            #Auswahl färben
            if ($type[$_.Type.ToString()] -is [System.Drawing.Color])
            {          
                
                #Infos als Objekt liegen lassen
                #[PSCustomObject]  @{Content = $_.Content; Type = $_.Type}
                $textBox.SelectionColor = $type[$_.Type.ToString()]
            }
            
        }
        #Auswahl aufheben
        $textBox.DeselectAll()                  
        
    }
    
    
    $snippets = $psIse.CurrentPowerShellTab.Snippets
    
    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing
    
    
    [System.Windows.Forms.Application]::EnableVisualStyles()
    
    
    # create the form
    $form = New-Object System.Windows.Forms.Form
    $form.Text = "Snippet-Browser"
    $form.Size = New-Object System.Drawing.Size(700,700) 
    $form.StartPosition = "CenterScreen"
    $form.SizeGripStyle = "Hide"
    $form.Icon = [System.Drawing.SystemIcons]::Information
    #TabControl1
    $tabControl1 = New-Object System.Windows.Forms.TabControl
    $tabControl1.Font = New-Object System.Drawing.Font("Arial", 9, [System.Drawing.FontStyle]::Regular, [System.Drawing.GraphicsUnit]::Point, 0);
    $tabControl1.Dock = [System.Windows.Forms.DockStyle]::Fill
    $tabControl1.Name = "tabControl1"
    $tabControl1.SelectedIndex = 0
    $tabControl1.TabIndex = 0;
    
    $i=0
    foreach($snippet in $snippets | Sort-Object -Property DisplayTitle )
    {    
        # für jedes snippet einen Tab
        $tabpage = New-Object System.Windows.Forms.TabPage   
        $tabPage.Name = "tabPage$($i)"
        $tabPage.Padding = New-Object System.Windows.Forms.Padding(3)   
        $tabPage.TabIndex = $i
        $tabPage.Text = $snippet.DisplayTitle
        $tabPage.UseVisualStyleBackColor = $true
        $tabControl1.Controls.Add($tabPage)
        
        # in jedem Tab ein Textfeld mit dem Snippet
        $textBox = New-Object System.Windows.Forms.RichTextBox
        $textBox.Dock = [System.Windows.Forms.DockStyle]::Fill    
        $textBox.Font = New-Object Drawing.Font("Lucida Console", 9, [Drawing.FontStyle]::Regular)
        $textBox.Multiline = $true;
        $textBox.Name = "textBox$($i)";
        $textBox.ScrollBars = [System.Windows.Forms.ScrollBars]::Both;
        $textBox.TabIndex = $i;
        $textbox.Text = $snippet.CodeFragment
        $textbox.BackColor = [System.Drawing.Color]::FromArgb(42,42,42)
        $textbox.ForeColor = [System.Drawing.Color]::White
        $tabpage.Controls.Add($textBox)
        Colorize-Code -textBox $textBox -code $snippet.CodeFragment        
        $i+=1
    }
    
    $form.Controls.Add($tabControl1)
    
    #Form anzeigen
    $form.Topmost = $True
    <#$result = $form.ShowDialog()#>
    [System.Windows.Forms.Application]::Run($form)
    $form = $null
}

# Markierter Text wird mit Codeblock umschlossen
function Recode-Selected
{
    [cmdletbinding()]    
    param
    (
    [ValidateSet('TryCatch','TryCatchLog','DoWhile','While','For','ForEach','PipeForEach')]
    [String] $Codetyp
    )
    if($selected_Code  = $psIse.CurrentFile.Editor.SelectedText)
    {
        $code = New-Object System.Text.StringBuilder
        switch($Codetyp)
        {
            'TryCatch'
            {
                $code.AppendLine("Try") | out-null
                $code.AppendLine("{") | out-null
                $code.AppendLine($selected_Code) | out-null
                $code.AppendLine("}") | out-null
                $code.AppendLine("Catch") | out-null
                $code.AppendLine("{") | out-null
                $code.AppendLine("}") | out-null
                $code.AppendLine("Finally") | out-null
                $code.AppendLine("{") | out-null
                $code.AppendLine("}") | out-null                                                           
            }      
            'TryCatchLog'
            {
                $code.AppendLine('Log -Logmessage "ACTION -- "') | out-null
                $code.AppendLine("Try") | out-null
                $code.AppendLine("{") | out-null
                $code.AppendLine($selected_Code) | out-null
                $code.AppendLine("}") | out-null
                $code.AppendLine("Catch") | out-null
                $code.AppendLine("{") | out-null
                $code.AppendLine('Log -Logmessage "ERROR --  $($_.Exception.Message)"') | out-null
                $code.AppendLine('$bug = $true') | out-null
                $code.AppendLine("}") | out-null
                $code.AppendLine("Finally") | out-null
                $code.AppendLine("{") | out-null
                $code.AppendLine('If ($bug -ne $true)') | out-null                
                $code.AppendLine("{") | out-null
                $code.AppendLine('Log -Logmessage "RESULT -- "') | out-null
                $code.AppendLine("}") | out-null
                $code.AppendLine("}") | out-null
                $code.AppendLine('$bug = $null')  | out-null                                           
            }            
            'DoWhile'
            {
                $code.AppendLine("Do") | out-null
                $code.AppendLine("{") | out-null
                $code.AppendLine($selected_Code) | out-null
                $code.AppendLine("}") | out-null
                $code.AppendLine("While()") | out-null
            }
            'While'
            {
                $code.AppendLine("While()") | out-null
                $code.AppendLine("{") | out-null
                $code.AppendLine($selected_Code) | out-null
                $code.AppendLine("}") | out-null                
            } 
            'For'
            {
                $code.AppendLine("For()") | out-null
                $code.AppendLine("{") | out-null
                $code.AppendLine($selected_Code) | out-null
                $code.AppendLine("}") | out-null
            } 
            'ForEach'
            {
                $code.AppendLine("ForEach()") | out-null
                $code.AppendLine("{") | out-null
                $code.AppendLine($selected_Code) | out-null
                $code.AppendLine("}") | out-null
            }
            'PipeForEach'
            {
                $code.AppendLine(" | ForEach {") | out-null
                $code.AppendLine($selected_Code) | out-null
                $code.AppendLine("}") | out-null
            }          
        }   
        $psIse.CurrentFile.Editor.InsertText($code.ToString())
        $psIse.CurrentFile.Editor.SetCaretPosition($psIse.CurrentFile.Editor.CaretLine,$psIse.CurrentFile.Editor.CaretColumn)
    }
}


# Fügt im Aktuellen Editor(Script) an der Cursorposition "Templatecode" ein
function Insert-Code
{
    [cmdletbinding()]
    param(
    [ValidateSet('Function')]
    [string] $Codeblock
    )
    switch($Codeblock)
    {
        'Function' 
        {
            $code = "function Name" +[Environment]::NewLine + "{" + [Environment]::NewLine + "[CmdletBinding()]" + [Environment]::NewLine + "param(" + [Environment]::NewLine + ")" + [Environment]::NewLine + "}" 
            $psIse.CurrentFile.Editor.InsertText($code)
        }
        
    }    
}

# Simple-Script-Tabber 
function Retab-Script
{
    [CmdletBinding()]
    param(
    [string] $Code = $psISE.CurrentFile.Editor.Text,
    [int] $TabWidth = 4
    )
    
    
    $CurrentLevel = 0
    $Tokens = $ParseError = $null	
    $AST = [System.Management.Automation.Language.Parser]::ParseInput($Code, [ref]$Tokens, [ref]$ParseError) 
    $sbCode = New-Object System.Text.StringBuilder $code	
    
    if($ParseError) 
    { 
        $ParseError | Write-Error
        throw "The parser will not work properly with errors in the script, please modify based on the above errors and retry."
    }
    
    for($index = $Tokens.Count -2; $index -ge 1; $index--)
    {
        
        $Token = $Tokens[$index]
        $NextToken = $Tokens[$index-1]
        
        if ($token.Kind -match '(L|At)Curly')
        { 
            if($CurrentLevel -gt 0){$CurrentLevel--} 
        }  
        
        if ($NextToken.Kind -eq 'NewLine' ) 
        {
            # Grab Placeholders for the Space Between the New Line and the next token.
            $RemoveStart = $NextToken.Extent.EndOffset  
            $RemoveEnd = $Token.Extent.StartOffset - $RemoveStart
            $tabText = <#"`t"#>' ' * 4 * $CurrentLevel 
            $sbCode = $sbCode.Remove($RemoveStart,$RemoveEnd).Insert($RemoveStart,$tabText)
        }		
        
        if ($token.Kind -eq 'RCurly')
        { 
            $CurrentLevel++
            
        }     
    }
    
    
    ## Blast it back...
    $caretline = $psIse.CurrentFile.Editor.CaretLine
    
    $psIse.CurrentFile.Editor.Text = $sbCode.ToString()
    $psIse.CurrentFile.Editor.SetCaretPosition(1,1)
    
    # Die Zeile in welcher sich der Cursor befindet, möglichst nicht am unteren Rand positionieren 
    if(($caretline+15) -le $psIse.CurrentFile.Editor.LineCount) 
    {
        $psIse.CurrentFile.Editor.SetCaretPosition($caretline+15,1)
    }
    else
    {
        $psIse.CurrentFile.Editor.SetCaretPosition($psIse.CurrentFile.Editor.LineCount,1) 
    }
    
    $psIse.CurrentFile.Editor.SetCaretPosition($caretline,1)
}

# Variablechecker
function Adjust-Variables
{
    [CmdletBinding()]
    param(
    [string] $Code = $psISE.CurrentFile.Editor.Text
    )
    $obj = @{}         
    #$text = $psIse.CurrentFile.Editor.Text
    $sb = New-Object System.Text.StringBuilder $Code    
    
    # alle Variablen im Text suchen. (?i) stellt auf ignore-Case 
    $vars = ([regex]'(?i)\$[a-z0-9_]+').Matches($Code) 
    
    foreach($var in $vars)
    {
        try
        {
            $obj.add($var.value,$var.value)
        }
        catch
        {
            $sb.Remove($var.Index,$var.Length) | Out-Null                                
            $sb.Insert($var.Index,$obj[$var.value]) | Out-Null
        }               
        
    }
    #$psIse.CurrentFile.Editor.Text = $sb.ToString()
    #$sb.ToString()
    ## Blast it back...
    $caretline = $psIse.CurrentFile.Editor.CaretLine
    
    $psIse.CurrentFile.Editor.Text = $sb.ToString()
    $psIse.CurrentFile.Editor.SetCaretPosition(1,1)
    
    # Die Zeile in welcher sich der Cursor befindet, möglichst nicht am unteren Rand positionieren 
    if(($caretline+15) -le $psIse.CurrentFile.Editor.LineCount) 
    {
        $psIse.CurrentFile.Editor.SetCaretPosition($caretline+15,1)
    }
    else
    {
        $psIse.CurrentFile.Editor.SetCaretPosition($psIse.CurrentFile.Editor.LineCount,1) 
    }
    
    $psIse.CurrentFile.Editor.SetCaretPosition($caretline,1)
}


### Menüeinträge

function Add-PMenu
{
    
    $psISE.CurrentPowerShellTab.AddOnsMenu.Submenus.Add("Search-MSDN(Powershell Command Only)", {Search-Selected}, "Shift+F1") | Out-Null
    $psISE.CurrentPowerShellTab.AddOnsMenu.Submenus.Add("Search-MSDN(Full)", {Search-Selected -full}, "Alt+F1") | Out-Null
    
    $psISE.CurrentPowerShellTab.AddOnsMenu.Submenus.Add("ISE Function Explorer", {Show-FunctionExplorer}, "Ctrl+Alt+X") | Out-Null
    $psIse.CurrentPowerShellTab.AddOnsMenu.Submenus.Add("Loaded-Assembly-Browser",{Get-AssemblyBrowser},"Ctrl+Alt+A") | Out-Null
    $psIse.CurrentPowerShellTab.AddOnsMenu.Submenus.Add("Remove-ISEAlias",{Remove-ISEAlias},"Ctrl+Alt+R") | Out-Null
    $psIse.CurrentPowerShellTab.AddOnsMenu.Submenus.Add("SnippetBrowser",{Show-SnippetBrowser},"Ctrl+Alt+S") | Out-Null
    $psIse.CurrentPowerShellTab.AddOnsMenu.Submenus.Add("Comment-Selected",{Comment-Selected},"Ctrl+K") | Out-Null
    $psIse.CurrentPowerShellTab.AddOnsMenu.Submenus.Add("Adjust Variables",{Adjust-Variables},"Ctrl+Alt+V") | Out-Null
    
    # Format Code Menu
    $formatmenu = $psIse.CurrentPowerShellTab.AddOnsMenu.Submenus.Add("Format-Code",$null,$null) 
    $formatmenu.SubMenus.Add("Version_1",{Retab-Script},"F7") | Out-Null
    $formatmenu.SubMenus.Add("Version_2",{Format-Code -Alternate},"Shift+F7") | Out-Null
    
    # Recode-Selected Menu
    $formatmenu = $psIse.CurrentPowerShellTab.AddOnsMenu.Submenus.Add("Recode-Selected",$null,$null) 
    $formatmenu.SubMenus.Add("TryCatch",{Recode-Selected -Codetyp TryCatch},"Ctrl+Alt+T") | Out-Null    
    $formatmenu.SubMenus.Add("TryCatchLog",{Recode-Selected -Codetyp TryCatchLog},"Ctrl+Alt+E") | Out-Null
    $formatmenu.SubMenus.Add("DoWhile",{Recode-Selected -Codetyp DoWhile},"Ctrl+Alt+D") | Out-Null
    $formatmenu.SubMenus.Add("While",{Recode-Selected -Codetyp While},"Ctrl+Alt+W") | Out-Null
    $formatmenu.SubMenus.Add("For",{Recode-Selected -Codetyp For},"Ctrl+Alt+1")|out-null
    $formatmenu.SubMenus.Add("ForEach",{Recode-Selected -Codetyp ForEach},"Ctrl+Alt+2")|out-null
    $formatmenu.SubMenus.Add("PipeForEach",{Recode-Selected -Codetyp PipeForEach},"Ctrl+Alt+3")|out-null


    
    # Insert-Code                    
    $formatmenu = $psIse.CurrentPowerShellTab.AddOnsMenu.Submenus.Add("Insert-Code",$null,$null) 
    $formatmenu.SubMenus.Add("Function",{Insert-Code -Codeblock Function},"Ctrl+Shift+F") | Out-Null
    
}

## add it
Add-PMenu

## AddTools
add-type -path "$PSScriptRoot\ADD_DLL\RegexTool.dll"
$psISE.CurrentPowerShellTab.VerticalAddOnTools.Add(‘Regex Helper’, [RegexTool.RegexToolWindow], $false) | Out-null