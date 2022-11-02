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
# SIG # Begin signature block
# MIIj3wYJKoZIhvcNAQcCoIIj0DCCI8wCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCA34noHJmto+I3C
# 0DxxVsQrLTmVgGdfOktJRLREal2TwaCCHgswggUyMIIDGqADAgECAgIQADANBgkq
# hkiG9w0BAQsFADBOMQswCQYDVQQGEwJMVTEmMCQGA1UECgwdTWFqb3JlbCBHcm91
# cCBMdXhlbWJvdXJnIFMuQS4xFzAVBgNVBAMMDk1ham9yZWwgQ0EgSU0xMB4XDTIy
# MDUwNjE0MzkxNloXDTIzMDUxNjE0MzkxNlowdTELMAkGA1UEBhMCREUxEDAOBgNV
# BAoMB01ham9yZWwxCzAJBgNVBAsMAklUMRowGAYDVQQDDBFCZXJnZW1hbm4sIERl
# bm5pczErMCkGCSqGSIb3DQEJARYcZGVubmlzLmJlcmdlbWFubkBtYWpvcmVsLmNv
# bTCCASIwDQYJKoZIhvcNAQEBBQADggEPADCCAQoCggEBAM7La3acKk/NlQB9LKEC
# trRJEoRDn5Xiazp1lo2fCufO0Ip0MA0oOD+jTpHLzhI319JJwQYW4evacyw+hTXc
# z5J5SiRwsbEfUB8sJPn/TbD2ceFx537BvjuIKDepjEi9SkuBkT8FChgOO3RpEHWH
# uf4XgRfpf/K+C3ZVv+1EWmTdsPw1wL3rE4/7ddRAhMeGITn0W64XGWizZtrghm0J
# fRj2zYs5A/jMG045EW+WwyNornZ9L3QDLH2zSJ80/TmNFoK4Lw/lvTZBO0aZLHFp
# nGZBFodgaBdptg933K2wbolQEB21LPMhGKgNYVqYwhK1wRylhpDTz9+dXivTfXUn
# 8ZMCAwEAAaOB8jCB7zAJBgNVHRMEAjAAMBEGCWCGSAGG+EIBAQQEAwIEEDA5Bglg
# hkgBhvhCAQ0ELBYqT3BlblNTTCBHZW5lcmF0ZWQgQ29kZSBTaWduaW5nIENlcnRp
# ZmljYXRlMB0GA1UdDgQWBBS0G+Kk5fwugDe7BEioOUbnG/18LTAfBgNVHSMEGDAW
# gBRXXCJ+a3JZbjbTzjSMflV9Llwb9DAOBgNVHQ8BAf8EBAMCBsAwRAYDVR0lAQH/
# BDowOAYIKwYBBQUHAwMGCisGAQQBgjcCARUGCisGAQQBgjcCARYGCisGAQQBgjcK
# AwEGCCsGAQUFBwMIMA0GCSqGSIb3DQEBCwUAA4ICAQA0em0HG75Rd+6tEg1LucmB
# hWkFjVLq7Lim62IiEIyk7+njH+JpZyFfLFcWMhzHR94Ve3bGlbRhp4fsJgfNgegs
# Ytkh0hZGTSuezHktARDsX0TBVbaBaBB0RHWyCfc9ep6Ey+BFmlB3++DZm/tR7V8b
# JCxjylLNGa0fRUEvqR+fTh4PtRgV+mgUFYBNAw5xEeTmywCnkVTIUUt3OaagJHhc
# Y0GEbarKbb2RHAKafGwlDNSh3+7mw939G30+EgkChwWQzNYVIlY/C9QFZcMFGPZP
# NAy+k+Sg72h+fe7T92bbt/2qEKpQrz49VIhpjKEtbrBJ88vRm0MUIPUCPpIXmc0O
# O+10eHTOV/cK20fCSX+E0FZSH8sLN2nimX+EzaTYsUcAsM5mK1gd6RhFvr5+myV6
# WaAUi1k7140VDkITVz8E8Bh3aU5o9f1vy2Hj0hG0lfXHqLjFo+dORTTV6pXfBbn8
# ZlLBCZXu2m9ZZkrLn/URe77UMHYI1yMzOTPwRo8O6oj2wHrVhpw3tDZhKiTr9XVd
# FRcu0dBT+IR+6P9NhzkmUtm/qZ1mqr5pwlyPHvpmqkHg02R4kUdu30E9bgblYiJ1
# vnZ6yu2h8Gq0W9h3Qd4qF5fPoXK4uj4KHXj8wVRyqvib/4JG5pe7Cswr9swkl531
# 9vvFoFM7Bb+xap47ZRVYwTCCBaAwggOIoAMCAQICAhAAMA0GCSqGSIb3DQEBCwUA
# MHAxCzAJBgNVBAYTAkxVMSYwJAYDVQQKDB1NYWpvcmVsIEdyb3VwIEx1eGVtYm91
# cmcgUy5BLjEYMBYGA1UECwwPd3d3Lm1ham9yZWwuY29tMR8wHQYDVQQDDBZNYWpv
# cmVsIEdsb2JhbCBSb290IENBMB4XDTIyMDUwNTE2MTY0NVoXDTMyMDUwMjE2MTY0
# NVowTjELMAkGA1UEBhMCTFUxJjAkBgNVBAoMHU1ham9yZWwgR3JvdXAgTHV4ZW1i
# b3VyZyBTLkEuMRcwFQYDVQQDDA5NYWpvcmVsIENBIElNMTCCAiIwDQYJKoZIhvcN
# AQEBBQADggIPADCCAgoCggIBALRq1JBEHHNEUt7bXbq86MwMoP86ZoXD8fJM65AK
# jj9uk6MH2WxtiDX2OqmriDpFsktMYq4AxDoVNIupDWlhQdasJU5/bVHAACdpZNcC
# VUHay2VUc9j3pYdiqgBhWFnmAgLpI2jQqxjlqKk5kFpsmVsv1773sYxdfyBkKMkX
# dy9ATBpT3jG/ZUjwJF+KQa46EtxvXRoF22oLgVJmjLSOtWA3bBVKbNRY6adbfTtT
# XXFMAn/FLYxKzUD1F3U8eBGCwCHGf///KhxMnLijjHtkQSkXySEhY7fqSfGN21mo
# XXjgwRHNFyRJE8D9JoKC45W8oXbGw/g94d3blHhbG9AnRtEgjqzfLnnEPL4AoAWp
# k5pRlho38dIjS+C8+clIAzbEfDBfOHsB1+jwOucjV6sq7rZ1YfpJME5ZsLZw0Gdh
# EHiAV4CHrEXJO1mReCot6bjPbkhubR4L7RF33Wmdv1qULcWsCPkWbqIe6opMxagF
# SSpdaMdT1K70fVwoHczocAaklmqnWn892SwhxKuwcavB8doDlrjahydpzM1eB0u6
# KTS9cLELmYc/F3Mm8gNfSeXSo3JS60sLcVYNqRO8ToItCWlqcnfM2+qV9IzadgZE
# eSv4N79J538WSY30o2m6iORF4S8fQPTIFWmMCM1yg3jC66NVDQApMtCvp0/mvteM
# Fw/dAgMBAAGjZjBkMB0GA1UdDgQWBBRXXCJ+a3JZbjbTzjSMflV9Llwb9DAfBgNV
# HSMEGDAWgBRrX51KyY1iex7xhua3XCRTkESvkzASBgNVHRMBAf8ECDAGAQH/AgEA
# MA4GA1UdDwEB/wQEAwIBhjANBgkqhkiG9w0BAQsFAAOCAgEACHgr9WqIFDDXaSLS
# ghc+CUR8nyDgcaSEVQC1asPJMqJVm3p1Do64V0us/lEpqtIIItTZRDQtuRoZVvlf
# KtciksG4U1z8QP11M2Ba+G9JF5vGUVAOWoZ2TuP8yUeXl4eibR9krI2s+CnCTNsF
# 2SLwIHPqLzclvaN9udT5oeKWCbCaeIgza32S4OqG60o3LdapB4veYq8W9jJb0Ar0
# bW/s5j5xpaC/dG94+9cl5cNLgFcrkT03ganYz7fxUuwc+TYrCCYvVjPSIV3MrYnn
# CfcOMJJFXdQt2rBNBZMXnSnNqd1iIhdh7kDJ3cKRFT3ynGbimhD5xri84/PaxMpj
# 2Yx3+IlKFmiczm+Mjz5qfbPzW3l5TrCj/2pwihpjkEFU1G2t1uFg6rwCFVd5NRgv
# HG+YNZK8HM3pjuZ/OC04uIRIbyifS+bLhebtxeAEgClV4RsFd00aSL7yMo3A8jS2
# zti0KJmwQ1pi04wkIpzlwpKMMfqv6nwxHSYiyx3PGOT1GFTRulXLOLv7kj0MU7OX
# k6FiILrNVgRn745g/cgbbTy28qLlEsXuguYxmr6gBs1IJvxG8xr94FT5easMbB1Y
# ivBmzkmcfT1iRYF9oqz1EZMWbGTMbNYREjOx2VwNO3dTRqnnY2z+BuHzBK+Mb3HD
# 4DCWjB1aAul65pBYjDwC81mcCogwggWxMIIEmaADAgECAhABJAr7HjgLihbxS3Gd
# 9NPAMA0GCSqGSIb3DQEBDAUAMGUxCzAJBgNVBAYTAlVTMRUwEwYDVQQKEwxEaWdp
# Q2VydCBJbmMxGTAXBgNVBAsTEHd3dy5kaWdpY2VydC5jb20xJDAiBgNVBAMTG0Rp
# Z2lDZXJ0IEFzc3VyZWQgSUQgUm9vdCBDQTAeFw0yMjA2MDkwMDAwMDBaFw0zMTEx
# MDkyMzU5NTlaMGIxCzAJBgNVBAYTAlVTMRUwEwYDVQQKEwxEaWdpQ2VydCBJbmMx
# GTAXBgNVBAsTEHd3dy5kaWdpY2VydC5jb20xITAfBgNVBAMTGERpZ2lDZXJ0IFRy
# dXN0ZWQgUm9vdCBHNDCCAiIwDQYJKoZIhvcNAQEBBQADggIPADCCAgoCggIBAL/m
# kHNo3rvkXUo8MCIwaTPswqclLskhPfKK2FnC4SmnPVirdprNrnsbhA3EMB/zG6Q4
# FutWxpdtHauyefLKEdLkX9YFPFIPUh/GnhWlfr6fqVcWWVVyr2iTcMKyunWZanMy
# lNEQRBAu34LzB4TmdDttceItDBvuINXJIB1jKS3O7F5OyJP4IWGbNOsFxl7sWxq8
# 68nPzaw0QF+xembud8hIqGZXV59UWI4MK7dPpzDZVu7Ke13jrclPXuU15zHL2pNe
# 3I6PgNq2kZhAkHnDeMe2scS1ahg4AxCN2NQ3pC4FfYj1gj4QkXCrVYJBMtfbBHMq
# bpEBfCFM1LyuGwN1XXhm2ToxRJozQL8I11pJpMLmqaBn3aQnvKFPObURWBf3JFxG
# j2T3wWmIdph2PVldQnaHiZdpekjw4KISG2aadMreSx7nDmOu5tTvkpI6nj3cAORF
# JYm2mkQZK37AlLTSYW3rM9nF30sEAMx9HJXDj/chsrIRt7t/8tWMcCxBYKqxYxhE
# lRp2Yn72gLD76GSmM9GJB+G9t+ZDpBi4pncB4Q+UDCEdslQpJYls5Q5SUUd0vias
# tkF13nqsX40/ybzTQRESW+UQUOsxxcpyFiIJ33xMdT9j7CFfxCBRa2+xq4aLT8LW
# RV+dIPyhHsXAj6KxfgommfXkaS+YHS312amyHeUbAgMBAAGjggFeMIIBWjAPBgNV
# HRMBAf8EBTADAQH/MB0GA1UdDgQWBBTs1+OC0nFdZEzfLmc/57qYrhwPTzAfBgNV
# HSMEGDAWgBRF66Kv9JLLgjEtUYunpyGd823IDzAOBgNVHQ8BAf8EBAMCAYYwEwYD
# VR0lBAwwCgYIKwYBBQUHAwgweQYIKwYBBQUHAQEEbTBrMCQGCCsGAQUFBzABhhho
# dHRwOi8vb2NzcC5kaWdpY2VydC5jb20wQwYIKwYBBQUHMAKGN2h0dHA6Ly9jYWNl
# cnRzLmRpZ2ljZXJ0LmNvbS9EaWdpQ2VydEFzc3VyZWRJRFJvb3RDQS5jcnQwRQYD
# VR0fBD4wPDA6oDigNoY0aHR0cDovL2NybDMuZGlnaWNlcnQuY29tL0RpZ2lDZXJ0
# QXNzdXJlZElEUm9vdENBLmNybDAgBgNVHSAEGTAXMAgGBmeBDAEEAjALBglghkgB
# hv1sBwEwDQYJKoZIhvcNAQEMBQADggEBAJoWAqUB74H7DbRYsnitqCMZ2XM32mCe
# UdfL+C9AuaMffEBOMz6QPOeJAXWF6GJ7HVbgcbreXsY3vHlcYgBN+El6UU0GMvPF
# 0gAqJyDqiS4VOeAsPvh1fCyCQWE1DyPQ7TWV0oiVKUPL4KZYEHxTjp9FySA3FMDt
# Gbp+dznSVJbHphHfNDP2dVJCSxydjZbVlWxHEhQkXyZB+hpGvd6w5ZFHA6wYCMvL
# 22aJfyucZb++N06+LfOdSsPMzEdeyJWVrdHLuyoGIPk/cuo260VyknopexQDPPtN
# 1khxehARigh0zWwbBFzSipUDdlFQU9Yu90pGw64QLHFMsIe2JzdEYEQwggauMIIE
# lqADAgECAhAHNje3JFR82Ees/ShmKl5bMA0GCSqGSIb3DQEBCwUAMGIxCzAJBgNV
# BAYTAlVTMRUwEwYDVQQKEwxEaWdpQ2VydCBJbmMxGTAXBgNVBAsTEHd3dy5kaWdp
# Y2VydC5jb20xITAfBgNVBAMTGERpZ2lDZXJ0IFRydXN0ZWQgUm9vdCBHNDAeFw0y
# MjAzMjMwMDAwMDBaFw0zNzAzMjIyMzU5NTlaMGMxCzAJBgNVBAYTAlVTMRcwFQYD
# VQQKEw5EaWdpQ2VydCwgSW5jLjE7MDkGA1UEAxMyRGlnaUNlcnQgVHJ1c3RlZCBH
# NCBSU0E0MDk2IFNIQTI1NiBUaW1lU3RhbXBpbmcgQ0EwggIiMA0GCSqGSIb3DQEB
# AQUAA4ICDwAwggIKAoICAQDGhjUGSbPBPXJJUVXHJQPE8pE3qZdRodbSg9GeTKJt
# oLDMg/la9hGhRBVCX6SI82j6ffOciQt/nR+eDzMfUBMLJnOWbfhXqAJ9/UO0hNoR
# 8XOxs+4rgISKIhjf69o9xBd/qxkrPkLcZ47qUT3w1lbU5ygt69OxtXXnHwZljZQp
# 09nsad/ZkIdGAHvbREGJ3HxqV3rwN3mfXazL6IRktFLydkf3YYMZ3V+0VAshaG43
# IbtArF+y3kp9zvU5EmfvDqVjbOSmxR3NNg1c1eYbqMFkdECnwHLFuk4fsbVYTXn+
# 149zk6wsOeKlSNbwsDETqVcplicu9Yemj052FVUmcJgmf6AaRyBD40NjgHt1bicl
# kJg6OBGz9vae5jtb7IHeIhTZgirHkr+g3uM+onP65x9abJTyUpURK1h0QCirc0PO
# 30qhHGs4xSnzyqqWc0Jon7ZGs506o9UD4L/wojzKQtwYSH8UNM/STKvvmz3+Drhk
# Kvp1KCRB7UK/BZxmSVJQ9FHzNklNiyDSLFc1eSuo80VgvCONWPfcYd6T/jnA+bIw
# pUzX6ZhKWD7TA4j+s4/TXkt2ElGTyYwMO1uKIqjBJgj5FBASA31fI7tk42PgpuE+
# 9sJ0sj8eCXbsq11GdeJgo1gJASgADoRU7s7pXcheMBK9Rp6103a50g5rmQzSM7TN
# sQIDAQABo4IBXTCCAVkwEgYDVR0TAQH/BAgwBgEB/wIBADAdBgNVHQ4EFgQUuhbZ
# bU2FL3MpdpovdYxqII+eyG8wHwYDVR0jBBgwFoAU7NfjgtJxXWRM3y5nP+e6mK4c
# D08wDgYDVR0PAQH/BAQDAgGGMBMGA1UdJQQMMAoGCCsGAQUFBwMIMHcGCCsGAQUF
# BwEBBGswaTAkBggrBgEFBQcwAYYYaHR0cDovL29jc3AuZGlnaWNlcnQuY29tMEEG
# CCsGAQUFBzAChjVodHRwOi8vY2FjZXJ0cy5kaWdpY2VydC5jb20vRGlnaUNlcnRU
# cnVzdGVkUm9vdEc0LmNydDBDBgNVHR8EPDA6MDigNqA0hjJodHRwOi8vY3JsMy5k
# aWdpY2VydC5jb20vRGlnaUNlcnRUcnVzdGVkUm9vdEc0LmNybDAgBgNVHSAEGTAX
# MAgGBmeBDAEEAjALBglghkgBhv1sBwEwDQYJKoZIhvcNAQELBQADggIBAH1ZjsCT
# tm+YqUQiAX5m1tghQuGwGC4QTRPPMFPOvxj7x1Bd4ksp+3CKDaopafxpwc8dB+k+
# YMjYC+VcW9dth/qEICU0MWfNthKWb8RQTGIdDAiCqBa9qVbPFXONASIlzpVpP0d3
# +3J0FNf/q0+KLHqrhc1DX+1gtqpPkWaeLJ7giqzl/Yy8ZCaHbJK9nXzQcAp876i8
# dU+6WvepELJd6f8oVInw1YpxdmXazPByoyP6wCeCRK6ZJxurJB4mwbfeKuv2nrF5
# mYGjVoarCkXJ38SNoOeY+/umnXKvxMfBwWpx2cYTgAnEtp/Nh4cku0+jSbl3ZpHx
# cpzpSwJSpzd+k1OsOx0ISQ+UzTl63f8lY5knLD0/a6fxZsNBzU+2QJshIUDQtxMk
# zdwdeDrknq3lNHGS1yZr5Dhzq6YBT70/O3itTK37xJV77QpfMzmHQXh6OOmc4d0j
# /R0o08f56PGYX/sr2H7yRp11LB4nLCbbbxV7HhmLNriT1ObyF5lZynDwN7+YAN8g
# Fk8n+2BnFqFmut1VwDophrCYoCvtlUG3OtUVmDG0YgkPCr2B2RP+v6TR81fZvAT6
# gt4y3wSJ8ADNXcL50CN/AAvkdgIm2fBldkKmKYcJRyvmfxqkhQ/8mJb2VVQrH4D6
# wPIOK+XW+6kvRBVK5xMOHds3OBqhK/bt1nz8MIIGxjCCBK6gAwIBAgIQCnpKiJ7J
# mUKQBmM4TYaXnTANBgkqhkiG9w0BAQsFADBjMQswCQYDVQQGEwJVUzEXMBUGA1UE
# ChMORGlnaUNlcnQsIEluYy4xOzA5BgNVBAMTMkRpZ2lDZXJ0IFRydXN0ZWQgRzQg
# UlNBNDA5NiBTSEEyNTYgVGltZVN0YW1waW5nIENBMB4XDTIyMDMyOTAwMDAwMFoX
# DTMzMDMxNDIzNTk1OVowTDELMAkGA1UEBhMCVVMxFzAVBgNVBAoTDkRpZ2lDZXJ0
# LCBJbmMuMSQwIgYDVQQDExtEaWdpQ2VydCBUaW1lc3RhbXAgMjAyMiAtIDIwggIi
# MA0GCSqGSIb3DQEBAQUAA4ICDwAwggIKAoICAQC5KpYjply8X9ZJ8BWCGPQz7sxc
# bOPgJS7SMeQ8QK77q8TjeF1+XDbq9SWNQ6OB6zhj+TyIad480jBRDTEHukZu6aNL
# SOiJQX8Nstb5hPGYPgu/CoQScWyhYiYB087DbP2sO37cKhypvTDGFtjavOuy8YPR
# n80JxblBakVCI0Fa+GDTZSw+fl69lqfw/LH09CjPQnkfO8eTB2ho5UQ0Ul8PUN7U
# WSxEdMAyRxlb4pguj9DKP//GZ888k5VOhOl2GJiZERTFKwygM9tNJIXogpThLwPu
# f4UCyYbh1RgUtwRF8+A4vaK9enGY7BXn/S7s0psAiqwdjTuAaP7QWZgmzuDtrn8o
# LsKe4AtLyAjRMruD+iM82f/SjLv3QyPf58NaBWJ+cCzlK7I9Y+rIroEga0OJyH5f
# sBrdGb2fdEEKr7mOCdN0oS+wVHbBkE+U7IZh/9sRL5IDMM4wt4sPXUSzQx0jUM2R
# 1y+d+/zNscGnxA7E70A+GToC1DGpaaBJ+XXhm+ho5GoMj+vksSF7hmdYfn8f6Cvk
# FLIW1oGhytowkGvub3XAsDYmsgg7/72+f2wTGN/GbaR5Sa2Lf2GHBWj31HDjQpXo
# nrubS7LitkE956+nGijJrWGwoEEYGU7tR5thle0+C2Fa6j56mJJRzT/JROeAiylC
# cvd5st2E6ifu/n16awIDAQABo4IBizCCAYcwDgYDVR0PAQH/BAQDAgeAMAwGA1Ud
# EwEB/wQCMAAwFgYDVR0lAQH/BAwwCgYIKwYBBQUHAwgwIAYDVR0gBBkwFzAIBgZn
# gQwBBAIwCwYJYIZIAYb9bAcBMB8GA1UdIwQYMBaAFLoW2W1NhS9zKXaaL3WMaiCP
# nshvMB0GA1UdDgQWBBSNZLeJIf5WWESEYafqbxw2j92vDTBaBgNVHR8EUzBRME+g
# TaBLhklodHRwOi8vY3JsMy5kaWdpY2VydC5jb20vRGlnaUNlcnRUcnVzdGVkRzRS
# U0E0MDk2U0hBMjU2VGltZVN0YW1waW5nQ0EuY3JsMIGQBggrBgEFBQcBAQSBgzCB
# gDAkBggrBgEFBQcwAYYYaHR0cDovL29jc3AuZGlnaWNlcnQuY29tMFgGCCsGAQUF
# BzAChkxodHRwOi8vY2FjZXJ0cy5kaWdpY2VydC5jb20vRGlnaUNlcnRUcnVzdGVk
# RzRSU0E0MDk2U0hBMjU2VGltZVN0YW1waW5nQ0EuY3J0MA0GCSqGSIb3DQEBCwUA
# A4ICAQANLSN0ptH1+OpLmT8B5PYM5K8WndmzjJeCKZxDbwEtqzi1cBG/hBmLP13l
# hk++kzreKjlaOU7YhFmlvBuYquhs79FIaRk4W8+JOR1wcNlO3yMibNXf9lnLocLq
# THbKodyhK5a4m1WpGmt90fUCCU+C1qVziMSYgN/uSZW3s8zFp+4O4e8eOIqf7xHJ
# MUpYtt84fMv6XPfkU79uCnx+196Y1SlliQ+inMBl9AEiZcfqXnSmWzWSUHz0F6aH
# ZE8+RokWYyBry/J70DXjSnBIqbbnHWC9BCIVJXAGcqlEO2lHEdPu6cegPk8QuTA2
# 5POqaQmoi35komWUEftuMvH1uzitzcCTEdUyeEpLNypM81zctoXAu3AwVXjWmP5U
# bX9xqUgaeN1Gdy4besAzivhKKIwSqHPPLfnTI/KeGeANlCig69saUaCVgo4oa6TO
# nXbeqXOqSGpZQ65f6vgPBkKd3wZolv4qoHRbY2beayy4eKpNcG3wLPEHFX41tOa1
# DKKZpdcVazUOhdbgLMzgDCS4fFILHpl878jIxYxYaa+rPeHPzH0VrhS/inHfypex
# 2EfqHIXgRU4SHBQpWMxv03/LvsEOSm8gnK7ZczJZCOctkqEaEf4ymKZdK5fgi9Oc
# zG21Da5HYzhHF1tvE9pqEG4fSbdEW7QICodaWQR2EaGndwITHDGCBSowggUmAgEB
# MFQwTjELMAkGA1UEBhMCTFUxJjAkBgNVBAoMHU1ham9yZWwgR3JvdXAgTHV4ZW1i
# b3VyZyBTLkEuMRcwFQYDVQQDDA5NYWpvcmVsIENBIElNMQICEAAwDQYJYIZIAWUD
# BAIBBQCggYQwGAYKKwYBBAGCNwIBDDEKMAigAoAAoQKAADAZBgkqhkiG9w0BCQMx
# DAYKKwYBBAGCNwIBBDAcBgorBgEEAYI3AgELMQ4wDAYKKwYBBAGCNwIBFjAvBgkq
# hkiG9w0BCQQxIgQgNebMsltLwxwoGFC0vR9AC5g86+yc2vl6F6yhLnven/4wDQYJ
# KoZIhvcNAQEBBQAEggEAE8pqGQS5pQahpGdMdXE0e2xCwO+nkVIe/Wk/vpaoQPd6
# f5p+P46fYMt4i6a8/nP8BslT1YQOz7k53ha+WJVSIcRekbf94Z4VpFAj7Rrk/a1a
# Mm97LzLiL866OhXkmBvPw1DIzP0Gjk8wHDq+4eFh49AVHCLNPwY4mzg7+oYyq4jr
# DSyrwcF1hgzwLWH5Z24RZLdg7ujYAHHM4PlsSxoDXJNm65kgEnbpheiTWv7PVXZ7
# kEd1PgD4srq8In2ET1EPOlMQ4Q5ylFWnFC8gYZhPbdJzQ+N9QXGsQTzPFmnUdI0m
# Za1JElwO+Ip7PqezjrC8da3odnAU+KIskvf/ONLC+qGCAyAwggMcBgkqhkiG9w0B
# CQYxggMNMIIDCQIBATB3MGMxCzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2Vy
# dCwgSW5jLjE7MDkGA1UEAxMyRGlnaUNlcnQgVHJ1c3RlZCBHNCBSU0E0MDk2IFNI
# QTI1NiBUaW1lU3RhbXBpbmcgQ0ECEAp6SoieyZlCkAZjOE2Gl50wDQYJYIZIAWUD
# BAIBBQCgaTAYBgkqhkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEP
# Fw0yMjA3MDYwOTIwMTRaMC8GCSqGSIb3DQEJBDEiBCDSkqz8WKmLeBjxa4r02bT5
# KAF3hGIAPf6Toy1vbzFVEDANBgkqhkiG9w0BAQEFAASCAgBgfAnAKNDDV05UF4LS
# OMxuxvEBTjyJdA1UobfrVeWSXcgUWdzTqt3mkRjjr6K+JRzt5pZ6vAbw8YnMuIPz
# 0aZxlxOqfyG2IlhezbJo5BhAgps+tQRwzqwum1RsoG7rsGEAxcWijChFb9xLFDbB
# kl83AegnEYi4C/p7JTgXg914GWbHg3mNSOO/k6eah1nmqtcU3g2Yp4vg8+AUn2Fq
# skjCFqKI3Xl5I2rUOPa76iXEDMfu1g17ncJTjkcAG9K14Fseo/+C9LWRdmEzbqOH
# 4XbN4acw2EEvN9NIwTglPK/ISkmKv0EMrhayI/BAKLQjVwx0+dyjPAhgBr6FIkNN
# ueaN4Z40j09ZsQYc3gaoHTfME9Cjzo0xKfIArXJtRagj6IyPYaW+5z5LYNQIG1yU
# HfdHV7AyyTNedsjkaX4cpKoD1syuK4IgrqHpgHsUmALD2JxwDvPuWt6NDzVqa2GJ
# vHd4TjrmCmjVyxLALzTAluKfombdWJdV8JBcVDIDqrWTPyf8o6nZio+73NMHC7AY
# i7loJ0ehodcvpU6I+cD629j0M78m6lDJdwGC+sn/lv4VowxZ7ply5V4n9DKXBtmn
# V027gmPnbF9zLCz6hT/CwgTiXA4WSfIzYYAuPBeGIRgbCUSapaM5yrN9L4M5dBTa
# bk0tWbkS6ahdCJRdqGYg7H3pLw==
# SIG # End signature block
