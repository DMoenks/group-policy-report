# Author: Mönks, Dominik
# Version: 4.0

# Save NewLine character for easier use
$nl = [Environment]::NewLine
# Save script path for easier use
$path = Split-Path -parent $MyInvocation.MyCommand.Definition

#region HTML preparations
$XMLWriterSettings = New-Object System.Xml.XmlWriterSettings
$XMLWriterSettings.ConformanceLevel = [System.Xml.ConformanceLevel]::Fragment
$XMLWriterSettings.Indent = $true
$XMLWriter = $null

function openFile([string]$Domain)
{
    [System.Xml.XmlWriter]::Create("$path\$Domain\GPOReport.html", $XMLWriterSettings)
}

function closeFile()
{
    $XMLWriter.Close()
    $XMLWriter = $null
}

function openTag([string]$Name)
{
    $XMLWriter.WriteStartElement($Name)
}

function closeTag()
{
    $XMLWriter.WriteEndElement()
}

function attribute([string]$Name, [string]$Value)
{
    $XMLWriter.WriteAttributeString($Name, $Value)
}

function content([string]$Value)
{
    $XMLWriter.WriteString($Value)
}
#endregion

foreach ($Domain in (Get-ADForest).Domains)
{
    $policies = @{}
    # Request the list of GPOs from the target domain
    foreach ($GPO in (Get-GPO -All -Domain $Domain))
    {
        $policies.Add($GPO.DisplayName, @{})
        # Check if the target folder exists and, if not create it
        if (!(Test-Path ("$path\$Domain\" + $($GPO.DisplayName.Replace(":", "")))))
        {
            New-Item -Path ("$path\$Domain\" + $GPO.DisplayName.Replace(":", "")) -ItemType directory
        }
        # Check if the GPO report for the most recent version already exists in the target folder and, if not, create it
        if (!(Test-Path ("$path\$Domain\" + $GPO.DisplayName.Replace(":", "") + "\Report" + $GPO.ModificationTime.ToString("yyyyMMdd") + ".xml")))
        {
            Get-GPOReport -Guid $GPO.Id -ReportType Xml -Domain $Domain -Path ("$path\$Domain\" + $GPO.DisplayName.Replace(":", "") + "\Report" + $GPO.ModificationTime.ToString("yyyyMMdd") + ".xml")
        }
        # Check if there is more than one report for the current GPO and, if so, create an entry for the domain's diff file
        if ((Get-ChildItem -Path ("$path\$Domain\" + $GPO.DisplayName.Replace(":", "")) -Filter "*.xml").Count -gt 1)
        {
            for ($i = 0;$i -lt (Get-ChildItem -Path ("$path\$Domain\" + $GPO.DisplayName.Replace(":", "")) -Filter "*.xml").Count - 1;$i++)
            {
                $CurVerFil = (Get-ChildItem -Path ("$path\$Domain\" + $GPO.DisplayName.Replace(":", "")) -Filter "*.xml" | Sort-Object -Property Name -Descending)[$i].FullName
                $PreVerFil = (Get-ChildItem -Path ("$path\$Domain\" + $GPO.DisplayName.Replace(":", "")) -Filter "*.xml" | Sort-Object -Property Name -Descending)[$i+1].FullName
                $CurVer = Get-Content $CurVerFil
                $PreVer = Get-Content $PreVerFil
                $CurVerDif = (Compare-Object $CurVer $PreVer | Where-Object -Property SideIndicator -EQ "<=").InputObject
                $PreVerDif = (Compare-Object $CurVer $PreVer | Where-Object -Property SideIndicator -EQ "=>").InputObject
                $policies[$GPO.DisplayName].Add([Regex]::Match($CurVerFil, "\d{8}").Value + "|" + [Regex]::Match($PreVerFil, "\d{8}").Value, @{})
                $policies[$GPO.DisplayName][[Regex]::Match($CurVerFil, "\d{8}").Value + "|" + [Regex]::Match($PreVerFil, "\d{8}").Value].Add([Regex]::Match($CurVerFil, "\d{8}").Value, @())
                $diffcounter = 0
                foreach ($line in $CurVer)
                {
                    if ($line -eq $CurVerDif[$diffcounter])
                    {
                        $policies[$GPO.DisplayName][[Regex]::Match($CurVerFil, "\d{8}").Value + "|" + [Regex]::Match($PreVerFil, "\d{8}").Value][[Regex]::Match($CurVerFil, "\d{8}").Value] += ("1:" + $line)
                        $diffcounter++
                    }
                    else
                    {
                        $policies[$GPO.DisplayName][[Regex]::Match($CurVerFil, "\d{8}").Value + "|" + [Regex]::Match($PreVerFil, "\d{8}").Value][[Regex]::Match($CurVerFil, "\d{8}").Value] += ("0:" + $line)
                    }
                }
                $policies[$GPO.DisplayName][[Regex]::Match($CurVerFil, "\d{8}").Value + "|" + [Regex]::Match($PreVerFil, "\d{8}").Value].Add([Regex]::Match($PreVerFil, "\d{8}").Value, @())
                $diffcounter = 0
                foreach ($line in $PreVer)
                {
                    if ($line -eq $PreVerDif[$diffcounter])
                    {
                        $policies[$GPO.DisplayName][[Regex]::Match($CurVerFil, "\d{8}").Value + "|" + [Regex]::Match($PreVerFil, "\d{8}").Value][[Regex]::Match($PreVerFil, "\d{8}").Value] += ("1:" + $line)
                        $diffcounter++
                    }
                    else
                    {
                        $policies[$GPO.DisplayName][[Regex]::Match($CurVerFil, "\d{8}").Value + "|" + [Regex]::Match($PreVerFil, "\d{8}").Value][[Regex]::Match($PreVerFil, "\d{8}").Value] += ("0:" + $line)
                    }
                }
            }
        }
        else
        {
            $CurVerFil = (Get-ChildItem -Path ("$path\$Domain\" + $GPO.DisplayName.Replace(":", "")) -Filter "*.xml" | Sort-Object -Property Name -Descending)[0].FullName
            $CurVer = Get-Content $CurVerFil
            $policies[$GPO.DisplayName].Add([Regex]::Match($CurVerFil, "\d{8}").Value, @{})
            $policies[$GPO.DisplayName][[Regex]::Match($CurVerFil, "\d{8}").Value].Add([Regex]::Match($CurVerFil, "\d{8}").Value, @())
            foreach ($line in $CurVer)
            {
                $policies[$GPO.DisplayName][[Regex]::Match($CurVerFil, "\d{8}").Value][[Regex]::Match($CurVerFil, "\d{8}").Value] += ("0:" + $line)
            }
        }
    }
    $XMLWriter = openFile $Domain
        openTag "head"
            openTag "link"
                attribute "rel" "stylesheet"
                attribute "type" "text/css"
                attribute "href" "..\GPOReports.css"
            closeTag
            openTag "title"
                content "GPOReports"
            closeTag
        closeTag
        openTag "body"
            foreach ($policy in ($policies.Keys | Sort-Object))
            {
                openTag "div"
                    attribute "class" "policy"
                    openTag "p"
                        content $policy
                        openTag "span"
                            #Hier fehlt noch das Datum der aktuellen Version
                        closeTag
                    closeTag
                    openTag "div"
                        openTag "div"
                            openTag "select"
                                attribute id $policy
                                attribute size 1
                                foreach ($diff in ($policies[$policy].Keys | Sort-Object -Descending))
                                {
                                    openTag "option"
                                        attribute "value" "$policy.$diff"
                                        content $diff.Replace("|", " vs. ")
                                    closeTag
                                }
                            closeTag
                            openTag "button"
                                attribute "onclick" "openSelected('$policy')"
                                content "Open"
                            closeTag
                            openTag "button"
                                attribute "onclick" "closeAll()"
                                content "Close"
                            closeTag
                        closeTag
                        openTag "div"
                            foreach ($diff in ($policies[$policy].Keys | Sort-Object -Descending))
                            {
                                openTag "div"
                                    attribute "class" "comparison"
                                    attribute "id" "$policy.$diff"
                                    openTag "table"
                                        if ($diff.Contains("|"))
                                        {
                                            openTag "thead"
                                                openTag "tr"
                                                    openTag "th"
                                                        content $diff.Split("|")[0]
                                                    closeTag
                                                    openTag "th"
                                                        content $diff.Split("|")[1]
                                                    closeTag
                                                closeTag
                                            closeTag
                                            openTag "tbody"
                                                openTag "tr"
                                                    openTag "td"
                                                        openTag "div"
                                                            attribute "id" "$policy.$diff.left"
                                                            attribute "onmousedown" "unbindScroll('$policy.$diff.right')"
                                                            attribute "onmouseup" "bindScroll('$policy.$diff.right')"
                                                            attribute "onscroll" "scroll('$policy.$diff.left', '$policy.$diff.right')"
                                                            foreach ($line in $policies[$policy][$diff][$diff.Split("|")[0]])
                                                            {
                                                                if ($line.StartsWith("1:"))
                                                                {
                                                                    openTag "span"
                                                                        content $line.SubString(2)
                                                                    closeTag
                                                                }
                                                                else
                                                                {
                                                                    content $line.SubString(2)
                                                                }
                                                                openTag "br"
                                                                closeTag
                                                            }
                                                        closeTag
                                                    closeTag
                                                    openTag "td"
                                                        openTag "div"
                                                            attribute "id" "$policy.$diff.right"
                                                            attribute "onmousedown" "unbindScroll('$policy.$diff.left')"
                                                            attribute "onmouseup" "bindScroll('$policy.$diff.left')"
                                                            attribute "onscroll" "scroll('$policy.$diff.right', '$policy.$diff.left')"
                                                            foreach ($line in $policies[$policy][$diff][$diff.Split("|")[1]])
                                                            {
                                                                if ($line.StartsWith("1:"))
                                                                {
                                                                    openTag "span"
                                                                        content $line.SubString(2)
                                                                    closeTag
                                                                }
                                                                else
                                                                {
                                                                    content $line.SubString(2)
                                                                }
                                                                openTag "br"
                                                                closeTag
                                                            }
                                                        closeTag
                                                    closeTag
                                                closeTag
                                            closeTag
                                        }
                                        else
                                        {
                                            openTag "thead"
                                                openTag "tr"
                                                    openTag "th"
                                                        content $diff
                                                    closeTag
                                                closeTag
                                            closeTag
                                            openTag "tbody"
                                                openTag "tr"
                                                    openTag "td"
                                                        openTag "div"
                                                            foreach ($line in $policies[$policy][$diff][$diff])
                                                            {
                                                                content $line.SubString(2)
                                                                openTag "br"
                                                                closeTag
                                                            }
                                                        closeTag
                                                    closeTag
                                                closeTag
                                            closeTag
                                        }
                                    closeTag
                                closeTag
                            }
                        closeTag
                    closeTag
                closeTag
            }
            openTag "script"
                attribute "type" "text/javascript"
                attribute "src" "..\GPOReports.js"
                content ""
            closeTag
        closeTag
    closeFile
}
