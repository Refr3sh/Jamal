If ($env:USERNAME -eq "C3585149")
{
	
#ClearUpTemps
	$ClsTmps = "$env:TEMP", "$env:UserProfile\Appdata\Roaming\Microsoft\Windows\Recent\*" , "$env:UserProfile\Appdata\Roaming\Microsoft\Windows\Cookies\*"
	rm $ClsTmps -force -Recurse -ErrorAction SilentlyContinue

#Excel
    $sourceUpd = "$env:APPDATA\Microsoft\Excel\XLSTART\"
    $source = "\\gbhgmercser0021\pdg\DESPATCH MISC\SYW Order Compliance\Empty Jamal\Script\XLSTART\PERSONAL.XLSB"
    mkdir $sourceUpd -Force -ErrorAction SilentlyContinue
    Copy-Item -Path $source -Destination $sourceUpd -Recurse -Force -ErrorAction SilentlyContinue
	Remove-Variable source
	Remove-Variable sourceUpd

#Start
	$JamalStart = "$env:APPDATA\Microsoft\Windows\Start Menu\Programs\Startup\"
	$JamalScriptS = "\\gbhgmercser0021\pdg\DESPATCH MISC\SYW Order Compliance\Empty Jamal\UpdateJamalScript.lnk"
	$JamalClockS = "\\gbhgmercser0021\pdg\DESPATCH MISC\SYW Order Compliance\Empty Jamal\Script\clock.lnk"
	#Clear	
	rm -path $JamalStart\* -Force -Recurse -ErrorAction SilentlyContinue
	#Do
	Copy-Item -Path $JamalScriptS -Destination $JamalStart -Recurse -Force -ErrorAction SilentlyContinue
	Copy-Item -Path $JamalClockS -Destination $JamalStart -Recurse -Force -ErrorAction SilentlyContinue

#Desktop
	$JamalComp = "$env:UserProfile\Desktop\"
	$JamalMain = "\\gbhgmercser0021\pdg\DESPATCH MISC\SYW Order Compliance\Empty Jamal\Script\Desktop\*"
	Copy-Item -Path $JamalMain -Destination $JamalComp -Recurse -Force -ErrorAction SilentlyContinue



#Fonts
	Add-Type -Name Session -Namespace "" -Member @"
	[DllImport("gdi32.dll")]
	public static extern int AddFontResource(string filePath);
"@

	$files = Get-ChildItem -path "\\gbhgmercser0021\pdg\mine\fonts\*" -Recurse -Include *.ttf, *.otf

	foreach($font in $files)
	{
    echo $font
    [Session]::AddFontResource($font.FullName)
	}

#Quiet	
	Start-Sleep 10
	$makequiet = "*authman*","*concen*","*wfcrun*","*redirect*","*selfservicePlug*","*receiver*","*chrome*","*lync*","*teams*"
	foreach( $process in $makequiet )
	{
    kill -name $makequiet -Force
	}
	
}else{
exit
}
