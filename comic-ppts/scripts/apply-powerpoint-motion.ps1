param(
 [Parameter(Mandatory=$true)][string]$InputPath,
 [Parameter(Mandatory=$true)][string]$OutputPath,
 [Parameter(Mandatory=$true)][string]$PlanPath,
 [switch]$CheckPlayback
)
$ErrorActionPreference='Stop'
$inputFile=(Resolve-Path -LiteralPath $InputPath).Path
$planFile=(Resolve-Path -LiteralPath $PlanPath).Path
$outputFile=[IO.Path]::GetFullPath($OutputPath)
if($inputFile -eq $outputFile){throw 'Input and output must be different files.'}
if(Test-Path -LiteralPath $outputFile){throw 'Output already exists. Choose a new filename.'}
$plan=Get-Content -LiteralPath $planFile -Raw -Encoding UTF8 | ConvertFrom-Json
$duration=0.35
if($null -ne $plan.duration){$duration=[double]$plan.duration}
if($duration -lt 0.1 -or $duration -gt 2){throw 'Duration must be 0.1 to 2 seconds.'}
New-Item -ItemType Directory -Path ([IO.Path]::GetDirectoryName($outputFile)) -Force | Out-Null
$app=New-Object -ComObject PowerPoint.Application
$deck=$null;$view=$null
try {
 $deck=$app.Presentations.Open($inputFile,0,0,0)
 foreach($slide in $deck.Slides){
  $slide.SlideShowTransition.EntryEffect=1793
  $slide.SlideShowTransition.Duration=$duration
  $slide.SlideShowTransition.AdvanceOnClick=-1
  $slide.SlideShowTransition.AdvanceOnTime=0
 }
 foreach($entry in $plan.slides){
  $slide=$deck.Slides.Item([int]$entry.slide)
  # The explicit plan replaces only this slide's main sequence.
  for($i=$slide.TimeLine.MainSequence.Count;$i -ge 1;$i--){$slide.TimeLine.MainSequence.Item($i).Delete()}
  foreach($target in $entry.targets){
   $shape=$null
   foreach($item in $slide.Shapes){if($item.Name -eq $target.name){$shape=$item;break}}
   if($null -eq $shape){throw "Missing target '$($target.name)' on slide $($entry.slide)."}
   $id=10
   if($target.effect -eq 'Appear'){$id=1}
   elseif($target.effect -and $target.effect -ne 'Fade'){throw 'Supported effects are Fade and Appear.'}
   $effect=$slide.TimeLine.MainSequence.AddEffect($shape,$id,0,1)
   $effect.Timing.Duration=$duration
  }
 }
 $deck.SaveAs($outputFile,24)
 $deck.Close();$deck=$null
 $windowFlag=0
 if($CheckPlayback){$app.Visible=-1;$windowFlag=-1}
 $deck=$app.Presentations.Open($outputFile,0,0,$windowFlag)
 $results=@()
 foreach($entry in $plan.slides){
  $slide=$deck.Slides.Item([int]$entry.slide)
  $actual=@();foreach($effect in $slide.TimeLine.MainSequence){$actual+=$effect.Shape.Name}
  $expected=@($entry.targets | ForEach-Object {$_.name})
  if(($actual -join '|') -ne ($expected -join '|')){throw "Saved sequence mismatch on slide $($entry.slide)."}
  $results+=[pscustomobject]@{slide=[int]$entry.slide;targets=$actual;effects=$actual.Count;manualAdvance=($slide.SlideShowTransition.AdvanceOnTime -eq 0)}
 }
 $steps=@()
 if($CheckPlayback){
  $deck.SlideShowSettings.ShowType=2
  $deck.Windows.Item(1).Activate()
  Start-Sleep -Milliseconds 500
  $window=$deck.SlideShowSettings.Run()
  for($attempt=0;$attempt -lt 10 -and $null -eq $window;$attempt++){
   Start-Sleep -Milliseconds 200
   if($app.SlideShowWindows.Count -gt 0){$window=$app.SlideShowWindows.Item($app.SlideShowWindows.Count)}
  }
  $view=$window.View
  if($null -eq $view){throw 'PowerPoint did not start a slideshow window.'}
  $wait=[int]($duration*1000+150)
  foreach($entry in $plan.slides){
   $num=[int]$entry.slide;$view.GotoSlide($num,-1)
   Start-Sleep -Milliseconds $wait
   if($view.GetClickIndex() -ne 0){throw "Slide $num did not start before its first click."}
   for($i=1;$i -le @($entry.targets).Count;$i++){
    $view.Next();Start-Sleep -Milliseconds $wait
    $index=$view.GetClickIndex();$current=$view.Slide.SlideIndex
    if($current -ne $num -or $index -ne $i){throw "Unexpected playback state on slide $num click $i."}
    $steps+=[pscustomobject]@{slide=$current;click=$index}
   }
  }
  $view.Exit();$view=$null
 }
 [pscustomobject]@{output=$outputFile;structure=$results;playbackChecked=[bool]$CheckPlayback;clicks=$steps} | ConvertTo-Json -Depth 8
}finally{
 if($null -ne $view){try{$view.Exit()}catch{}}
 if($null -ne $deck){$deck.Close()}
 [void][Runtime.InteropServices.Marshal]::ReleaseComObject($app)
}
