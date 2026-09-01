param([Parameter(Mandatory=$true)][string]$SampleDirectory,[Parameter(Mandatory=$true)][string]$OutputDirectory,[string]$Password='',[string]$PythonPath='')
$ErrorActionPreference='Stop'
function T($v){ if($null -eq $v){''}else{([string]$v).Trim()} }
function N($v){
  if($null -eq $v -or [string]::IsNullOrWhiteSpace([string]$v)){ return [double]0 }
  try { return [double]([string]$v) } catch { return [double]0 }
}
function Release-Com($o){if($null -ne $o){[void][Runtime.InteropServices.Marshal]::ReleaseComObject($o)}}
function Change($s,$c,$b,$a,$src,$rule){if((T $b) -ne (T $a)){$script:changes += [ordered]@{sheet=$s;cell=$c;before=$b;after=$a;source=$src;rule=$rule}}}
$master=Get-ChildItem -LiteralPath $SampleDirectory -Filter '202607*工资核算总表-v2.xlsx'|select -First 1
$salary=Get-ChildItem -LiteralPath $SampleDirectory -Filter '薪资数据-科园-7月薪资*.xlsx'|select -First 1
$social=@(Get-ChildItem -LiteralPath $SampleDirectory -Filter '*2026.07.xlsx'|where {$_.Name -match '益药科园|鹤安长泰'})
$taxDir=Get-ChildItem -LiteralPath $SampleDirectory -Directory|where {$_.Name -like '回复：202608*'}|select -First 1
if(!$master -or !$salary -or $social.Count -ne 2 -or !$taxDir){throw '演示素材不完整'}
New-Item -ItemType Directory -Path $OutputDirectory -Force|Out-Null
$formal=Join-Path $OutputDirectory '202608（所属月202607）-北京科园-鹤安-大药房工资核算总表-正式稿.xlsx'
$review=Join-Path $OutputDirectory '202608（所属月202607）-北京科园-鹤安-大药房工资核算总表-修改稿.xlsx'
$issuesPath=Join-Path $OutputDirectory '待人工确认-科园202607.xlsx';$reportPath=Join-Path $OutputDirectory '验收报告-科园202607.json';$working=Join-Path $OutputDirectory '.working.xlsx'
$python=if($PythonPath){$PythonPath}else{Join-Path (Split-Path $PSScriptRoot -Parent) '.venv-codex\Scripts\python.exe'}
$decrypt=Join-Path $PSScriptRoot 'decrypt_workbook_copy.py'
$signature=[byte[]](Get-Content -LiteralPath $master.FullName -AsByteStream -TotalCount 2)
if($signature.Count -ge 2 -and $signature[0] -eq 0x50 -and $signature[1] -eq 0x4B){
 Copy-Item -LiteralPath $master.FullName -Destination $working -Force;$decryptExit=0
}else{
 & $python $decrypt $master.FullName $working $Password;$decryptExit=$LASTEXITCODE
}
if($decryptExit -ne 0){throw '总表副本解密失败'}
$changes=@();$issues=@();$app=$null;$book=$null;$srcBook=$null
try{
 try{$app=New-Object -ComObject Excel.Application}catch{$app=New-Object -ComObject KET.Application}
 $app.Visible=$false;$app.DisplayAlerts=$false;try{$app.AskToUpdateLinks=$false}catch{};try{$app.EnableEvents=$false}catch{};try{$app.Calculation=-4135}catch{}
  $basicProcessor=Join-Path $PSScriptRoot 'keyuan_basic_processor.py'
  $basicOutput=Join-Path $OutputDirectory '.keyuan-basics.xlsx'
  $basicJson=& $python -X utf8 $basicProcessor $working $salary.FullName $basicOutput
  if($LASTEXITCODE -ne 0){throw '基础薪资、考勤、值班或补发补扣处理失败'}
  $basicResult=$basicJson|ConvertFrom-Json
  foreach($x in @($basicResult.changes)){$changes += [ordered]@{sheet=$x.sheet;cell=$x.cell;before=$x.before;after=$x.after;source=$x.source;rule=$x.rule}}
  if(Test-Path $basicOutput){Copy-Item -LiteralPath $basicOutput -Destination $working -Force}
  $book=$app.Workbooks.Open($working,0,$false);$srcBook=$app.Workbooks.Open($salary.FullName,0,$true)
 $payslip=$book.Worksheets.Item('工资条');$payslip.Range('A1:BV41').Copy()|Out-Null;$payslip.Range('A46:BV86').PasteSpecial(-4163)|Out-Null;$payslip.Range('A1:BV41').Copy()|Out-Null;$payslip.Range('A46:BV86').PasteSpecial(-4122)|Out-Null;$payslip.Cells.Item(45,1).Value2='累计工资条';Change '工资条' 'A45:BV86' '' '上月工资条快照' '原始总表' '更新前沉淀上月工资条'
 $oa=$book.Worksheets.Item('OA请款及审批');$sourceHistory=$oa.Range('A8:I13');$targetHistory=$oa.Range('A9:I14');$sourceHistory.Copy()|Out-Null;$targetHistory.PasteSpecial(-4163)|Out-Null;$b=$oa.Cells.Item(8,1).Value2;$oa.Cells.Item(8,1).Value2='7月';Change 'OA请款及审批' 'A8:I14' $b '7月并顺延历史' '操作手册' '更新前冻结上月数据并顺延'
 $p=$book.Worksheets.Item('工资核算');$bonus=$srcBook.Worksheets.Item('奖金-7月');$byName=@{}
 $bonusRows=[int]$bonus.UsedRange.Rows.Count
 for($r=4;$r -le $bonusRows;$r++){
   $n=T ($bonus.Cells.Item($r,4).Value2)
   if($n){
     [double]$j=N ($bonus.Cells.Item($r,10).Value2)
     [double]$k=N ($bonus.Cells.Item($r,11).Value2)
     [double]$l=N ($bonus.Cells.Item($r,12).Value2)
     [double]$sum=$j+$k
     $byName[$n]=[double[]]@($sum,$l)
   }
 }
 for($r=4;$r -le 43;$r++){ $n=T $p.Cells.Item($r,2).Value2;if($byName.ContainsKey($n)){foreach($z in @(@(30,$byName[$n][0]),@(31,$byName[$n][1]))){$c=$p.Cells.Item($r,$z[0]);$b=$c.Value2;$c.Value2=[double]$z[1];Change '工资核算' ($c.Address($false,$false)) $b $z[1] $salary.Name '奖金J+K/L';Release-Com $c}}}
 $att=$srcBook.Worksheets.Item('考勤-7月');$at=$book.Worksheets.Item('考勤');$ab=@{}
 for($r=9;$r -le $att.UsedRange.Rows.Count;$r++){ $n=T $att.Cells.Item($r,1).Value2;if($n){$g=[double](N ($att.Cells.Item($r,7).Value2));$q=[double](N ($att.Cells.Item($r,17).Value2));$ad=[double](N ($att.Cells.Item($r,30).Value2));$af=[double](N ($att.Cells.Item($r,32).Value2));$ag=[double](N ($att.Cells.Item($r,33).Value2));$ah=[double](N ($att.Cells.Item($r,34).Value2));[double]$overtime=$ad+$af;$ab[$n]=[double[]]@($g,$q,$overtime,$ag,$ah)} }
 for($r=2;$r -le 48;$r++){ $n=T $at.Cells.Item($r,2).Value2;if($ab.ContainsKey($n)){ $v=$ab[$n];foreach($z in @(@(7,$v[0]),@(13,$v[1]),@(9,$v[2]),@(10,$v[3]),@(16,$v[4]))){$c=$at.Cells.Item($r,$z[0]);$b=$c.Value2;$c.Value2=[double]$z[1];Change '考勤' ($c.Address($false,$false)) $b $z[1] $salary.Name '考勤G/Q/AD+AF/AG/AH';Release-Com $c}}}
 $duty=$srcBook.Worksheets.Item('值班');$dt=$book.Worksheets.Item('配送员值班费');$counts=@{};for($r=3;$r -le $duty.UsedRange.Rows.Count;$r++){ $n=T $duty.Cells.Item($r,3).Value2;if($n){$counts[$n]=1+([int]($counts[$n]))} };for($r=2;$r -le 6;$r++){ $n=T $dt.Cells.Item($r,2).Value2;$v=if($counts.ContainsKey($n)){$counts[$n]}else{0};$c=$dt.Cells.Item($r,4);$b=$c.Value2;$c.Value2=[int]$v;Change '配送员值班费' ($c.Address($false,$false)) $b $v $salary.Name '值班次数';Release-Com $c}
 $adj=$book.Worksheets.Item('其他调差累计');[void]$adj.Range('A2:C39').ClearContents();$add=$srcBook.Worksheets.Item('补发补扣');$rout=2;for($r=1;$r -le $add.UsedRange.Rows.Count;$r++){ $n=T $add.Cells.Item($r,1).Value2;$note=T $add.Cells.Item($r,2).Value2;if($n -and $note){$m=[regex]::Matches($note,'(?<!\d)(\d+(?:\.\d{1,2})?)(?=元)');if($m.Count){$v=[double]$m[$m.Count-1].Groups[1].Value;$adj.Cells.Item($rout,1).Value2=[string]$n;$adj.Cells.Item($rout,2).Value2=[double]$v;$adj.Cells.Item($rout,3).Value2=[string]$note;Change '其他调差累计' "A$rout:C$rout" '' "$n / $v" $salary.Name '补发补扣';$rout++}else{$issues+= [ordered]@{item='补发补扣';detail="$n 金额无法识别"}}}}
 $heat=$book.Worksheets.Item('防暑降温费');$b=$heat.Range('K1').Value2;$heat.Range('K1').Value2=23;Change '防暑降温费' 'K1' $b 23 '操作手册' '当月应出勤23天';
 for($r=4;$r -le 43;$r++){ $n=T $p.Cells.Item($r,2).Value2;$c=$p.Cells.Item($r,47);$f=T $c.Formula;if($n -ne '李楠' -and $f -match '/21'){$a=$f.Replace('/21','/23');$c.Formula=$a;Change '工资核算' ($c.Address($false,$false)) $f $a '操作手册' 'AU21改23'};Release-Com $c;if($n -in @('李晓苛','袁艳')){$c=$p.Cells.Item($r,48);$f=T $c.Formula;if($f -match '/21'){$a=$f.Replace('/21','/23');$c.Formula=$a;Change '工资核算' ($c.Address($false,$false)) $f $a '操作手册' 'AV21改23'};Release-Com $c}}
 $p.Columns.Item(1).Insert()|Out-Null;$p.Cells.Item(2,1).Value2='班制';for($r=4;$r -le 43;$r++){$p.Cells.Item($r,1).Formula="=VLOOKUP(C$r,考勤!B:Q,16,0)"};Change '工资核算' 'A:A' '' '班制' '操作手册' '新增班制列'
 $book.SaveAs($formal,51);$book.Close($false);$book=$null;$srcBook.Close($false);$srcBook=$null
 $externalScript=Join-Path $PSScriptRoot 'apply_keyuan_external_sources.py';$externalJson=& $python -X utf8 $externalScript $formal $SampleDirectory;if($LASTEXITCODE -ne 0){throw '社保或个税附件处理失败'};$externalChanges=$externalJson|ConvertFrom-Json;foreach($x in $externalChanges){$changes += [ordered]@{sheet=$x.sheet;cell=$x.cell;before=$x.before;after=$x.after;source=$x.source;rule=$x.rule}}
 $book=$app.Workbooks.Open($formal,0,$false);try{try{$app.Calculation=-4105}catch{};try{$book.ForceFullCalculation=$true}catch{};try{$app.CalculateFullRebuild()}catch{try{$app.Calculate()}catch{}};$book.Save()}finally{$book.Close($false);Release-Com $book;$book=$null}
 Copy-Item -LiteralPath $formal -Destination $review -Force;$reviewBook=$app.Workbooks.Open($review,0,$false);$log=$reviewBook.Worksheets.Add();$log.Name='修改记录';$headers=@('序号','工作表','单元格','修改前','修改后','来源','规则');for($c=1;$c -le $headers.Count;$c++){$log.Cells.Item(1,$c).Value2=[string]$headers[$c-1]};$rr=2;foreach($x in $changes){$log.Cells.Item($rr,1).Value2=[string]($rr-1);$log.Cells.Item($rr,2).Value2=[string]$x.sheet;$log.Cells.Item($rr,3).Value2=[string]$x.cell;$log.Cells.Item($rr,4).Value2=[string]$x.before;$log.Cells.Item($rr,5).Value2=[string]$x.after;$log.Cells.Item($rr,6).Value2=[string]$x.source;$log.Cells.Item($rr,7).Value2=[string]$x.rule;try{$reviewBook.Worksheets.Item($x.sheet).Range($x.cell).Interior.Color=65535}catch{};$rr++};$log.Rows.Item(1).Font.Bold=$true;$log.Columns.AutoFit()|Out-Null;$reviewBook.Save();$reviewBook.Close($false);Release-Com $log;Release-Com $reviewBook
 $ib=$app.Workbooks.Add();$is=$ib.Worksheets.Item(1);$is.Name='待人工确认';$is.Cells.Item(1,1).Value2='事项';$is.Cells.Item(1,2).Value2='说明';$is.Rows.Item(1).Font.Bold=$true;$rr=2;if($issues.Count -eq 0){$is.Cells.Item(2,1).Value2='无阻塞项';$is.Cells.Item(2,2).Value2='本次样本已按手册自动处理'}else{foreach($x in $issues){$is.Cells.Item($rr,1).Value2=$x.item;$is.Cells.Item($rr,2).Value2=$x.detail;$rr++}};$is.Columns.AutoFit()|Out-Null;$ib.SaveAs($issuesPath,51);$ib.Close($false);Release-Com $is;Release-Com $ib
}finally{if($null -ne $srcBook){try{$srcBook.Close($false)}catch{};Release-Com $srcBook};if($null -ne $book){try{$book.Close($false)}catch{};Release-Com $book};if($null -ne $app){try{$app.Quit()}catch{};Release-Com $app};[gc]::Collect();[gc]::WaitForPendingFinalizers();if(Test-Path $working){Remove-Item -LiteralPath $working -Force}}
$report=[ordered]@{status=if($issues.Count){'needs_review'}else{'passed'};generated_at=(Get-Date).ToString('s');source_master=$master.FullName;source_files_modified=$false;formal_workbook=$formal;review_workbook=$review;unresolved_workbook=$issuesPath;change_count=$changes.Count;changes=$changes;unresolved_count=$issues.Count;issues=$issues};$report|ConvertTo-Json -Depth 8|Set-Content -LiteralPath $reportPath -Encoding utf8;$report|ConvertTo-Json -Depth 6

