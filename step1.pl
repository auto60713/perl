use strict;
use Win32::OLE qw(in with);
use Win32::OLE::Const 'Microsoft Excel';
use FindBin;
my $path = $FindBin::Bin;
#使用FindBin可以得到當前地址

my $Excel = Win32::OLE->GetActiveObject('Excel.Application')|| Win32::OLE->new('Excel.Application', 'Quit');
   $Excel ->{Visible} = 1;
#要使用excel前的宣告 Visible=1會顯示excel視窗

my $book    =  $Excel -> Workbooks -> Open (  $path."/book.xls"  );
#用Excel變數打開檔案並宣告成book變數 
my $Sheet1  =  $book ->  Worksheets(1) ;
#選擇book的工作表1並宣告成Sheet1變數


foreach  my  $row  (  2  ..  30  ){  

my $nember =  $Sheet1 -> Cells ( $row , 1 ) -> { Value };
#宣告變數nember 並且等於Sheet1工作表( 第N列 , 第1欄 )的值
my $name  =  $Sheet1 -> Cells ( $row , 2 ) -> { Value };
my $score =  $Sheet1 -> Cells ( $row , 3 ) -> { Value };

next  unless  defined  $score;  #如果沒值就結束這次迴圈
#所以沒有分數的學生不會被印出來

print "學號 ".$nember."   姓名 ".$name."   分數 ".$score."\n";

last  unless  defined  $nember;  #如果沒值就跳出這個迴圈
#所以李祥後面的人都沒被印出來
}





$book -> Close ();
$Excel -> Quit ();

