use strict;
use Win32::OLE qw(in with);
use Win32::OLE::Const 'Microsoft Excel';
use FindBin;
my $path = $FindBin::Bin;


print "請輸入要出表的檔案名稱(不用副檔名)\n";
chomp(my $file=<STDIN>);
#宣告$file 並帶入使用者輸入

my $Excel = Win32::OLE->GetActiveObject('Excel.Application')|| Win32::OLE->new('Excel.Application', 'Quit');
   #$Excel ->{Visible} = 1;
   #不設定Visible excel視窗將不顯現

my $book    =  $Excel -> Workbooks -> Open (  $path."/book.xls" );
my $Sheet1  =  $book ->  Worksheets(1) ;

my $titleScore = 0;
#宣告一個變數來存放總分

foreach  my  $row  (  2  ..  21  ){  

    my $score =  $Sheet1 -> Cells ( $row , 3 ) -> { Value };
    next  unless  defined  $score;  

    $titleScore += $score;
    #把總分往上加

    if($score >= 60){
    	$Sheet1->Cells ( $row , 4 )->{Value} = "O";
    }
    else{
    	$Sheet1->Cells ( $row , 4 )->{Value} = "X";
    }

}

$Sheet1->Cells ( 23 , 4 )->{Value} = $titleScore;
#算出總分



#$book -> Save();
#儲存檔案
$book -> SaveAs( $path."/"."$file.xls" );
#另存新檔
$book -> Close ();
$Excel -> Quit ();

