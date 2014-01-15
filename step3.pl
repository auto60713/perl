use strict;
use Win32::OLE qw(in with);
use Win32::OLE::Const 'Microsoft Excel';
use FindBin;
my $path = $FindBin::Bin;


print "點數卡庫存(XXXX)\n請輸入結算的日期  XXXX = ?    ";
chomp(my $date=<STDIN>);
#宣告$file 並帶入使用者輸入

my $Excel = Win32::OLE->GetActiveObject('Excel.Application')|| Win32::OLE->new('Excel.Application', 'Quit');
   #$Excel ->{Visible} = 1;
   #不設定Visible excel視窗將不顯現

my $book    =  $Excel -> Workbooks -> Open (  $path."/點數卡庫存(".$date.").xls" );
my $book2   =  $Excel -> Workbooks -> Open (  $path."/點數卡結算.xls" );

#資料是在第二個工作表
my $Sheet1  =  $book  ->  Worksheets(2) ;
my $Sheet2  =  $book2 ->  Worksheets(1) ;



my %gameStock;
#宣告一個hash陣列 來存放點數卡的剩餘數量

#設置一個最大數目 資料量不能超過這個數目
foreach  my  $row  (  2  ..  1000  ){  

    my $gameName =  $Sheet1 -> Cells ( $row , 2 ) -> { Value };
    last  unless  defined  $gameName;  
    #如果沒有值就表示結束了

    my $stock =  $Sheet1 -> Cells ( $row , 7 ) -> { Value };

    $gameStock{$gameName} += $stock;
    #這時會在hash陣列裡面成立以gameName命名的變數 並且給他帶入相對應的庫存數


}

foreach  my  $row  (  2  ..  20  ){  

    my $gameName =  $Sheet2 -> Cells ( $row , 1 ) -> { Value };
    last  unless  defined  $gameName;  
    #如果沒有值就表示結束了

    $Sheet2->Cells ( $row , 2 )->{Value} =  $gameStock{$gameName};  
    #抓取第一欄的遊戲名稱 在第二欄寫入該遊戲名稱的庫存

}

print "\n更新完成!!";

#儲存檔案
$book2 -> Save();
$book  -> Close ();
$book2 -> Close ();
$Excel -> Quit ();

