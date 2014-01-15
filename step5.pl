use strict;
use Win32::OLE qw(in with);
use Win32::OLE::Const 'Microsoft Excel';
use FindBin;
my $path = $FindBin::Bin;


my $Excel = Win32::OLE->GetActiveObject('Excel.Application')|| Win32::OLE->new('Excel.Application', 'Quit');

my $book  =  $Excel -> Workbooks -> Open (  $path."/book.xls" );

my $Sheet1  =  $book  ->  Worksheets(2) ;
#抓取第二個工作表的資料


#特殊符號要用\來輔助 以免遭特殊解讀
print "\n判斷字串是否包含\"ACG\"==========\n\n";
foreach  my  $row  (  2  ..  10  ){  

    my $val =  $Sheet1 -> Cells ( $row , 1 ) -> { Value };
    last  unless  defined  $val;  

    if($val =~ "ACG"){

        print "\"".$val."\" 之中包含了ACG\n";
    }
}


print "\n判斷字串開頭是否為AE==========\n\n";
foreach  my  $row  (  2  ..  10  ){  

    my $val =  $Sheet1 -> Cells ( $row , 2 ) -> { Value };
    last  unless  defined  $val;  

    if($val =~ m/^AE/){

       print "\"".$val."\" 是AE為開頭的字串\n";
    }
}


print "\n去除字串後面多於的空白==========\n\n";
foreach  my  $row  (  2  ..  10  ){  

    my $val =  $Sheet1 -> Cells ( $row , 3 ) -> { Value };
    last  unless  defined  $val;  
    #如果沒有值就表示結束了

    $val =~ s/\s+$//;
    print $val;
}
    print "\n因為字都連在一起表示空白已經去除了\n";



print "\n切割字串==========\n\n";
foreach  my  $row  (  2  ..  10  ){  

    my $val =  $Sheet1 -> Cells ( $row , 4 ) -> { Value };
    last  unless  defined  $val;  
    #如果沒有值就表示結束了

    my @cut = split('\.', $val);  
    #以.切開字串 切開後存成陣列

    print $cut[1]."月".$cut[2]."號\n";
    #我們只取日期
}