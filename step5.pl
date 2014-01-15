use strict;
use Win32::OLE qw(in with);
use Win32::OLE::Const 'Microsoft Excel';
use FindBin;
my $path = $FindBin::Bin;


my $Excel = Win32::OLE->GetActiveObject('Excel.Application')|| Win32::OLE->new('Excel.Application', 'Quit');

my $book  =  $Excel -> Workbooks -> Open (  $path."/book.xls" );

my $Sheet1  =  $book  ->  Worksheets(2) ;
#����ĤG�Ӥu�@�����


#�S��Ÿ��n��\�ӻ��U �H�K�D�S���Ū
print "\n�P�_�r��O�_�]�t\"ACG\"==========\n\n";
foreach  my  $row  (  2  ..  10  ){  

    my $val =  $Sheet1 -> Cells ( $row , 1 ) -> { Value };
    last  unless  defined  $val;  

    if($val =~ "ACG"){

        print "\"".$val."\" �����]�t�FACG\n";
    }
}


print "\n�P�_�r��}�Y�O�_��AE==========\n\n";
foreach  my  $row  (  2  ..  10  ){  

    my $val =  $Sheet1 -> Cells ( $row , 2 ) -> { Value };
    last  unless  defined  $val;  

    if($val =~ m/^AE/){

       print "\"".$val."\" �OAE���}�Y���r��\n";
    }
}


print "\n�h���r��᭱�h�󪺪ť�==========\n\n";
foreach  my  $row  (  2  ..  10  ){  

    my $val =  $Sheet1 -> Cells ( $row , 3 ) -> { Value };
    last  unless  defined  $val;  
    #�p�G�S���ȴN��ܵ����F

    $val =~ s/\s+$//;
    print $val;
}
    print "\n�]���r���s�b�@�_��ܪťդw�g�h���F\n";



print "\n���Φr��==========\n\n";
foreach  my  $row  (  2  ..  10  ){  

    my $val =  $Sheet1 -> Cells ( $row , 4 ) -> { Value };
    last  unless  defined  $val;  
    #�p�G�S���ȴN��ܵ����F

    my @cut = split('\.', $val);  
    #�H.���}�r�� ���}��s���}�C

    print $cut[1]."��".$cut[2]."��\n";
    #�ڭ̥u�����
}