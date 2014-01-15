use strict;
use Win32::OLE qw(in with);
use Win32::OLE::Const 'Microsoft Excel';
use FindBin;
my $path = $FindBin::Bin;


print "�I�ƥd�w�s(XXXX)\n�п�J���⪺���  XXXX = ?    ";
chomp(my $date=<STDIN>);
#�ŧi$file �ña�J�ϥΪ̿�J

my $Excel = Win32::OLE->GetActiveObject('Excel.Application')|| Win32::OLE->new('Excel.Application', 'Quit');
   #$Excel ->{Visible} = 1;
   #���]�wVisible excel�����N����{

my $book    =  $Excel -> Workbooks -> Open (  $path."/�I�ƥd�w�s(".$date.").xls" );
my $book2   =  $Excel -> Workbooks -> Open (  $path."/�I�ƥd����.xls" );

#��ƬO�b�ĤG�Ӥu�@��
my $Sheet1  =  $book  ->  Worksheets(2) ;
my $Sheet2  =  $book2 ->  Worksheets(1) ;



my %gameStock;
#�ŧi�@��hash�}�C �Ӧs���I�ƥd���Ѿl�ƶq

#�]�m�@�ӳ̤j�ƥ� ��ƶq����W�L�o�Ӽƥ�
foreach  my  $row  (  2  ..  1000  ){  

    my $gameName =  $Sheet1 -> Cells ( $row , 2 ) -> { Value };
    last  unless  defined  $gameName;  
    #�p�G�S���ȴN��ܵ����F

    my $stock =  $Sheet1 -> Cells ( $row , 7 ) -> { Value };

    $gameStock{$gameName} += $stock;
    #�o�ɷ|�bhash�}�C�̭����ߥHgameName�R�W���ܼ� �åB���L�a�J�۹������w�s��


}

foreach  my  $row  (  2  ..  20  ){  

    my $gameName =  $Sheet2 -> Cells ( $row , 1 ) -> { Value };
    last  unless  defined  $gameName;  
    #�p�G�S���ȴN��ܵ����F

    $Sheet2->Cells ( $row , 2 )->{Value} =  $gameStock{$gameName};  
    #����Ĥ@�檺�C���W�� �b�ĤG��g�J�ӹC���W�٪��w�s

}

print "\n��s����!!";

#�x�s�ɮ�
$book2 -> Save();
$book  -> Close ();
$book2 -> Close ();
$Excel -> Quit ();

