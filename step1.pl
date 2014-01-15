use strict;
use Win32::OLE qw(in with);
use Win32::OLE::Const 'Microsoft Excel';
use FindBin;
my $path = $FindBin::Bin;
#�ϥ�FindBin�i�H�o���e�a�}

my $Excel = Win32::OLE->GetActiveObject('Excel.Application')|| Win32::OLE->new('Excel.Application', 'Quit');
   $Excel ->{Visible} = 1;
#�n�ϥ�excel�e���ŧi Visible=1�|���excel����

my $book    =  $Excel -> Workbooks -> Open (  $path."/book.xls"  );
#��Excel�ܼƥ��}�ɮרëŧi��book�ܼ� 
my $Sheet1  =  $book ->  Worksheets(1) ;
#���book���u�@��1�ëŧi��Sheet1�ܼ�


foreach  my  $row  (  2  ..  30  ){  

my $nember =  $Sheet1 -> Cells ( $row , 1 ) -> { Value };
#�ŧi�ܼ�nember �åB����Sheet1�u�@��( ��N�C , ��1�� )����
my $name  =  $Sheet1 -> Cells ( $row , 2 ) -> { Value };
my $score =  $Sheet1 -> Cells ( $row , 3 ) -> { Value };

next  unless  defined  $score;  #�p�G�S�ȴN�����o���j��
#�ҥH�S�����ƪ��ǥͤ��|�Q�L�X��

print "�Ǹ� ".$nember."   �m�W ".$name."   ���� ".$score."\n";

last  unless  defined  $nember;  #�p�G�S�ȴN���X�o�Ӱj��
#�ҥH�����᭱���H���S�Q�L�X��
}





$book -> Close ();
$Excel -> Quit ();

