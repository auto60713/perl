use strict;
use Win32::OLE qw(in with);
use Win32::OLE::Const 'Microsoft Excel';
use FindBin;
my $path = $FindBin::Bin;


print "�п�J�n�X���ɮצW��(���ΰ��ɦW)\n";
chomp(my $file=<STDIN>);
#�ŧi$file �ña�J�ϥΪ̿�J

my $Excel = Win32::OLE->GetActiveObject('Excel.Application')|| Win32::OLE->new('Excel.Application', 'Quit');
   #$Excel ->{Visible} = 1;
   #���]�wVisible excel�����N����{

my $book    =  $Excel -> Workbooks -> Open (  $path."/book.xls" );
my $Sheet1  =  $book ->  Worksheets(1) ;

my $titleScore = 0;
#�ŧi�@���ܼƨӦs���`��

foreach  my  $row  (  2  ..  21  ){  

    my $score =  $Sheet1 -> Cells ( $row , 3 ) -> { Value };
    next  unless  defined  $score;  

    $titleScore += $score;
    #���`�����W�[

    if($score >= 60){
    	$Sheet1->Cells ( $row , 4 )->{Value} = "O";
    }
    else{
    	$Sheet1->Cells ( $row , 4 )->{Value} = "X";
    }

}

$Sheet1->Cells ( 23 , 4 )->{Value} = $titleScore;
#��X�`��



#$book -> Save();
#�x�s�ɮ�
$book -> SaveAs( $path."/"."$file.xls" );
#�t�s�s��
$book -> Close ();
$Excel -> Quit ();

