use strict;


&happy("���ߵo�]�B�s�~�ּ�!!",8);
#�I�s�Ƶ{��happy �ña�J��ӰѼ�

#�ŧi�Ƶ{��happy
sub happy {

    my ($str,$times) = @_;
    #������ӰѼ� �ëŧi������ܼ�

    foreach  my  $row  (  1  ..  $times ){   

    print $str."\n";
    }

}