use strict;


&happy("恭喜發財、新年快樂!!",8);
#呼叫副程式happy 並帶入兩個參數

#宣告副程式happy
sub happy {

    my ($str,$times) = @_;
    #接取兩個參數 並宣告成兩個變數

    foreach  my  $row  (  1  ..  $times ){   

    print $str."\n";
    }

}