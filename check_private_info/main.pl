#package Horn;
use Encode;
use Net::SMTP;
use MIME::Base64;
use Spreadsheet::ParseExcel;

#put config.txt nearby

my $args   = read_config();
my %args   = %$args;
my $id_msg = crop_xls(
	$args{"info_path"},  $args{"info_sheet"},
	$args{"info_value"}, $args{"info_key"}
);

my $id_address =
  crop_xls( $args{"db_path"}, $args{"db_sheet"},, $args{"db_value"},
	$args{"db_key"} );
my %id_msg     = %$id_msg;
my %id_address = %$id_address;

foreach $id ( keys(%id_msg) ) {
	my $to        = $id_address{$id};
	my $mail_body = $id_msg{$id};
	if ( $args{'debug'} eq '1' || $to eq "" ) {

		$to = $args{'auth'};

	}

	send_mail( $args{'host'}, $args{'auth'}, $args{'passwd'}, $to,
		$args{'subject'}, $mail_body );

}

sub crop_xls() {
	my ( $filename, $iSheet, $arrange, $id_index ) = @_;
	my %out;

	#printf("$filename, $iSheet, $arrange, $id_index \n");
	my ( $left, $top, $right, $bottom ) = split( ',', $arrange );
	my $oExcel = new Spreadsheet::ParseExcel;
	my $oBook  = $oExcel->Parse("$filename");

	#print "FILE  :",      $oBook->{File},       "\n";
	#print "SHEETCOUNT :", $oBook->{SheetCount}, "\n";

	my ( $iR, $oWkS, $oWkC );

	$oWkS = $oBook->{Worksheet}[ $iSheet - 1 ];
	for (

		my $iR = $oWkS->{MinRow} + ( $top - 1 ) ;
		defined $oWkS->{MaxRow} && $iR <= $oWkS->{MinRow} + ( $bottom - 1 ) ;
		$iR++
	  )
	{

		my $msg = "";
		for (
			my $i = ord($left) - ord('A') ;
			$i <= ord($right) - ord('A') ;
			++$i
		  )
		{

			my $oWkC = $oWkS->{Cells}[$iR][$i];

			if ($oWkC) {
				my $field = $oWkC->Value;
				if ( $field ne "" ) {
					$msg = $msg . "$field ";
				}
			}

		}
		if ( $msg ne "" ) {
			$id = $oWkS->{Cells}[$iR][ ord($id_index) - ord('A') ]->Value;
			$msg =~ s/\n/ /g;
			$msg =~ s/\s+$//g;

			#printf("[$id]=> [$msg]\n");
			$out{$id} = $msg;

		}

	}
	return \%out;
}

sub read_config($) {
	my $SCRIPT = __FILE__;
	my $dir    = $SCRIPT;
	$dir =~ s/[\/\\][^\/\\]*$//;

	my %args;

	open( HANDLE,      "config.txt" )
	  or open( HANDLE, "$dir/config.txt" )
	  or die "config.txt not found at $dir/config.txt";
	while (<HANDLE>) {
		if (/^#/) {
			next;
		}
		{
			if (/([^\s]*)\s*\=\s*\"(.*)\"/) {

				#$1 =~ s/\\/\//g;
				# $1 is a const value
				my $tmp = $1;
				$tmp =~ s/\\/\//g;
				$args{$tmp} = $2;

			}
		}
	}
	close HANDLE;
	return ( \%args );
}

sub send_mail {
	my ( $host, $auth, $password, $to, $subject, $mail_body ) = @_;
	printf("send to $to [$mail_body]\n");
	{
		$mail_body = $mail_body;
		my $smtp = Net::SMTP->new(
			Host    => $host,
			Debug   => 1,
			Timeout => 30
		);

		#print "$host $auth $to $mail_body\n";
		$smtp->command('AUTH LOGIN')->response();
		my $userpass = encode_base64($auth);
		$userpass =~ s/\n//g;
		$smtp->command($userpass)->response();
		$userpass = encode_base64($password);
		$userpass =~ s/\n//g;
		$smtp->command($userpass)->response();
		$smtp->mail($auth);
		$smtp->to($to);
		$smtp->bcc($auth);
		$smtp->data();
		$smtp->datasend("Content-Type:text/plain;charset=utf-8\n");
		$smtp->datasend("Content-Transfer-Encoding:utf-8\n");
		$smtp->datasend("From:$auth \n");
		$smtp->datasend("To:$to \n");
		$smtp->datasend(
			"Subject:=?gb2312?B?" . encode_base64( $subject, '' ) . "?=\n\n" );
		$smtp->datasend("\n");
		$smtp->datasend( $mail_body, '' . " \n" );
		$smtp->dataend();
		$smtp->quit;
	}
}
__END__
config.txt like this

#------------------------------------------------------------
#隐私信息文件路径
info_path = "C:\Users\Administrator\Desktop\check_private_info\旅游登记.xls"

# 表单序号
info_sheet = "1"

# 信息范围：如 (A,1,D,16) （行从A开始，列从1开始。依次为左,上,右,下）
info_value = "B,3,K,4"

# 工号所在列
info_key = "B"


#-------------------------------------------------------------
# 邮箱数据库文件路径
db_path = "C:\Users\Administrator\Desktop\check_private_info\邮件统计.xls"

# 表单序号
db_sheet = "1"

# 工号所在列
db_key = "B"

# 邮箱地址范围
db_value ="C,3,C,159"



#-----------------------------------------------------------
# 邮箱账户设置
host ="smtp.oppo.com"
auth = "zhr@oppo.com"
passwd = "******"
subject = "Please check your information"
# 调试模式 (1 or 0) 当为1时邮件将全部发送给auth.
debug = "1"
#end
