package nimsoft_generic;
#----------------------------------------------------------------
# read UIM Restful Server, Port, Username and Password from file-
#----------------------------------------------------------------
sub getRestData 
{

  my $uim_server = "x";
  my $uim_port = "x";
  my $uim_https = "x";
  my $uim_user = "x";
  my $uim_password = "x";
  my $uim_domain = "x";
  my $uim_hub = "x";
  my $uim_robot = "x";
  my $uim_realm = "x";
  my $sql_server = "x";
  my $sql_user = "x";
  my $sql_password = "x";
  my $sql_db ="x";
  my $sql_type ="x";
  my $sql_driver ="x";
  my $sql_dsn ="x";

# - do we find the nimsoft_generic.dat file?

  $genericdat="nimsoft_generic.dat";
  unless (-e $genericdat) 
   { 
     print "File: $genericdat Doesn't Exist!\n"; 
     print "(this file contains all needed connection parameters for this script)\n";
     exit 2; 
   }
  
# -
  if (open( $FP, $genericdat)) 
  {
    while ($entry = <$FP>) 
    {
      chomp $entry;
      #print "$entry\n";
	  my ($key, $value) = split("=", $entry);
	  if ($key eq "uim_server") { $uim_server = $value; $uim_server =~ s/^\s+|\s+$//g;};
	  if ($key eq "uim_port") { $uim_port = $value; $uim_port =~ s/^\s+|\s+$//g;};
	  if ($key eq "uim_https") { $uim_https = $value; $uim_https =~ s/^\s+|\s+$//g;};
	  if ($key eq "uim_user") { $uim_user = $value; $uim_user =~ s/^\s+|\s+$//g;};
	  if ($key eq "uim_password") { $uim_password = $value; $uim_password =~ s/^\s+|\s+$//g;};
	  if ($key eq "uim_domain") { $uim_domain = $value; $uim_domain =~ s/^\s+|\s+$//g;};
	  if ($key eq "uim_hub") { $uim_hub = $value; $uim_hub =~ s/^\s+|\s+$//g;};
	  if ($key eq "uim_robot") { $uim_robot = $value; $uim_robot =~ s/^\s+|\s+$//g;};
	  if ($key eq "uim_realm") { $uim_realm = $value; $uim_realm =~ s/^\s+|\s+$//g;};
	  if ($key eq "sql_server") { $sql_server = $value; $sql_server =~ s/^\s+|\s+$//g;};
	  if ($key eq "sql_user") { $sql_user = $value; $sql_user =~ s/^\s+|\s+$//g;};
	  if ($key eq "sql_password") { $sql_password = $value; $sql_password =~ s/^\s+|\s+$//g;};
	  if ($key eq "sql_db") { $sql_db = $value; $sql_db =~ s/^\s+|\s+$//g;};
	  if ($key eq "sql_type") { $sql_type = $value; $sql_type =~ s/^\s+|\s+$//g;};
	  if ($key eq "sql_driver") { $sql_driver = $value; $sql_driver =~ s/^\s+|\s+$//g;};
	  if ($key eq "sql_dsn") { $sql_dsn = $value; $sql_driver =~ s/^\s+|\s+$//g;};
     }
  }
  else
  {
    print "Error during open of file: $genericdat\n";
    print "(this file contains all needed connection parameters for this script)\n";
    exit 3;
  }  

# -
  return ($uim_server, $uim_port, $uim_user, $uim_password, $uim_domain, $uim_hub, $uim_robot, $sql_server, $sql_user, $sql_password, $sql_db, $sql_type, $uim_https, $sql_driver, $sql_dsn, $uim_realm);
}

1; 