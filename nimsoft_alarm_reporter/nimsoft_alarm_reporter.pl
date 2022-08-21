#
# read mysql/mssql db CA_UIM table nas_transaction_log or nas_transaction_summary (and optionally the site wide: AlarmTransactionLog table)
#
# create HTML/CSV report with all current alarms
#
# In NAS setup you need: (in NIS Bridge tab)
# - activate NIS bridge
# - log transaction details
# 
# input parameters:
#   -di: output directory (default: c:\temp)
#   -fi: output file (default: report_nimsoft_alarm_reporter)
#   -rs: (optional) number of seconds between report + html updates
#   -su: (nimsoft_generic.dat) sql userid (default: mysql: root mssql: sa) if you use -u"trusted" we will use Windows trusted authentication
#   -sp: (nimsoft_generic.dat) sql password *** required ***
#   -sr: (nimsoft_generic.dat) repository/server name: default local host
#   -wr: create overvieW report (Y, N) Default Y
#   -ty: (nimsoft_generic.dat) 1: mssql 2: mysql (default: 1)
#   -al: y: query mssql AlarmTransactionLog table (only after manual steps to create this table) Default: n
#           (remark that this table has not all fields available for reporting & that we use a "customized" version of the proposed table that has some important fields added)
#      : t: DEFAULT: nas_transaction_log
#      : s: nas_transaction_summary
#   -nb: maximum # lines printed on detail report (to avoid that IE doesn't open) Default: 5000
#   -8: (nimsoft_generic.dat) database name (default CA_UIM)
#
# Date Selection:
#   -bh: (optional) number of hours to go back to start report (default 24, only today) 48: yesterday & today 72:....
#   -bm: (optional) number of minutes to go back (default 60)
#   -mm:(optional) number of months to go back (0: current month, 1: previous month,...)
#   -1: (optional) start date of report, format: "yyyy-mm-dd hh:mm"
#   -2: (optional) start time, format: "hh:mm" (works only with -bh)
#   -3: (optional) end date of report, format: "yyyy-mm-dd hh:mm"
#   -4: (optional) end time, format: "hh:mm" (works only with -bh)
#
# OR you use -bh (= -b) or -bm, OR you use -1 and -3
# If no date option is selected, the default is -bh"24" (= report on the last 24 hours)
#
# Report Columns:
#   -co: give the columns you want in your report: 
#        duration,hub,origin,robot,source,subsys,probe,severity,level,visible,suppcount,message,suppkey,sid,acknowledged,user_tag1,user_tag2,assigned_by,assigned_to,type
#
# SQL where like filters (to limit the amount of selected records)
#   -lh: like hostname (as where like clause in sql)\n";
#   -lg: like origin (as where like clause in sql)\n";
#   -lm: like message (as where like clause in sql)\n";
#   -lo: like probe (as where like clause in sql)\n";
#   -lu: like hub (as where like clause in sql)\n";
#   -lt: like robot (as where like clause in sql)\n";
#
# Search Criteria: (regex)
#   -mi/me: message text include/exclude
#   -gi/ge: oriGin include/exclude
#   -ui/ue: hUb include/exclude
#   -hi/he: hostname include/exclude
#   -ti/te: roboT include/exclude
#   -oi/oe: prObe include/exclude
#   -ai/ae: assigned_by include/exclude
#   -si/se: assigned_to include/exclude
#   -ki/ke: acknowledged_by include/exclude
#   -li/le: level include/exclude
#
#   -vi: VIsible y: show visible (default) n: no visible reported o: only visible
#
# output:
#   -9 : view report directly (y or n) Default n
#   -x : create csv file (instead of html file) (work in progress)
#
# - HTML report in directory -d and file -f (by default: c:\temp\report_nimsoft_alarm_reporter.html)
# 
#
# Count number of events in table: select COUNT(*) from nas_transaction_log
#
# db: C:\Users\All Users\MySQL\MySQL Server 5.6\data\nimsoftslm
#
# Updated:
# - 25/03/14: add visible keyword
#             add supp_key field
# - 26/03/14: remove milliseconds 
#           : added duration of the outstanding message
# - 10/04/14: add overview counters (-w"Y")
# - 10/04/14: add option -y to make the selection mssql or mysql
# - 14/04/14: create overview report (instead of overview print)
# - 17/04/14: add usage when no parameters are passed
#           : reset counters when executing in loop 
# - 18/04/14: added suppcount in order statement (in case same message in same second)
#           : add sid column
# - 25/04/14: if mssql and no userid/password, try windows integrated security
# - 07/05/14: add -a: query from mssql AlarmTransactionLog table (common ado mesage table)
# - 20/05/14: use updated version of the AlarmTransactionLog to have the fields: domain, hub, robot and probe
#           : updated exclude count if no filters 
# - 23/06/14: added binmode for output files to avoid 'wide characters' warnings
# - 26/06/14: remove (possible) new line characters from messages
# - 27/06/14: added some use statements to compile this perl (so it can be used on systems without perl or perl without all modules)
# - 26/07/14: correct -o (probe selector)
# - 28/07/14: -bh is now in hours
#           : exit code is # of selected messages    
# - 12/08/14: -n to specify the maximum # of lines in detail report, this to avoid that your browser doesn't open correctly    
# - 27/08/14: add field acknowledged_by  
# - 29/08/14: -c gives the possibility to name the columns you want/need in your report (duration,robot,source,subsys,probe,severity,level,visible,suppcount,message,suppkey,sid,acknowledged)
# - 22/09/14: corrected the -s loop logic 
#           : corrected report column count variable
# - 26/09/14: add option -8 to define database name (default: CA_UIM)
# - 01/10/14: add option -1 to define report start date/time
#           : add option -3 to define report end date/time (if only -1 is used the end date/time is now)
# - 24/10/14: changes switch names to be more logical
#           : add host exclude (-he), robot exclude (-te) and probe exclude (-oe)
# - 05/11/14: add overview report by date and by hour
#           : use only DBI
# - 09/11/14: add -c columns user_tag1/2
# - 17/11/14: add -c column hub
#           : add -ui and -ue hUb include/exclude
# - 11/12/14: add -gi and -ge oriGin include/exclude as filter + origin column
#           : when reporting from alarmtransactionlog remove blanks from the reported fields
# - 24/03/15: invalid -u switch when not sa used
# - 25/09/15: add nimsoft_generic.pm to get common parameters (logon, servername...)
# - 23/10/15: -9: view generated report directly (y,n) default n
# - 13/11/15: -x: create csv file (y) or html report (n)
# - 07/12/15: nimsoft_generic.dat will only use passwords that are crypted via nimsoft_crypt.exe
# - 06/01/16: -c becomes -co
#           : option -cv (CountValue) makes it possible to create a regex filter on the suppression count (ex. "^1$" to have only the new messages)
# - 07/01/16: option -2 defines a start hour (hh:mm) works only if other date parameters are also used 
#           : option -4 defines an end hour (hh:mm) works only if other date parameters are used.  This makes it possible to create a report of the previous dates via -bh but only reporting on the working hours via -2 and -4
# - 18/01/16: add new overview report that will show the messages and probe
# - 20/01/16: add sql_type in nimsoft_generic
# - 22/01/16: add version number in last line of report
# - 21/02/16: add -ai (assigned_by include) and -ae (assigned_by exclude) regex filters (to select all assigned alarms: -si"(.|\s)*\S(.|\s)*")
#             add -si (assigned_to include) and -se (assigned_to exclude) regex filters 
# - 22/02/16: add -ki (acknowledged_by include) and -ke (acknowledged_by exclude) regex filters
#           : add -li (level include) and -le (level exclude) regex filters
# - 15/03/16: you can specify up to 18 user columns in report
# - 05/05/16: to select only the first alarms you can use -cv"1" (this will be tested against "", "0" and "1")
#           : add reporting of nas_transaction_summary
#           : add overview report: acknowledged_by, assigned_by and assigned_to  
#           : when reporting on nas_transaction_summary (-al"s") you can use as columns: custom1, custom2, custom3, custom4 and custom5
# - 27/06/16: mysql only -bh is validated
# - 05/07/16: add assigned and acknowledged in the report header filter
# - 18/11/16: test if we run on MSWin32, else take other defaults (-di"/opt/temp")
#           : change report rename to move (so that it works on Linux)
# - 28/11/16: tested on Windows->mssql, Centos->Mysql, Compiled version->Windows&Centos
# - 02/12/16: add some extra used regex filter is title of report
# - 11/01/17: in the overview reports we replaced all figures with a x, but when more than 9 it was replaced as xx => now always replaced as xx
# - 02/02/17: add use_https in nimsoft_generic.dat/pm
# - 28/04/17: add with(nolock) for all tables in all sql queries
# - 12/05/17: when printing to csv we received (some) error messages when encountering non-ascii values
# - 11/09/17: -co was not always accepted to define your own columns
# - 11/10/17: -bh was not accepting correctly the # of hours
# - 08/11/17: -mm: report on a month (o: current month, 1: previous month)
#           : csv output file has several new fields: domain, hub, origin, user_tag, custom1-5
# - 22/01/18: -bm: report on the xx last minutes
#             -bh or -b for the last xx hours
# - 01/05/19: add some help fields in color (as test)
# - 17/05/19: added source column in default report
# - 13/06/19: adapted colors for the severity column to correspond with the colors used in alarm subconsole
# - 06/09/19: tested with version 9.20
# - 16/10/19: add parameter $sql_driver in nimsoft_generic.dat so that its possible to install and reference "ODBC Driver 17 for SQL Server" that should be compatible with TLS 1.2
# - 11/02/20: -db"y" running debug mode 
# - 24/02/20: if only -2 (start hour) or -4 (end hour) are used to indicate reporting period, by default it will be todays date
# - 13/03/20: If you need/want to use a specific, pre-defined ODBC System DSN, you can use the nimsoft_generic.dat parameter sql_dsn with as value the ODBC defined name
# - 06/05/20: add overview report based on source & probe # of alarms
# - 25/10/20: tested with uim 20.3
# - 19/11/20: tested on Perl 5.32
# - 16/05/22: nimsoft_generic.pm and nimsoft_generic.dat can be in same  directory as perl source (or in perl/lib)
#           : tested on 20.4
# - 15/06/22: add sql WHERE clause parameters: -lg (origin), -lh (hostname), -lo (probe), -lm (message), -lt (robot) and -lu (hub)
#           : add selected sql and regex filters in report header and in Field Explanation
#           : changed the column source with hostname in report (source is sometimes ip address and hostname is translated name)
#           : added hostname as possible column for -co
# - 16/06/22: add also sql filters: -lt (robot) and -lu (hub)
#           : for a first run without nimsoft_generic.dat, you can use -sr (sql server), -su (sql user) and -sp (sql password)  
#           : -fq can strip fqdn domain name for the hostname column y: strip n: keep original, default: y
#           : added origin in default -co report columns
#
# Note:
# MySQL can be a little difficult to accept your userid and password, this must match exactly what is defined in the GUI - Users and Privileges - nodes
#
#
# ---------------------------------------------------------------------------------------------------------------------------------------------------------------------
# - This program is only an example program that can be used on your own risk
# ---------------------------------------------------------------------------------------------------------------------------------------------------------------------
#

# --- Import external modules
use DBI;
use PerlIO;
use Getopt::Long;
# use Posix;
use DBD::ODBC;
# crypt
use Crypt::RC4;
use MIME::Base64;

# - to find if an exe is available
use File::Which;
use File::Which qw (which where);
use File::Copy;
# - to get the current os type
use Config;
# --- to use colors in the print statements
use Win32::Console::ANSI;
use Term::ANSIColor;
# - change directory to directory where perl source is located to find our nimsoft_generic.dat file
use Cwd 'chdir';
use File::Basename;
# --- modif to have nimsoft_generic.pm is same directory as perl source (like a use lib is executed before the real execution we need to use BEGIN)
BEGIN
{
use Cwd;
chdir dirname $0;
$curdir=cwd();
}

use lib "$curdir";
use nimsoft_generic;

# --- variables ---
$SERVER=$ENV{"computername"};
$myos=$Config{osname};
@hostnsplit=split(/\\/, $SERVER);
$hostnn=$#hostnsplit;
$hostn = $hostnsplit[$hostnn];
# $hostn=uc($hostn);
$domatch=0;
$version="1.9.4";
$pgm="Nimsoft_Alarm_Reporter (version: $version)";
$colspan1=4;
$debug="n";
# - sql driver (tls 1.2: "ODBC Driver 17 for SQL Server", normal: "SQL Server")
#$sql_driver="ODBC Driver 17 for SQL Server";
# - crypt key
$crkey       = "ab_def%hij&5";

# --- get parameters
&getparams();

# --- user defined columns? 
if ($report ne '')
  {
# --- split the given columns and create table
    &splitreport();
    $colspan=3+$repcount;
    if ($verbose eq 'Y') {print "Received $repcount report columns: $report\n";}
  }
else
 {
   if ($sqlcollection ne 's')
     {
       $colspan=15;
     }
   else
     {
       $colspan=16;
     }
 }

# --- loop to create report (based on -s parameter)
if ($reportloop eq 'Y')
 {
loop:
  print "-\n";
  &calcdate(); 
  print "$intim Start alarm loop with $reportseconds sleep period\n";
  &base_loop;
  sleep($reportseconds);
  goto loop;
 }
else
 {
  print "-\n"; 
  &calcdate(); 
  &base_loop;
 }
print "\n";
if ($do_csv eq 'n')
 {
   print "Created report: $outfil.html\n";
 }
else
 {
   print "Created report: $outfil.csv\n";
 }
print "\n";
exit $msg_selected;

# --- procedures

sub base_loop
{
# --- loop variables
$sqlcount=1;
$teller=0;
$tel=0;
$totcol=0;
$o_nimid="";

# --- reset counters for the next loop
$msg_selected=0;
$msg_excluded=0;
$msg_read=0;

# -- counters
$c_msg=0;
$c_msg_probe=0;
$c_origin=0;
$c_acknowledged=0;
$c_assigned_to=0;
$c_assigned_by=0;
$c_hub=0;
$c_robot=0;
$c_source=0;
$c_probe=0;
$c_sid=0;
%cacknowledged=();
%cassigned_by=();
%cassigned_to=();
%cmsg=();
%cmsg_probe=();
%corigin=();
%chub=();
%crobot=();
%csource=();
%csourceprobe=();
%cprobe=();
%csid=();

# -- if gui is selected start it
if ($gui eq 'y')
{
  tkgui();
}

# --- connect DB ---
if ($sqltype eq 'my')
  {
    $dbh = DBI->connect("DBI:mysql:database=$database;host=$sqlserver", "$sqluser", "$sqlpass") || die "Could not connect to nimsoftslm database on $sqlserver: $DBI::errstr";
  }
if ($sqltype eq 'ms')
  {
    if ($trusted eq 'yes')
      {
        $dbh = DBI->connect("DBI:ODBC:Driver={$sql_driver};Server=$sql_server;Database=$database;Trusted_Connection=Yes;") or die("\n\nCONNECT ERROR:\n\n$DBI::errstr");
        print "Using mssql trusted connection to connect to: $sqlserver with driver: $sql_driver\n";
      }
    else
      {
         if ($sql_dsn eq '')
          {
            $dbh = DBI->connect("DBI:ODBC:Driver={$sql_driver};Server=$sql_server;Database=$sql_db;UID=$sql_user;PWD=$sql_password") or die("\n\nCONNECT ERROR:\n\n$DBI::errstr");
            print "Connect to server: $sql_server db: $sql_db with userid: $sql_user\n";
          }
         else
          {
# -- with dsn
            $dbh = DBI->connect("DBI:ODBC:$sql_dsn","$sql_user","$sql_password") or die("\n\nCONNECT ERROR:\n\n$DBI::errstr");
            print "Connect to server: $sql_server db: $sql_db with userid: $sql_user and DSN: $sql_dsn\n";
          }
      } 
  }

# --- open HTML
# print "open $outfil\_new.html\n";
open(OUTFILE1, ">$outfil\_new.html") || die("Cannot open file $outfil\_new.html : $!");
binmode OUTFILE1, ":encoding(UTF-8)";
if ($s_overview eq 'Y')
 {
   open(OUTFILE2, ">$outfil\_overview\_new.html") || die("Cannot open file $outfil\_overview\_new : $!");
   binmode OUTFILE2, ":encoding(UTF-8)";
 }
# --- open CSV
if ($do_csv eq 'y')
 {
   open(OUTCSV, ">$outfil.csv") || die("Cannot open file $outfil.csv : $!");
#   print OUTCSV "sqlcount,time,duration,robot,source,subsys,prid,severity,level,visible,suppcount,message,suppkey,sid,acknowledged_by,type,assigned_to,assigned_by\n";
   print OUTCSV "sqlcount,time,duration,domain,hub,origin,robot,source,subsys,probe,severity,level,visible,suppcount,message,suppkey,sid,acknowledged_by,type,assigned_by,assigned_to,user_tag1,user_tag2,custom1,custom2,custom3,custom4,custom5\n";
 } 

# --- prepare HTML 
&report_header1;
&report_header2;

# --- do we need an sql like statement? (we need 2 statements, for the old query and the new one)
&sql_filters;

# --- sql query

# for mysql
if ($sqltype eq 'my')
  {
    if ($sql_filt ne 'y')
      {
        $sql = "select l.time, l.type, l.nimid, l.level, l.severity, l.message, l.subsys, l.source, l.hostname, l.prid, l.robot, l.hub, l.nas, l.domain, l.suppcount
, a.suppcount, l.visible, a.supp_key, l.nimts, TIME_FORMAT(TIMEDIFF(l.nimts,l.time), '%H:%i:%s') as duration, l.sid, l.acknowledged_by, l.user_tag1, l.user_tag2,l.origin,l.assigned_by,l.assigned_to from nas_transaction_log l 
left outer join nas_alarms a on l.nimid = a.nimid
 WHERE l.time >= (NOW() - interval $s_days HOUR ) order by l.time, l.suppcount";
      }
    else
      {
        $sql = "select l.time, l.type, l.nimid, l.level, l.severity, l.message, l.subsys, l.source, l.hostname, l.prid, l.robot, l.hub, l.nas, l.domain, l.suppcount
, a.suppcount, l.visible, a.supp_key, l.nimts, TIME_FORMAT(TIMEDIFF(l.nimts,l.time), '%H:%i:%s') as duration, l.sid, l.acknowledged_by, l.user_tag1, l.user_tag2,l.origin,l.assigned_by,l.assigned_to from nas_transaction_log l 
left outer join nas_alarms a on l.nimid = a.nimid
 WHERE l.time >= (NOW() - interval $s_days HOUR ) and $sql_f2 order by l.time, l.suppcount";
      }   
  } 
# for mssql (only date select portion is different)
if ($s_days ne '')
 {
   $sql_days="l.time >= DATEADD(hh, -$s_days, GETDATE())";
   $sql_days1="Created >= DATEADD(hh, -$s_days, GETDATE())";
   if ($sql_endhh eq '' && $sql_starthh eq '')
    {
      print (color("black on_green "),"Reporting period: last $s_days hours",color("reset"),"\n");
    }
   if ($sql_starthh eq '' && $sql_endhh ne '')
    {
      print (color("black on_green "),"Reporting period: start $s_days hours back endtime filter: $sql_endhh hhmm",color("reset"),"\n");
    }
   if ($sql_endhh eq '' && $sq_starthh ne '')
    {
      print (color("black on_green "),"Reporting period: start $s_days hours back starttime filter: $sql_starthh hhmm",color("reset"),"\n");
    }
   if ($sql_endhh ne '' && $sql_starthh ne '')
    {
      print (color("black on_green "),"Reporting period: start $s_days hours back starttime filter: $sql_starthh endtime: $sql_endhh",color("reset"),"\n");
    }
 }
else
 {
   if ($s_minutes ne '')
    {
      $sql_days="l.time >= DATEADD(minute, -$s_minutes, GETDATE())";
      $sql_days1="Created >= DATEADD(minute, -$s_minutes, GETDATE())";
      if ($sql_endhh eq '')
       {
         print (color("black on_green "),"Reporting period: last $s_minutes minutes",color("reset"),"\n");
       }
      else
       {
         print (color("black on_green "),"Reporting period: start $s_minutes minutes back until end time $sql_endhh hhmm",color("reset"),"\n");
       }
    }
   else
    {  
      if ($sql_start ne '' && $sql_end ne '')
       {
         $sql_days="l.time >= '$sql_start' and l.time <= '$sql_end'";
         $sql_days1="Created >= '$sql_start' and Created <= '$sql_end'";
         print (color("black on_green "),"Reporting period: start: $sql_start end: $sql_end",color("reset"),"\n");
       }
      else
       {
         if ($s_month ne '')
          {
            $sql_days="DatePart(\"m\", l.time) = DatePart(\"m\", DateAdd(\"m\", -$s_month, getdate())) AND DatePart(\"yyyy\", l.time) = DatePart(\"yyyy\", DateAdd(\"m\", -$s_month, getdate()))";
            $sql_days1="DatePart(\"m\", created) = DatePart(\"m\", DateAdd(\"m\", -$s_month, getdate())) AND DatePart(\"yyyy\", created) = DatePart(\"yyyy\", DateAdd(\"m\", -$s_month, getdate()))";
            print (color("black on_green "),"Reporting period: last: $s_month month(s) (0 is current month)",color("reset"),"\n");
          }
         else
          {  
            if ($sql_starthh ne '' || $sql_endhh ne '')
             {
               if ($sql_endhh eq '')
                 {
                   $sql_start="$sql_today $sql_starthh";
                   $sql_days="l.time >= '$sql_start'";
                   $sql_days1="Created >= '$sql_start'";
                   print (color("black on_green "),"Reporting period: start: $sql_start end: now",color("reset"),"\n");
                 }
               if ($sql_starthh eq '')
                 {
                   $sql_start="$sql_today 00:00";
                   $sql_end="$sql_today $sql_endhh";  
                   $sql_days="l.time >= '$sql_start' and l.time <= '$sql_end'";
                   $sql_days1="Created >= '$sql_start' and Created <= '$sql_end'";
                   print (color("black on_green "),"Reporting period: start: $sql_start end: $sql_end",color("reset"),"\n");
                 }
              if ($sql_starthh ne '' && $sql_endhh ne '')
                 {
                   $sql_start="$sql_today $sql_starthh";
                   $sql_end="$sql_today $sql_endhh";  
                   $sql_days="l.time >= '$sql_start' and l.time <= '$sql_end'";
                   $sql_days1="Created >= '$sql_start' and Created <= '$sql_end'";
                   print (color("black on_green "),"Reporting period: start: $sql_start end: $sql_end",color("reset"),"\n");
                 }
              }
            else
              { 
                 $sql_days="l.time >= '$sql_start'";
                 $sql_days1="Created >= '$sql_start'";
                 print (color("black on_green "),"Reporting period: start: $sql_start end: now",color("reset"),"\n");
              }
          } 
       }
    }
 }
if ($sqltype eq 'ms')
  {
# - nas_transaction_log
    if ($sqlcollection eq 't')
     {
       if ($sql_filt ne 'y')
         {
           $sql = "select l.time, l.type, l.nimid, l.level, l.severity, l.message, l.subsys, l.source, l.hostname, l.prid, l.robot, l.hub, l.nas, 
l.domain, l.suppcount, a.suppcount, l.visible, a.supp_key, l.nimts, DATEDIFF(DAY, l.nimts,l.time) durationdd, CONVERT(varchar(8), DATEADD(minute, DATEDIFF(minute, l.nimts,l.time), 0), 114) durationhh, l.sid,l.acknowledged_by, l.user_tag1, l.user_tag2,l.origin,l.assigned_by,l.assigned_to 
from nas_transaction_log l with(nolock) 
left outer join nas_alarms a with(nolock) on l.nimid = a.nimid 
WHERE $sql_days
 order by l.time, l.suppcount";
         }
       else
         {
           $sql = "select l.time, l.type, l.nimid, l.level, l.severity, l.message, l.subsys, l.source, l.hostname, l.prid, l.robot, l.hub, l.nas, 
l.domain, l.suppcount, a.suppcount, l.visible, a.supp_key, l.nimts, DATEDIFF(DAY, l.nimts,l.time) durationdd, CONVERT(varchar(8), DATEADD(minute, DATEDIFF(minute, l.nimts,l.time), 0), 114) durationhh, l.sid,l.acknowledged_by, l.user_tag1, l.user_tag2,l.origin,l.assigned_by,l.assigned_to 
from nas_transaction_log l with(nolock) 
left outer join nas_alarms a with(nolock) on l.nimid = a.nimid 
WHERE $sql_days and $sql_f2
 order by l.time, l.suppcount";
         }
     }
# - nas_transaction_summary
    if ($sqlcollection eq 's')
     {
        $sql = "select l.time, l.nimid, l.level, l.severity, l.message, l.subsys, l.source, l.hostname, l.prid, l.robot, l.hub, l.nas, 
l.domain, l.suppcount, l.visible, l.supp_key, l.nimts, l.sid,l.acknowledged_by, 
l.user_tag1, l.user_tag2,l.origin,l.assigned_by,l.assigned_to,l.custom_1,l.custom_2,l.custom_3,l.custom_4,l.custom_5 
from nas_transaction_summary l with(nolock) 
WHERE $sql_days 
 order by l.time";
     }
# - alarmtransactionlog (multi nas)
    if ($sqlcollection eq 'y')
     {
       $sql = "select processed ,Typeid, AlarmId, AlarmSeverity, AlarmSeverityDesc, AlarmMessage, AlarmSubsystem, hostname, source, created, DATEDIFF(DAY, created,processed) durationdd, CONVERT(varchar(8), DATEADD(minute, DATEDIFF(minute, created,processed), 0), 114) durationhh, AlarmSid, Domain, Hub, Robot, Probe, origin from AlarmTransactionLog with(nolock) WHERE $sql_days1 group by alarmid,Created, Processed, TypeId, AlarmSeverity, AlarmSeverityDesc, AlarmSid, AlarmSubsystem, AlarmMessage, hostname, source, Domain, Origin, Hub, Robot, Probe order by processed,created";
     }
  }

# - if debug print used sql statement
 if ($debug eq 'y') 
   { 
     print "sql_start: $sql_start sql_starthh: $sql_starthh sql_end: $sql_end sql_endhh: $sqlendhh s_days: $s_days s_minutes: $s_minutes s_month: $s_month\n";
     print "sql: $sql\n"; 
   }

$sth = $dbh->prepare($sql);
$sth->execute or die "SQL Error: $DBI::errstr\n";
while (@row = $sth->fetchrow_array) 
 {
# - decode line
  $msg_read++;
# - mysql
if ($sqltype eq 'my')
  {
    ($a_time, $a_type, $a_nimid, $a_level, $a_severity, $a_message, $a_subsys, $a_source, $a_hostname, $a_prid, $a_robot, $a_hub, $a_nas, $a_domain, $l_suppcount, $a_suppcount, $a_visible, $a_suppkey, $a_nimtime, $a_duration, $a_sid, $a_acknowledged_by, $a_user_tag1, $a_user_tag2, $a_origin, $a_assigned_by, $a_assigned_to) = @row;
# --- romove possible - sign from $a_duration
    $a_duration=~ s/^-//g;
  }
else
  {
# - mssql
    if ($sqlcollection eq 't')
     {
      ($a_time, $a_type, $a_nimid, $a_level, $a_severity, $a_message, $a_subsys, $a_source, $a_hostname, $a_prid, $a_robot, $a_hub, $a_nas, $a_domain, $l_suppcount, $a_suppcount, $a_visible, $a_suppkey, $a_nimtime, $a_durationdd,$a_durationhh, $a_sid, $a_acknowledged, $a_user_tag1, $a_user_tag2, $a_origin, $a_assigned_by, $a_assigned_to) = @row;
      $a_duration="$a_durationdd.$a_durationhh";
      $a_duration=~ s/\./d/g;
#      $a_duration=~ s/0d/   /g;
      $a_duration=~ s/0d//g;
     }
    if ($sqlcollection eq 's')
     {
      $a_prid="-";
      $a_robot="-";
      $a_origin="-";
      $a_hub="-";
      $a_nas="-";
      $a_domain="-";
      $l_suppcount="-";
      $a_suppcount="-";
      $a_visible="-";
      $a_suppkey="-";
      ($a_time, $a_nimid, $a_level, $a_severity, $a_message, $a_subsys, $a_source, $a_hostname, $a_prid, $a_robot, $a_hub, $a_nas, $a_domain, $a_suppcount, $a_visible, $a_suppkey,$a_nimtime, $a_sid, $a_acknowledged, $a_user_tag1, $a_user_tag2, $a_origin, $a_assigned_by, $a_assigned_to,$a_custom1,$a_custom2,$a_custom3,$a_custom4,$a_custom5) = @row;
      $a_duration="$a_durationdd.$a_durationhh";
      $a_duration=~ s/\./d/g;
#      $a_duration=~ s/0d/   /g;
      $a_duration=~ s/0d//g;
# --- remove left and right blanks
      $a_origin =~s/^\s+|\s+$//g;
      $a_subsys =~s/^\s+|\s+$//g;
      $a_source =~s/^\s+|\s+$//g;
      $a_hostname =~s/^\s+|\s+$//g;
      $a_domain =~s/^\s+|\s+$//g;
      $a_hub =~s/^\s+|\s+$//g;
      $a_robot =~s/^\s+|\s+$//g;
      $a_prid =~s/^\s+|\s+$//g;
     }
    if ($sqlcollection eq 'y')
     {
      $a_prid="-";
      $a_robot="-";
      $a_origin="-";
      $a_hub="-";
      $a_nas="-";
      $a_domain="-";
      $l_suppcount="-";
      $a_suppcount="-";
      $a_visible="-";
      $a_suppkey="-";
      ($a_time, $a_type, $a_nimid, $a_level, $a_severity, $a_message, $a_subsys, $a_source, $a_hostname, $a_nimtime, $a_durationdd,$a_durationhh, $a_sid, $a_domain, $a_hub, $a_robot, $a_prid,$a_origin) = @row;
      $a_duration="$a_durationdd.$a_durationhh";
      $a_duration=~ s/\./d/g;
#      $a_duration=~ s/0d/   /g;
      $a_duration=~ s/0d//g;
# --- remove left and right blanks
      $a_origin =~s/^\s+|\s+$//g;
      $a_subsys =~s/^\s+|\s+$//g;
      $a_source =~s/^\s+|\s+$//g;
      $a_hostname =~s/^\s+|\s+$//g;
      $a_domain =~s/^\s+|\s+$//g;
      $a_hub =~s/^\s+|\s+$//g;
      $a_robot =~s/^\s+|\s+$//g;
      $a_prid =~s/^\s+|\s+$//g;
     }
  }
# --- strip the milliseconds that are not used
  $a_time=substr($a_time,0,19);
  $a_nimtime=substr($a_nimtime,0,19);
  $domatch=0;

# --- create date and time field for overview report
  $a_timdd=substr($a_time,0,10); 
  $a_timhh=substr($a_time,0,13); 
# --- create hh:mm for -2 and -4 time limit
  $a_hhmm=substr($a_time,11,5);

# --- remove message new line characters
  $a_message =~ s/[\x0A\x0D]//g;
  $a_message =~ s///g;
#  $a_message =~ /\Q$a_message\E/;

# --- remove fqdn
  if ($fqdn_strip eq 'y')
   {
     $a_hostname =~ /([^.]+)/;
     $a_hostname=$1;
   }

# --- optional search criteria
     if ($s_message_exc ne '' && $a_message =~ /$s_message_exc/i)
      {
        $msg_excluded++;
        goto bypass;
      }
     if ($s_hostname_exc ne '' && $a_hostname =~ /$s_hostname_exc/i)
      {
        $msg_excluded++;
        goto bypass;
      }
     if ($s_origin_exc ne '' && $a_origin =~ /$s_origin_exc/i)
      {
        $msg_excluded++;
        goto bypass;
      }
     if ($s_hub_exc ne '' && $a_hub =~ /$s_hub_exc/i)
      {
        $msg_excluded++;
        goto bypass;
      }
     if ($s_assignedby_exc ne '' && $a_assigned_by =~ /$s_assignedby_exc/i)
      {
        $msg_excluded++;
        goto bypass;
      }
     if ($s_acknowledged_exc ne '' && $a_acknowledged =~ /$s_acknowledged_exc/i)
      {
        $msg_excluded++;
        goto bypass;
      }
     if ($s_level_exc ne '' && $a_level =~ /$s_level_exc/i)
      {
        $msg_excluded++;
        goto bypass;
      }
     if ($s_assignedto_exc ne '' && $a_assigned_to =~ /$s_assignedto_exc/i)
      {
        $msg_excluded++;
        goto bypass;
      }
     if ($s_robot_exc ne '' && $a_robot =~ /$s_robot_exc/i)
      {
        $msg_excluded++;
        goto bypass;
      }
     if ($s_probe_exc ne '' && $a_prid =~ /$s_probe_exc/i)
      {
        $msg_excluded++;
        goto bypass;
      }

     if ($s_message_inc ne '')
      { 
         if ($a_message =~ /$s_message_inc/i ) 
           {
             $domatch=1;
           }
         else
           {
             $msg_excluded++;
             if ($debug eq 'y') { print "Excluded $msg_excluded in message: $a_message\n"; }
             goto bypass;
           }
      }
     if ($s_count_value ne '')
      {
        if ($s_count_value eq '1')
         { 
           if (($l_suppcount eq '1') || ($l_suppcount eq '') || ($l_suppcount eq '0') || $a_suppcount eq '1' || $a_suppcount eq '0')
              {
                $domatch=1;
              }
           else
              {
                $msg_excluded++;
                if ($debug eq 'y') { print "Excluded $msg_excluded in suppresscount: $a_suppcount\n"; }
                goto bypass;
              }
         }
        if ($s_count_value ne '1')
         { 
           if ($a_suppcount =~ /$s_count_value/i ) 
              {
                $domatch=1;
              }
            else
              {
                $msg_excluded++;
                if ($debug eq 'y') { print "Excluded $msg_excluded in suppresscount: $a_suppcount\n"; }
                goto bypass;
              }
         }
      }
     if ($sql_starthh ne '')
      {
        if ($a_hhmm ge $sql_starthh) 
           {
            $domatch=1;
           }
         else
           {
             $msg_excluded++;
             if ($debug eq 'y') { print "Excluded $msg_excluded in starthh: $sql_starthh\n"; }
             goto bypass;
           }
      }
     if ($sql_endhh ne '')
      {
        if ($a_hhmm le $sql_endhh) 
           {
            $domatch=1;
           }
         else
           {
             $msg_excluded++;
             if ($debug eq 'y') { print "Excluded $msg_excluded in endhh: $sql_endhh\n"; }
             goto bypass;
           }
      }
     if ($s_hostname_inc ne '')
      {
        if ($a_hostname =~ /$s_hostname_inc/i ) 
           {
            $domatch=1;
           }
         else
           {
             $msg_excluded++;
             if ($debug eq 'y') { print "Excluded $msg_excluded in hostname: $a_hostname\n"; }
             goto bypass;
           }
      }
     if ($s_origin_inc ne '')
      {
        if ($a_origin =~ /$s_origin_inc/i ) 
           {
             $domatch=1;
           }
         else
           {
             $msg_excluded++;
             if ($debug eq 'y') { print "Excluded $msg_excluded in origin: $a_origin\n"; }
             goto bypass;
           }
      }
     if ($s_hub_inc ne '')
      {
        if ($a_hub =~ /$s_hub_inc/i ) 
           {
             $domatch=1;
           }
         else
           {
             $msg_excluded++;
             if ($debug eq 'y') { print "Excluded $msg_excluded in hub: $a_hub\n"; }
             goto bypass;
           }
      }
     if ($s_assignedby_inc ne '')
      {
        if ($a_assigned_by =~ /$s_assignedby_inc/i ) 
           {
             $domatch=1;
           }
         else
           {
             $msg_excluded++;
             if ($debug eq 'y') { print "Excluded $msg_excluded in assigned_by: $a_assigned_by\n"; }
             goto bypass;
           }
      }
     if ($s_acknowledged_inc ne '')
      {
        if ($a_acknowledged =~ /$s_acknowledged_inc/i ) 
           {
             $domatch=1;
           }
         else
           {
             $msg_excluded++;
             if ($debug eq 'y') { print "Excluded $msg_excluded in acknowledged_by: $a_acknowledged\n"; }
             goto bypass;
           }
      }
     if ($s_level_inc ne '')
      {
        if ($a_level =~ /$s_level_inc/i ) 
           {
             $domatch=1;
           }
         else
           {
             $msg_excluded++;
             if ($debug eq 'y') { print "Excluded $msg_excluded in level: $a_level\n"; }
             goto bypass;
           }
      }
     if ($s_assignedto_inc ne '')
      {
        if ($a_assigned_to =~ /$s_assignedto_inc/i ) 
           {
             $domatch=1;
           }
         else
           {
             $msg_excluded++;
             if ($debug eq 'y') { print "Excluded $msg_excluded in assigned_to: $a_assigned_to\n"; }
             goto bypass;
           }
      }
     if ($s_robot_inc ne '')
      {
        if ($a_robot =~ /$s_robot_inc/i ) 
           {
             $domatch=1;
           }
         else
           {
             $msg_excluded++;
             if ($debug eq 'y') { print "Excluded $msg_excluded in robot: $a_robot\n"; }
             goto bypass;
           }
      }
     if ($s_probe_inc ne '')
      {
#         print "filter: $s_probe_inc probe: $a_prid\n";
         if ($a_prid =~ /$s_probe_inc/i )
           {
             $domatch=1;
           }
         else
           {
             $msg_excluded++;
             if ($debug eq 'y') { print "Excluded $msg_excluded in probe: $a_prid\n"; }
             goto bypass;
           }
      }
     if ($a_visible eq '0')
      {
         if ($s_visible eq 'n' )
           {
             $msg_excluded++;
             if ($debug eq 'y') { print "Excluded $msg_excluded in visible: 0\n"; }
             goto bypass;
           }
         else
           {
             $domatch=1;
           }
      }
     if ($a_visible eq '1')
      {
         if ($s_visible eq 'o')
           {
             $msg_excluded++;
             if ($debug eq 'y') { print "Excluded $msg_excluded in visible: 1\n"; }
             goto bypass;
           }
         else
           {
             $domatch=1;
           }
      }


# --- set $d_suppcount as combination of suppresscount and maximum if available
   $d_suppcount="";
   if ($l_suppcount ne '' && $l_suppcount ne '0')
    {
      $d_suppcount=$l_suppcount;
    }
   if ($a_suppcount ne '' && $d_suppcount ne '')
    {
      $d_suppcount="$d_suppcount/$a_suppcount";
    }  
   if ($a_severity eq 'clear')
    {
      $d_suppcount="";
    }
   if ($sqlcollection eq 's')
    {
      $d_suppcount=$a_suppcount;
    } 
# --- if columns are defined via -co
# field names: a_duration,a_origin,a_hub,a_robot,a_source,a_subsys,a_prid,a_severity,a_level,a_visible,d_suppcount,a_message,a_suppkey,a_sid,a_acknowledged,a_user_tag1,a_user_tag2,a_assigned_by,a_assigned_to
      $duration=$a_duration;
      $robot=$a_robot;
      $origin=$a_origin;
      $hub=$a_hub;
      $source=$a_source;
      $hostname=$a_hostname;
      $subsys=$a_subsys;
      $probe=$a_prid;  
      $severity=$a_severity;
      $level=$a_level;
      $visible=$a_visible;
      $suppcount=$d_suppcount;
      $message=$a_message;
      $suppkey=$a_suppkey;
      $sid=$a_sid;
      $acknowledged=$a_acknowledged;
      $user_tag1=$a_user_tag1;
      $user_tag2=$a_user_tag2;
      $assigned_by=$a_assigned_by;
      $assigned_to=$a_assigned_to;
      $type=$a_type;
      $custom1=$a_custom1;
      $custom2=$a_custom2;
      $custom3=$a_custom3;
      $custom4=$a_custom4;
      $custom5=$a_custom5;
      $repcc=0;
      while ($rep[$repcc] ne '')
        { 
          if ($repcc eq '0') {$w0=$rep[$repcc];}  
          if ($repcc eq '1') {$w1=$rep[$repcc];} 
          if ($repcc eq '2') {$w2=$rep[$repcc];} 
          if ($repcc eq '3') {$w3=$rep[$repcc];} 
          if ($repcc eq '4') {$w4=$rep[$repcc];} 
          if ($repcc eq '5') {$w5=$rep[$repcc];}    
          if ($repcc eq '6') {$w6=$rep[$repcc];} 
          if ($repcc eq '7') {$w7=$rep[$repcc];} 
          if ($repcc eq '8') {$w8=$rep[$repcc];} 
          if ($repcc eq '9') {$w9=$rep[$repcc];} 
          if ($repcc eq '10') {$w10=$rep[$repcc];}  
          if ($repcc eq '11') {$w11=$rep[$repcc];} 
          if ($repcc eq '12') {$w12=$rep[$repcc];} 
          if ($repcc eq '13') {$w13=$rep[$repcc];} 
          if ($repcc eq '14') {$w14=$rep[$repcc];} 
          if ($repcc eq '15') {$w15=$rep[$repcc];}    
          if ($repcc eq '16') {$w16=$rep[$repcc];} 
          if ($repcc eq '17') {$w17=$rep[$repcc];} 
          if ($repcc eq '18') {$w18=$rep[$repcc];} 
          if ($repcc eq '19') {$w19=$rep[$repcc];} 
          if ($repcc eq '20') {$w20=$rep[$repcc];}          
          $repcc++;
        }
     if ($domatch eq '1')
      {
        if ($numlines > $msg_selected)
         {
            &detail_line();
         }
        else
         {
            $endmax=1;
         }
        &do_overview;
        $msg_selected++;
      }  
     else
      {  
#        $msg_excluded++;
      }   
# --- in case no filters are defined, print all
     if ($s_message_inc eq '' && $s_message_exc eq '' && $s_hostname_inc eq '' && $s_hostname_exc eq '' && $s_origin_inc eq '' && $s_origin_exc eq ''  && $s_hub_inc eq '' && $s_hub_exc eq '' && $s_assignedby_inc eq '' && $s_assignedby_exc eq '' && $s_acknowledged_inc eq '' && $s_acknowledged_exc eq '' && $s_level_inc eq '' && $s_level_exc eq '' && $s_assignedto_inc eq '' && $s_assignedto_exc eq '' && $s_robot_inc eq '' && $s_robot_exc eq '' && $s_probe_inc eq '' && $s_probe_exc eq '' && $s_visible eq 'y' && $s_starthh eq '' && $s_endhh eq '' && $domatch ne '1')
      {
        &detail_line();
        &do_overview;
        $msg_selected++;
      } 
bypass:
 } 

# --- disconnect ---
$dbh->disconnect();

# --- fill up the page
while ($teller < 45)
 {
  &report_emptyline;
  $teller = $teller + 1;
 }

# --- footers HTML
&report_end;

# --- close HTML
close (OUTFILE1);

# --- close CSV
if ($do_csv eq 'y')
 {
   close (OUTCSV);
 }

# --- rename reports (specially done for auto-update mode)
# system("erase /Q $outfil.html > NUL 2>&1");
# system("rename $outfil\_new.html $outfile.html");
unlink "$outfil.html";
move("$outfil\_new.html","$outfil.html");
print "Messages read: $msg_read  Messages selected: $msg_selected Messages excluded: $msg_excluded\n";

# --- print hub overview
if ($s_overview eq 'N')
 {
   goto overviewend;
 }

# --- count lines to decide on extra empty lines
$ocount=0;

# --- prepare HTML 
&report_over_header1;

# --- print date overview
$pcount=1;
$overtit="Date";
&report_over_header2;
foreach $dateddentry (sort keys(%cdatedd))
 {
   $over1=$dateddentry;
   $over2=$cdatedd{$dateddentry};
   $over0="Date $pcount";
   &report_over_detail(); 
#   print "Hub $pcount: $over1 #msg: $over2\n";
   $pcount++;
   $ocount++;
 }
&report_emptyline_over;

# --- print hour overview
$pcount=1;
$overtit="Hour";
&report_over_header2;
foreach $dathehentry (sort keys(%cdatheh))
 {
   $over1=$dathehentry;
   $over2=$cdatheh{$dathehentry};
   $over0="Hour $pcount";
   &report_over_detail(); 
#   print "Hub $pcount: $over1 #msg: $over2\n";
   $pcount++;
   $ocount++;
 }
&report_emptyline_over;

# --- print origin overview
$pcount=1;
$overtit="Origin";
&report_over_header2;
foreach $originentry (sort sortorigin keys(%corigin))
 {
   $over1=$originentry;
   $over2=$corigin{$originentry};
   $over0="Origin $pcount";
   &report_over_detail(); 
#   print "Origin $pcount: $over1 #msg: $over2\n";
   $pcount++;
   $ocount++;
 }
&report_emptyline_over;


# --- print hub overview
$pcount=1;
$overtit="Hub";
&report_over_header2;
foreach $hubentry (sort sorthub keys(%chub))
 {
   $over1=$hubentry;
   $over2=$chub{$hubentry};
   $over0="Hub $pcount";
   &report_over_detail(); 
#   print "Hub $pcount: $over1 #msg: $over2\n";
   $pcount++;
   $ocount++;
 }
&report_emptyline_over;

# --- print robot overview
$pcount=1;
$overtit="Robot";
&report_over_header2;
foreach $robotentry (sort sortrobot keys(%crobot))
 {
   $over1=$robotentry;
   $over2=$crobot{$robotentry};
   $over0="Robot $pcount";
   &report_over_detail(); 
#   print"Robot $pcount: $over1 #msg: $over2\n";
   $pcount++;
   $ocount++;
 }
&report_emptyline_over;
# --- print source overview
$pcount=1;
$overtit="Source";
&report_over_header2;
foreach $sourceentry (sort sortsource keys(%csource))
 {
   $over1=$sourceentry;
   $over2=$csource{$sourceentry};
   $over0="Source $pcount";
   &report_over_detail(); 
#   print"Source $pcount: $over1 #msg: $over2\n";
   $pcount++;
   $ocount++;
 }
&report_emptyline_over;

# --- print source+probe overview
$pcount=1;
$overtit="Source&Probe";
&report_over_header2;
foreach $sourceentry (sort sortsourceprobe keys(%csourceprobe))
 {
   $over1=$sourceentry;
   $over2=$csourceprobe{$sourceentry};
   $over0="Source $pcount";
   &report_over_detail(); 
#   print"Source $pcount: $over1 #msg: $over2\n";
   $pcount++;
   $ocount++;
 }
&report_emptyline_over;

# --- print probe overview
$pcount=1;
$overtit="Probe";
&report_over_header2;
foreach $probeentry (sort sortprobe keys(%cprobe))
 {
   $over1=$probeentry;
   $over2=$cprobe{$probeentry};
   $over0="Probe $pcount";
   &report_over_detail(); 
#   print"Probe $pcount: $over1 #msg: $over2\n";
   $pcount++;
   $ocount++;
 }
&report_emptyline_over;

# --- print sid overview
$pcount=1;
$overtit="Sid";
&report_over_header2;
foreach $sidentry (sort sortsid keys(%csid))
 {
   $over1=$sidentry;
   $over2=$csid{$sidentry};
   $over0="Sid $pcount";
   &report_over_detail(); 
#   print"Sid $pcount: $over1 #msg: $over2\n";
   $pcount++;
   $ocount++;
 }
&report_emptyline_over;

# --- print acknowledged_by overview
$pcount=1;
$overtit="acknowledged_by";
&report_over_header2;
foreach $ackentry (sort sortacknowledged keys(%cacknowledged))
 {
   $over1=$ackentry;
   $over2=$cacknowledged{$ackentry};
   $over0="ack $pcount";
   &report_over_detail(); 
#   print"Sid $pcount: $over1 #msg: $over2\n";
   $pcount++;
   $ocount++;
 }
&report_emptyline_over;

# --- print assigned_by overview
$pcount=1;
$overtit="assigned_by";
&report_over_header2;
foreach $asbyentry (sort sortassigned_by keys(%cassigned_by))
 {
   $over1=$asbyentry;
   $over2=$cassigned_by{$asbyentry};
   $over0="Assigned By $pcount";
   &report_over_detail(); 
#   print"Msg $pcount: $over1 #msg: $over2\n";
   $pcount++;
   $ocount++;
 }

&report_emptyline_over;

# --- print assigned_to overview
$pcount=1;
$overtit="assigned_to";
&report_over_header2;
foreach $astoentry (sort sortassigned_to keys(%cassigned_to))
 {
   $over1=$astoentry;
   $over2=$cassigned_to{$astoentry};
   $over0="Assigned To $pcount";
   &report_over_detail(); 
#   print"Msg $pcount: $over1 #msg: $over2\n";
   $pcount++;
   $ocount++;
 }

&report_emptyline_over;

# --- print msg_probe overview
$pcount=1;
$overtit="Msg and probe";
&report_over_header2;
foreach $msgprobeentry (sort sortmsgprobe keys(%cmsg_probe))
 {
   $over1=$msgprobeentry;
   $over2=$cmsg_probe{$msgprobeentry};
   $over0="Msg $pcount";
   &report_over_detail(); 
#   print"Msg $pcount: $over1 #msg: $over2\n";
   $pcount++;
   $ocount++;
 }

&report_emptyline_over;

# --- print msg overview
$pcount=1;
$overtit="Msg";
&report_over_header2;
foreach $msgentry (sort sortmsg keys(%cmsg))
 {
   $over1=$msgentry;
   $over2=$cmsg{$msgentry};
   $over0="Msg $pcount";
   &report_over_detail(); 
#   print"Msg $pcount: $over1 #msg: $over2\n";
   $pcount++;
   $ocount++;
 }

&report_emptyline_over;

while ($ocount < 30)
 {
   &report_emptyline_over;
   $ocount++;
 }

&report_end_over;
close (OUTFILE2);
# --- rename reports (specially done for auto-update mode)
# system("erase /Q $outfil\_overview.html > NUL 2>&1");
# system("rename $outfil\_overview\_new.html $outfile\_overview.html");
unlink "$outfil\_overview.html";
move("$outfil\_overview\_new.html","$outfil\_overview.html");

overviewend:

# -- direct view of report?
if ($repview eq 'y' && $do_csv eq 'n')
{
  print "Start IEXPLORE via System call\n";
  $exe_path1 = which 'iexplore.exe';

  if ($exe_path1 ne '')
   {
     system("\"$exe_path1\" \"$outdir\\$outfile.html\"");
   }
  else
   {
     system("\"C:\\Program Files\\Internet Explorer\\iexplore.exe\" \"$outdir\\$outfile.html\"");
   }
  }
}

sub do_overview
{
# -- create overview counters
  if ($s_overview eq 'Y')
   {
     $b_message=$a_message;
# - do not use the digits in overview counts
     $b_message=~ s/[0-9]+/xx/g;
     $cmsg{"$b_message"} += 1;
     if ($a_prid ne '')
      {
        $cmsg_probe{"$a_prid: $b_message"} += 1;
      }
     else
      {
        $cmsg_probe{"N\/A: $b_message"} += 1;
      } 
     $corigin{"$a_origin"} += 1;
     $chub{"$a_hub"} += 1;
     $crobot{"$a_robot"} += 1;
     $csource{"$a_source"} += 1;
     $csourceprobe{"$a_source\_$a_prid"} += 1;
     $cprobe{"$a_prid"} += 1;
     $csid{"$a_sid"} += 1;
     $cdatedd{"$a_timdd"} += 1;
     $cdatheh{"$a_timhh"} += 1;
     $cacknowledged{"$a_acknowledged"} += 1;
     $cassigned_by{"$a_assigned_by"} += 1;
     $cassigned_to{"$a_assigned_to"} += 1;
   }
}

sub sortorigin
{
  $corigin{$a} <=> $corigin{$b};
}

sub sorthub
{
  $chub{$a} <=> $chub{$b};
}

sub sortrobot
{
  $crobot{$a} <=> $crobot{$b};
}

sub sortsource
{
  $csource{$a} <=> $csource{$b};
}

sub sortsourceprobe
{
  $csourceprobe{$a} <=> $csourceprobe{$b};
}

sub sortprobe
{
  $cprobe{$a} <=> $cprobe{$b};
}

sub sortmsg
{
  $cmsg{$a} <=> $cmsg{$b};
}

sub sortacknowledged
{
  $cacknowledged{$a} <=> $cacknowledged{$b};
}

sub sortassigned_by
{
  $cassigned_by{$a} <=> $cassigned_by{$b};
}

sub sortassigned_to
{
  $cassigned_to{$a} <=> $cassigned_to{$b};
}

sub sortmsgprobe
{
  $cmsg_probe{$a} <=> $cmsg_probe{$b};
}

sub sortsid
{
  $csid{$a} <=> $csid{$b};
}

sub sortdatedd
{
  $cdatedd{$a} <=> $cdatedd{$b};
}

sub sortdatheh
{
  $cdatheh{$a} <=> $cdatheh{$b};
}


sub report_header1
 {
  print OUTFILE1 "<html>\n";
  print OUTFILE1 "<meta http-equiv=\"Content-Language\" content=\"en-us\">\n";
  print OUTFILE1 "<meta http-equiv=\"Content-Type\" content=\"text/html; charset=windows-1252\">\n";
  if ($reportloop eq 'Y')
   {
     print OUTFILE1 "<meta http-equiv=\"REFRESH\" CONTENT=\"$reportseconds\" content=\"no-cache\">\n"; 
   }
  print OUTFILE1 "<title>Nimsoft Alarm Reporter</title>\n";
  print OUTFILE1 "</head>\n";
  print OUTFILE1 "<body>\n";
#  This section is used a header to indicate the type of report and when it was generated
  print OUTFILE1 "<table border=\"0\" cellpadding=\"0\" cellspacing=\"0\" style=\"border-collapse: collapse\" bordercolor=\"#111111\" width=\"100%\" id=\"AutoNumber1\">\n";
# Put references to other reports
  print OUTFILE1 "<tr>\n";
  $colspana=$colspan+1;
  print OUTFILE1 "<td width=\"100%\" bgcolor=\"#009AD6\" colspan=\"$colspana\" align=\"left\" height=\"40\"><b>\n";

# --- set $s_filter
  $s_filter="Filters: last $s_days hours Visible: $s_visible ";
  if ($s_message_inc ne '')
   {
     $s_filter="$s_filter msg incl.: $s_message_inc";
   }
  if ($s_message_exc ne '')
   {
     $s_filter="$s_filter msg excl.: $s_message_exc";
   }
  if ($s_hostname_inc ne '' || $s_hostname_exc ne '')
   {
     $s_filter="$s_filter hostname: inc: $s_hostname_inc exc: $s_hostname_exc";
   }
  if ($s_origin_inc ne '' || $s_origin_exc ne '')
   {
     $s_filter="$s_filter origin: inc: $s_origin_inc exc: $s_origin_exc";
   }
  if ($s_hub_inc ne '' || $s_hub_exc ne '')
   {
     $s_filter="$s_filter hub: inc: $s_hub_inc exc: $s_hub_exc";
   }
  if ($s_robot_inc ne '' || $s_robot_exc ne '')
   {
     $s_filter="$s_filter robot: inc: $s_robot_inc exc: $s_robot_exc";
   }
  if ($s_probe_inc ne '' || $s_probe_exc ne '')
   {
     $s_filter="$s_filter probe: inc: $s_probe_inc exc: $s_probe_exc";
   }
  if ($s_count_value ne '')
   {
     $s_filter="$s_filter suppcount: $s_count_value";
   }
  if ($sql_starthh ne '')
   {
     $s_filter="$s_filter starttime: $sql_starthh";
   }
  if ($sql_endhh ne '')
   {
     $s_filter="$s_filter endtime: $sql_endhh";
   }
  if ($s_assignedby_inc ne '')
   {
     $s_filter="$s_filter assigned_by_inc: $s_assignedby_inc";
   }
  if ($s_assignedby_exc ne '')
   {
     $s_filter="$s_filter assigned_by_exc: $s_assignedby_exc";
   }
  if ($s_assignedto_inc ne '')
   {
     $s_filter="$s_filter assigned_to_inc: $s_assignedto_inc";
   }
  if ($s_assignedto_exc ne '')
   {
     $s_filter="$s_filter assigned_to_exc: $s_assignedto_exc";
   }
  if ($s_acknowledged_inc ne '')
   {
     $s_filter="$s_filter acknowledged_inc: $s_acknowledged_inc";
   }
   if ($s_acknowledged_exc ne '')
   {
     $s_filter="$s_filter acknowledged_exc: $s_acknowledged_exc";
   }
   if ($sql_origin ne '')
   {
     $s_filter="$s_filter sql_origin: $sql_origin";
   }
   if ($sql_hostname ne '')
   {
     $s_filter="$s_filter sql_hostname: $sql_hostname";
   }
   if ($sql_probe ne '')
   {
     $s_filter="$s_filter sql_probe: $sql_probe";
   }
   if ($sql_message ne '')
   {
     $s_filter="$s_filter sql_message: $sql_message";
   }
  if ($sql_robot ne '')
   {
     $s_filter="$s_filter sql_robot: $sql_robot";
   }
  if ($sql_hub ne '')
   {
     $s_filter="$s_filter sql_hub: $sql_hub";
   }
  if ($reportloop eq 'N')
   {
     print OUTFILE1 "<font color=\"#FFFFFF\" face=\"MS Sans Serif\" size=\"2\">&nbsp;Nimsoft $sqltable on $dag $dagnum $maand ($maandnum) $jaar $intim (no autoupdate) $s_filter</font>\n";
   }
  else
   {
    print OUTFILE1 "<font color=\"#FFFFFF\" face=\"MS Sans Serif\" size=\"2\">&nbsp;Nimsoft $sqltable on $dag $dagnum $maand ($maandnum) $jaar $intim (auto update after $reportseconds seconds) $s_filter</font>\n";
   }
  print OUTFILE1 "</b></td>\n";
  print OUTFILE1 "</tr>\n";
 }

sub report_over_header1
 {
  print OUTFILE2 "<html>\n";
  print OUTFILE2 "<meta http-equiv=\"Content-Language\" content=\"en-us\">\n";
  print OUTFILE2 "<meta http-equiv=\"Content-Type\" content=\"text/html; charset=windows-1252\">\n";
  if ($reportloop eq 'Y')
   {
     print OUTFILE2 "<meta http-equiv=\"REFRESH\" CONTENT=\"$reportseconds\" content=\"no-cache\">\n"; 
   }
  print OUTFILE2 "<title>Nimsoft Alarm Reporter Overview</title>\n";
  print OUTFILE2 "</head>\n";
  print OUTFILE2 "<body>\n";
#  This section is used a header to indicate the type of report and when it was generated
  print OUTFILE2 "<table border=\"0\" cellpadding=\"0\" cellspacing=\"0\" style=\"border-collapse: collapse\" bordercolor=\"#111111\" width=\"100%\" id=\"AutoNumber1\">\n";
# Put references to other reports
  print OUTFILE2 "<tr>\n";
  print OUTFILE2 "<td width=\"100%\" bgcolor=\"#009AD6\" colspan=\"$colspan1\" align=\"left\" height=\"30\"><b>\n";

# --- set $s_filter
  if ($s_days ne '') {$s_filter="Filters: last $s_days hours Visible: $s_visible ";}
  if ($s_month ne '') {$s_filter="Filters: last $s_days month Visible: $s_visible ";}
  if ($s_message_inc ne '')
   {
     $s_filter="$s_filter msg incl.: $s_message_inc";
   }
  if ($s_message_exc ne '')
   {
     $s_filter="$s_filter msg excl.: $s_message_exc";
   }
  if ($s_hostname ne '')
   {
     $s_filter="$s_filter hostname: $s_hostname";
   }
  if ($s_origin ne '')
   {
     $s_filter="$s_filter origin: $s_origin";
   }
  if ($s_hub ne '')
   {
     $s_filter="$s_filter hub: $s_hub";
   }
  if ($s_robot ne '')
   {
     $s_filter="$s_filter robot: $s_robot";
   }
  if ($s_probe ne '')
   {
     $s_filter="$s_filter probe: $s_probe";
   }
  if ($s_count_value ne '')
   {
     $s_filter="$s_filter suppcount: $s_count_value";
   }
  if ($sql_starthh ne '')
   {
     $s_filter="$s_filter starttime: $sql_starthh";
   }
  if ($sql_endhh ne '')
   {
     $s_filter="$s_filter endtime: $sql_endhh";
   }
  if ($sql_origin ne '')
   {
     $s_filter="$s_filter sql_origin: $sql_origin";
   }
  if ($sql_hostname ne '')
   {
     $s_filter="$s_filter sql_hostname: $sql_hostname";
   }
  if ($sql_probe ne '')
   {
     $s_filter="$s_filter sql_probe: $sql_probe";
   }
  if ($sql_message ne '')
   {
     $s_filter="$s_filter sql_message: $sql_message";
   }
  if ($sql_robot ne '')
   {
     $s_filter="$s_filter sql_robot: $sql_robot";
   }
  if ($sql_hub ne '')
   {
     $s_filter="$s_filter sql_hub: $sql_hub";
   }
 
  if ($reportloop eq 'N')
   {
     print OUTFILE2 "<font color=\"#FFFFFF\" face=\"MS Sans Serif\" size=\"2\">&nbsp;Nimsoft Overview Report on $dag $dagnum $maand ($maandnum) $jaar $intim (no autoupdate) $s_filter</font>\n";
   }
  else
   {
    print OUTFILE2 "<font color=\"#FFFFFF\" face=\"MS Sans Serif\" size=\"2\">&nbsp;Nimsoft Overview Report on $dag $dagnum $maand ($maandnum) $jaar $intim (auto update after $reportseconds seconds) $s_filter</font>\n";
   }
  print OUTFILE2 "</b></td>\n";
  print OUTFILE2 "</tr>\n";
 }

sub report_emptyline
 {
#   $oldprinttel=$printtel; 
  $telc=0;
  $colspan1=$colspan+1;
  if ($tel eq "0")
   { 
    print OUTFILE1 "<tr>\n";
    while ($telc < $colspan1)
     {
       print OUTFILE1 "<td><font face=\"Tahoma\" size=\"2\" color=\"#FFFFFF\">-</font></td>\n";
       $telc = $telc + 1;
     }
    print OUTFILE1 "</tr>\n";
    $tel = 1;
   }
  else
   {
    print OUTFILE1 "<tr>\n";
    while ($telc < $colspan1)
     {
       print OUTFILE1 "<td bgcolor=\"#D2FFFF\"><pre>\n";
       print OUTFILE1 "<font face=\"Tahoma\" size=\"2\" color=\"#D2FFFF\">-</font></pre></td>\n";
       $telc = $telc + 1;
     }
    print OUTFILE1 "</tr>\n";
    $tel = 0;
   }
 }

sub old_report_emptyline
 {
  if ($tel eq "0")
   { 
    print OUTFILE1 "<tr>\n";
    print OUTFILE1 "<td><font face=\"Tahoma\" size=\"2\" color=\"#FFFFFF\">-</font></td>\n";
    print OUTFILE1 "<td><font face=\"Tahoma\" size=\"2\" color=\"#FFFFFF\">-</font></td>\n";
    print OUTFILE1 "<td><font face=\"Tahoma\" size=\"2\" color=\"#FFFFFF\">-</font></td>\n";
    print OUTFILE1 "<td><font face=\"Tahoma\" size=\"2\" color=\"#FFFFFF\">-</font></td>\n";
    print OUTFILE1 "<td><font face=\"Tahoma\" size=\"2\" color=\"#FFFFFF\">-</font></td>\n";
    print OUTFILE1 "<td><font face=\"Tahoma\" size=\"2\" color=\"#FFFFFF\">-</font></td>\n";
    print OUTFILE1 "<td><font face=\"Tahoma\" size=\"2\" color=\"#FFFFFF\">-</font></td>\n";
    print OUTFILE1 "<td><font face=\"Tahoma\" size=\"2\" color=\"#FFFFFF\">-</font></td>\n";
    print OUTFILE1 "<td><font face=\"Tahoma\" size=\"2\" color=\"#FFFFFF\">-</font></td>\n";
    print OUTFILE1 "<td><font face=\"Tahoma\" size=\"2\" color=\"#FFFFFF\">-</font></td>\n";
    print OUTFILE1 "<td><font face=\"Tahoma\" size=\"2\" color=\"#FFFFFF\">-</font></td>\n";
    print OUTFILE1 "<td><font face=\"Tahoma\" size=\"2\" color=\"#FFFFFF\">-</font></td>\n";
    print OUTFILE1 "<td><font face=\"Tahoma\" size=\"2\" color=\"#FFFFFF\">-</font></td>\n";
    print OUTFILE1 "<td><font face=\"Tahoma\" size=\"2\" color=\"#FFFFFF\">-</font></td>\n";
    print OUTFILE1 "<td><font face=\"Tahoma\" size=\"2\" color=\"#FFFFFF\">-</font></td>\n";
    print OUTFILE1 "</tr>\n";
    $tel = 1;
   }
  else
   {
    print OUTFILE1 "<tr>\n";
    print OUTFILE1 "<td bgcolor=\"#D2FFFF\"><pre>\n";
    print OUTFILE1 "<font face=\"Tahoma\" size=\"2\" color=\"#D2FFFF\">-</font></pre></td>\n";
    print OUTFILE1 "<td bgcolor=\"#D2FFFF\"><pre>\n";
    print OUTFILE1 "<font face=\"Tahoma\" size=\"2\" color=\"#D2FFFF\">-</font></pre></td>\n";
    print OUTFILE1 "<td bgcolor=\"#D2FFFF\"><pre>\n";
    print OUTFILE1 "<font face=\"Tahoma\" size=\"2\" color=\"#D2FFFF\">-</font></pre></td>\n";
    print OUTFILE1 "<td bgcolor=\"#D2FFFF\"><pre>\n";
    print OUTFILE1 "<font face=\"Tahoma\" size=\"2\" color=\"#D2FFFF\">-</font></pre></td>\n";
    print OUTFILE1 "<td bgcolor=\"#D2FFFF\"><pre>\n";
    print OUTFILE1 "<font face=\"Tahoma\" size=\"2\" color=\"#D2FFFF\">-</font></pre></td>\n";
    print OUTFILE1 "<td bgcolor=\"#D2FFFF\"><pre>\n";
    print OUTFILE1 "<font face=\"Tahoma\" size=\"2\" color=\"#D2FFFF\">-</font></pre></td>\n";
    print OUTFILE1 "<td bgcolor=\"#D2FFFF\"><pre>\n";
    print OUTFILE1 "<font face=\"Tahoma\" size=\"2\" color=\"#D2FFFF\">-</font></pre></td>\n";
    print OUTFILE1 "<td bgcolor=\"#D2FFFF\"><pre>\n";
    print OUTFILE1 "<font face=\"Tahoma\" size=\"2\" color=\"#D2FFFF\">-</font></pre></td>\n";
    print OUTFILE1 "<td bgcolor=\"#D2FFFF\"><pre>\n";
    print OUTFILE1 "<font face=\"Tahoma\" size=\"2\" color=\"#D2FFFF\">-</font></pre></td>\n";
    print OUTFILE1 "<td bgcolor=\"#D2FFFF\"><pre>\n";
    print OUTFILE1 "<font face=\"Tahoma\" size=\"2\" color=\"#D2FFFF\">-</font></pre></td>\n";
    print OUTFILE1 "<td bgcolor=\"#D2FFFF\"><pre>\n";
    print OUTFILE1 "<font face=\"Tahoma\" size=\"2\" color=\"#D2FFFF\">-</font></pre></td>\n";
    print OUTFILE1 "<td bgcolor=\"#D2FFFF\"><pre>\n";
    print OUTFILE1 "<font face=\"Tahoma\" size=\"2\" color=\"#D2FFFF\">-</font></pre></td>\n";
    print OUTFILE1 "<td bgcolor=\"#D2FFFF\"><pre>\n";
    print OUTFILE1 "<font face=\"Tahoma\" size=\"2\" color=\"#D2FFFF\">-</font></pre></td>\n";
    print OUTFILE1 "<td bgcolor=\"#D2FFFF\"><pre>\n";
    print OUTFILE1 "<font face=\"Tahoma\" size=\"2\" color=\"#D2FFFF\">-</font></pre></td>\n";
    print OUTFILE1 "<td bgcolor=\"#D2FFFF\"><pre>\n";
    print OUTFILE1 "<font face=\"Tahoma\" size=\"2\" color=\"#D2FFFF\">-</font></pre></td>\n";
    print OUTFILE1 "</tr>\n";
    $tel = 0;
   }
 }

sub report_emptyline_over
 {
  if ($tel eq "0")
   { 
    print OUTFILE2 "<tr>\n";
    print OUTFILE2 "<td><font face=\"Tahoma\" size=\"2\" color=\"#FFFFFF\">-</font></td>\n";
    print OUTFILE2 "<td><font face=\"Tahoma\" size=\"2\" color=\"#FFFFFF\">-</font></td>\n";
    print OUTFILE2 "<td><font face=\"Tahoma\" size=\"2\" color=\"#FFFFFF\">-</font></td>\n";
    print OUTFILE2 "</tr>\n";
    $tel = 1;
   }
  else
   {
    print OUTFILE2 "<tr>\n";
    print OUTFILE2 "<td bgcolor=\"#D2FFFF\"><pre>\n";
    print OUTFILE2 "<font face=\"Tahoma\" size=\"2\" color=\"#D2FFFF\">-</font></pre></td>\n";
    print OUTFILE2 "<td bgcolor=\"#D2FFFF\"><pre>\n";
    print OUTFILE2 "<font face=\"Tahoma\" size=\"2\" color=\"#D2FFFF\">-</font></pre></td>\n";
    print OUTFILE2 "<td bgcolor=\"#D2FFFF\"><pre>\n";
    print OUTFILE2 "<font face=\"Tahoma\" size=\"2\" color=\"#D2FFFF\">-</font></pre></td>\n";
    print OUTFILE2 "</tr>\n";
    $tel = 0;
   }
 }

sub report_end
 {
  &report_emptyline; 
  print OUTFILE1 "</table>\n";
# html
#  print OUTFILE1 "</body>\n";
#  print OUTFILE1 "</html>\n";
if ($endmax eq '1')
 {
   print OUTFILE1 "<h5 align=\"center\" color=\"#FF0000\">&nbsp;STOPPED CREATING REPORT DUE TO MORE THAN $numlines LINES IN THE GENERATED REPORT (this can be modified via the -nb option)</font>\n";
 }
  &report_fieldex;

  print OUTFILE1 "</table>\n";
  print OUTFILE1 "</body>\n";
  print OUTFILE1 "</html>\n";

  print OUTFILE1 "<p align=\"center\"><font face=\"Arial\" size=\"-2\">Copyright $jaar (generated by $pgm)<br>\n";
  print OUTFILE1 "<a href=\"http://www.ca.com\">CA</a></font></p>\n";
 }

sub report_end_over
 {
  &report_fieldex;
  print OUTFILE2 "</table>\n";
  print OUTFILE2 "</body>\n";
  print OUTFILE2 "</html>\n";
  print OUTFILE2 "<p align=\"center\"><font face=\"Arial\" size=\"-2\">Copyright $jaar (generated by $pgm)<br>\n";
  print OUTFILE2 "<a href=\"http://www.ca.com\">CA</a></font></p>\n";
 }

sub report_header2
 {
# print month title
     print OUTFILE1 "<tr>\n";
     print OUTFILE1 "<td bgcolor=\"#6BC3CE\"><font face=\"Arial\" size=\"2\"><b>&nbsp;Count</b></font></td>\n";
     print OUTFILE1 "<td bgcolor=\"#6BC3CE\"><font face=\"Arial\" size=\"2\"><b>&nbsp;Created</b></font></td>\n";
     print OUTFILE1 "<td bgcolor=\"#6BC3CE\"><font face=\"Arial\" size=\"2\"><b>&nbsp;S</b></font></td>\n";
     for (my $i=0; $i <= $repcount; $i++) 
      {
        $tit0=$rep[$i];
        $tit0=~ s/^([a-z])/\u$1/;
#        print OUTFILE1 "<td bgcolor=\"#6BC3CE\"><pre><font face=\"Arial\" size=\"2\"><b>&nbsp;$tit0</b></font></pre></td>\n";
        print OUTFILE1 "<td bgcolor=\"#6BC3CE\"><font face=\"Arial\" size=\"2\"><b>&nbsp;$tit0</b></font></td>\n";
      }

#     print OUTFILE1 "<td bgcolor=\"#6BC3CE\"><font face=\"Arial\" size=\"2\"><b>&nbsp;Duration</b></font></td>\n";
#     print OUTFILE1 "<td bgcolor=\"#6BC3CE\"><font face=\"Arial\" size=\"2\"><b>&nbsp;Robot</b></font></td>\n";
#     print OUTFILE1 "<td bgcolor=\"#6BC3CE\"><font face=\"Arial\" size=\"2\"><b>&nbsp;Source</b></font></td>\n";
#     print OUTFILE1 "<td bgcolor=\"#6BC3CE\"><font face=\"Arial\" size=\"2\"><b>&nbsp;Subsys</b></font></td>\n";
#     print OUTFILE1 "<td bgcolor=\"#6BC3CE\"><font face=\"Arial\" size=\"2\"><b>&nbsp;Probe </b></font></td>\n";
#     print OUTFILE1 "<td bgcolor=\"#6BC3CE\"><font face=\"Arial\" size=\"2\"><b>&nbsp;Severity</b></font></td>\n";
#     print OUTFILE1 "<td bgcolor=\"#6BC3CE\"><font face=\"Arial\" size=\"2\"><b>&nbsp;S</b></font></td>\n";
#     print OUTFILE1 "<td bgcolor=\"#6BC3CE\"><font face=\"Arial\" size=\"2\"><b>&nbsp;V</b></font></td>\n";
#     print OUTFILE1 "<td bgcolor=\"#6BC3CE\"><font face=\"Arial\" size=\"2\"><b>&nbsp;#</b></font></td>\n";
#     print OUTFILE1 "<td bgcolor=\"#6BC3CE\"><font face=\"Arial\" size=\"2\"><b>&nbsp;Message</b></font></td>\n";
#     print OUTFILE1 "<td bgcolor=\"#6BC3CE\"><font face=\"Arial\" size=\"2\"><b>&nbsp;SupKey</b></font></td>\n";
#     print OUTFILE1 "<td bgcolor=\"#6BC3CE\"><font face=\"Arial\" size=\"2\"><b>&nbsp;Sid</b></font></td>\n";
#     print OUTFILE1 "<td bgcolor=\"#6BC3CE\"><font face=\"Arial\" size=\"2\"><b>&nbsp;Ack</b></font></td>\n";
     print OUTFILE1 "</tr>\n";
 }

sub report_over_header2
 {
# print month title
     print OUTFILE2 "<tr>\n";
     print OUTFILE2 "<td bgcolor=\"#6BC3CE\"><font face=\"Arial\" size=\"2\"><b>&nbsp;#</b></font></td>\n";
     print OUTFILE2 "<td bgcolor=\"#6BC3CE\"><font face=\"Arial\" size=\"2\"><b>&nbsp;Count</b></font></td>\n";
     print OUTFILE2 "<td bgcolor=\"#6BC3CE\"><font face=\"Arial\" size=\"2\"><b>&nbsp;$overtit</b></font></td>\n";
     print OUTFILE2 "</tr>\n";
 }

# $a_time, $a_type, $a_nimid, $a_level, $a_severity, $a_message, $a_subsys, $a_source, $a_hostname, $a_prid, $a_robot, $a_hub, $a_nas, $a_domain
sub report_detail
 {
   if ($do_csv eq 'y')
    {
#   print OUTCSV "sqlcount,time,duration,domain,hub,origin,robot,source,subsys,probe,severity,level,visible,suppcount,message,suppkey,sid,acknowledged_by,type,assigned_by,assigned_to,user_tag1,user_tag2,custom1,custom2,custom3,custom4,custom5\n";
        $tocsv="$sqlcount,$a_time,$a_duration,$a_domain,$a_hub,$a_origin,$a_robot,$a_source,$a_subsys,$a_prid,$a_severity,$a_level,$a_visible,$d_suppcount,$a_message,$a_suppkey,$a_sid,$a_acknowledged_by,$a_type,$a_assigned_by,$a_assigned_to,$a_user_tag1,$a_user_tag2,$a_custom1,$a_custom2,$a_custom3,$a_custom4,$a_custom5";
#    $tocsv="$sqlcount,$a_time,$a_duration,$a_robot,$a_source,$a_subsys,$a_prid,$a_severity,$a_level,$a_visible,$d_suppcount,$a_message,$a_suppkey,$a_sid,$a_acknowledged_by,$a_type,$a_assigned_by,$a_assigned_to";
#     print OUTCSV "$sqlcount,$a_time,$a_duration,$a_robot,$a_source,$a_subsys,$a_prid,$a_severity,$a_level,$a_visible,$d_suppcount,$a_message,$a_suppkey,$a_sid,$a_acknowledged_by,$a_type,$a_assigned_by,$a_assigned_to\n";
       $tocsv=~s/[^[:ascii:]]+//g;
	   print OUTCSV "$tocsv\n";
    } 
   if ($tel eq "0")
    { 
    print OUTFILE1 "<tr>\n";
    print OUTFILE1 "<td><pre><font face=\"Tahoma\" size=\"2\" color=\"#000080\">&nbsp;$sqlcount</font></pre></td>\n";
    print OUTFILE1 "<td><pre><font face=\"Tahoma\" size=\"2\" color=\"#000080\">&nbsp;$a_time</font></pre></td>\n";
    print OUTFILE1 "<td><pre>\n";
    if ($a_level eq '0')
    {
       print OUTFILE1 "<font face=\"Tahoma\" size=\"2\" color=\"#00FF00\">&nbsp;&#9632;</font>\n";
    }
    if ($a_level eq '1')
    {
       print OUTFILE1 "<font face=\"Tahoma\" size=\"2\" color=\"#42e5f4\">&nbsp;&#9632;</font>\n";
    }
    if ($a_level eq '2')
    {
       print OUTFILE1 "<font face=\"Tahoma\" size=\"2\" color=\"#1c0cce\">&nbsp;&#9632;</font>\n";
    }
    if ($a_level eq '3')
    {
       print OUTFILE1 "<font face=\"Tahoma\" size=\"2\" color=\"#e5e110\">&nbsp;&#9632;</font>\n";
    }
    if ($a_level eq '4')
    {
       print OUTFILE1 "<font face=\"Tahoma\" size=\"2\" color=\"#e58910\">&nbsp;&#9632;</font>\n";
    } 
    if ($a_level eq '5')
    {
       print OUTFILE1 "<font face=\"Tahoma\" size=\"2\" color=\"#FF0000\">&nbsp;&#9632;</font>\n";
    } 
    if ($a_level eq ' ')
    {
      print OUTFILE1 "<font face=\"Tahoma\" size=\"2\" color=\"#D2FFFF\">&nbsp;&#9632;</font>\n";
    } 


    print OUTFILE1 "</pre></td>\n";
    if ($repcount > 0) {print OUTFILE1 "<td><pre><font face=\"Tahoma\" size=\"2\" color=\"#000080\">&nbsp;$$w0</font></pre></td>\n";}
    if ($repcount > 1) {print OUTFILE1 "<td><pre><font face=\"Tahoma\" size=\"2\" color=\"#000080\">&nbsp;$$w1</font></pre></td>\n";}
    if ($repcount > 2) {print OUTFILE1 "<td><pre><font face=\"Tahoma\" size=\"2\" color=\"#000080\">&nbsp;$$w2</font></pre></td>\n";}
    if ($repcount > 3) {print OUTFILE1 "<td><pre><font face=\"Tahoma\" size=\"2\" color=\"#000080\">&nbsp;$$w3</font></pre></td>\n";}
    if ($repcount > 4) {print OUTFILE1 "<td><pre><font face=\"Tahoma\" size=\"2\" color=\"#000080\">&nbsp;$$w4</font></pre></td>\n";}
    if ($repcount > 5) {print OUTFILE1 "<td><pre><font face=\"Tahoma\" size=\"2\" color=\"#000080\">&nbsp;$$w5</font></pre></td>\n";}
    if ($repcount > 6) {print OUTFILE1 "<td><pre><font face=\"Tahoma\" size=\"2\" color=\"#000080\">&nbsp;$$w6</font></pre></td>\n";}
    if ($repcount > 7) {print OUTFILE1 "<td><pre><font face=\"Tahoma\" size=\"2\" color=\"#000080\">&nbsp;$$w7</font></pre></td>\n";}
    if ($repcount > 8) {print OUTFILE1 "<td><pre><font face=\"Tahoma\" size=\"2\" color=\"#000080\">&nbsp;$$w8</font></pre></td>\n";}
    if ($repcount > 9) {print OUTFILE1 "<td><pre><font face=\"Tahoma\" size=\"2\" color=\"#000080\">&nbsp;$$w9</font></pre></td>\n";}
    if ($repcount > 10) {print OUTFILE1 "<td><pre><font face=\"Tahoma\" size=\"2\" color=\"#000080\">&nbsp;$$w10</font></pre></td>\n";}
    if ($repcount > 11) {print OUTFILE1 "<td><pre><font face=\"Tahoma\" size=\"2\" color=\"#000080\">&nbsp;$$w11</font></pre></td>\n";}
    if ($repcount > 12) {print OUTFILE1 "<td><pre><font face=\"Tahoma\" size=\"2\" color=\"#000080\">&nbsp;$$w12</font></pre></td>\n";}
    if ($repcount > 13) {print OUTFILE1 "<td><pre><font face=\"Tahoma\" size=\"2\" color=\"#000080\">&nbsp;$$w13</font></pre></td>\n";}
    if ($repcount > 14) {print OUTFILE1 "<td><pre><font face=\"Tahoma\" size=\"2\" color=\"#000080\">&nbsp;$$w14</font></pre></td>\n";}
    if ($repcount > 15) {print OUTFILE1 "<td><pre><font face=\"Tahoma\" size=\"2\" color=\"#000080\">&nbsp;$$w15</font></pre></td>\n";}
    if ($repcount > 16) {print OUTFILE1 "<td><pre><font face=\"Tahoma\" size=\"2\" color=\"#000080\">&nbsp;$$w16</font></pre></td>\n";}
    if ($repcount > 17) {print OUTFILE1 "<td><pre><font face=\"Tahoma\" size=\"2\" color=\"#000080\">&nbsp;$$w17</font></pre></td>\n";}
    if ($repcount > 18) {print OUTFILE1 "<td><pre><font face=\"Tahoma\" size=\"2\" color=\"#000080\">&nbsp;$$w18</font></pre></td>\n";}
    print OUTFILE1 "<td><pre><font face=\"Tahoma\" size=\"2\" color=\"#000080\">&nbsp;</font></pre></td>\n";

# a_duration,a_robot,a_source,a_subsys,a_prid,a_severity,a_level,a_visible,d_suppcount,a_message,a_suppkey,a_sid,a_acknowledged

#    print OUTFILE1 "<td><pre><font face=\"Tahoma\" size=\"2\" color=\"#000080\">&nbsp;$a_duration</font></pre></td>\n";
#    print OUTFILE1 "<td><pre><font face=\"Tahoma\" size=\"2\" color=\"#000080\">&nbsp;$a_robot</font></pre></td>\n";
#    print OUTFILE1 "<td><pre><font face=\"Tahoma\" size=\"2\" color=\"#000080\">&nbsp;$a_source</font></pre></td>\n";
#    print OUTFILE1 "<td><pre><font face=\"Tahoma\" size=\"2\" color=\"#000080\">&nbsp;$a_subsys</font></pre></td>\n";
#    print OUTFILE1 "<td><pre><font face=\"Tahoma\" size=\"2\" color=\"#000080\">&nbsp;$a_prid</font></pre></td>\n";
#    print OUTFILE1 "<td><pre><font face=\"Tahoma\" size=\"2\" color=\"#000080\">&nbsp;$a_severity</font></pre></td>\n";
#    print OUTFILE1 "<td><pre>\n";
#    if ($a_level eq '1' || $a_level eq '0')
#    {
#       print OUTFILE1 "<font face=\"Tahoma\" size=\"2\" color=\"#00FF00\">&nbsp; &#9632;</font>\n";
#    }
#    else
#    { 
#      if ($a_level eq '5' || $a_level eq '4')
#       {
#         print OUTFILE1 "<font face=\"Tahoma\" size=\"2\" color=\"#FF0000\">&nbsp; &#9632;</font>\n";
#       } 
#      else
#       {
#         if ($a_level eq ' ')
#          {
#           print OUTFILE1 "<font face=\"Tahoma\" size=\"2\" color=\"#FFFFFF\">&nbsp; &#9632;</font>\n";
#          } 
#         else
#          {
#           print OUTFILE1 "<font face=\"Tahoma\" size=\"2\" color=\"#FFFF00\">&nbsp; &#9632;</font>\n";
#          }
#       } 
#    } 
#    print OUTFILE1 "</pre></td>\n";
#    print OUTFILE1 "<td><pre><font face=\"Tahoma\" size=\"2\" color=\"#000080\">&nbsp;$a_visible</font></pre></td>\n";
#    print OUTFILE1 "<td><pre><font face=\"Tahoma\" size=\"2\" color=\"#000080\">&nbsp;$d_suppcount</font></pre></td>\n";
#    print OUTFILE1 "<td><pre><font face=\"Tahoma\" size=\"2\" color=\"#000080\">&nbsp;$a_message</font></pre></td>\n";
#    print OUTFILE1 "<td><pre><font face=\"Tahoma\" size=\"2\" color=\"#000080\">&nbsp;$a_suppkey</font></pre></td>\n";
#    print OUTFILE1 "<td><pre><font face=\"Tahoma\" size=\"2\" color=\"#000080\">&nbsp;$a_sid</font></pre></td>\n";
#    print OUTFILE1 "<td><pre><font face=\"Tahoma\" size=\"2\" color=\"#000080\">&nbsp;$a_acknowledged_by</font></pre></td>\n";
    print OUTFILE1 "</tr>\n";
    $tel = 1;
    }
   else
    {
# html
    print OUTFILE1 "<tr>\n";
    print OUTFILE1 "<td bgcolor=\"#D2FFFF\"><pre>\n"; 
    print OUTFILE1 "<font face=\"Tahoma\" size=\"2\" color=\"#000080\">&nbsp;$sqlcount</font></pre></td>\n"; 
    print OUTFILE1 "<td bgcolor=\"#D2FFFF\"><pre>\n"; 
    print OUTFILE1 "<font face=\"Tahoma\" size=\"2\" color=\"#000080\">&nbsp;$a_time</font></pre></td>\n"; 
    print OUTFILE1 "<td bgcolor=\"#D2FFFF\"><pre>\n";
    if ($a_level eq '0')
    {
       print OUTFILE1 "<font face=\"Tahoma\" size=\"2\" color=\"#00FF00\">&nbsp;&#9632;</font>\n";
    }
    if ($a_level eq '1')
    {
       print OUTFILE1 "<font face=\"Tahoma\" size=\"2\" color=\"#42e5f4\">&nbsp;&#9632;</font>\n";
    }
    if ($a_level eq '2')
    {
       print OUTFILE1 "<font face=\"Tahoma\" size=\"2\" color=\"#1c0cce\">&nbsp;&#9632;</font>\n";
    }
    if ($a_level eq '3')
    {
       print OUTFILE1 "<font face=\"Tahoma\" size=\"2\" color=\"#e5e110\">&nbsp;&#9632;</font>\n";
    }
    if ($a_level eq '4')
    {
       print OUTFILE1 "<font face=\"Tahoma\" size=\"2\" color=\"#e58910\">&nbsp;&#9632;</font>\n";
    } 
    if ($a_level eq '5')
    {
       print OUTFILE1 "<font face=\"Tahoma\" size=\"2\" color=\"#FF0000\">&nbsp;&#9632;</font>\n";
    } 
    if ($a_level eq ' ')
    {
      print OUTFILE1 "<font face=\"Tahoma\" size=\"2\" color=\"#D2FFFF\">&nbsp;&#9632;</font>\n";
    } 

    print OUTFILE1 "</pre></td>\n"; 
    if ($repcount > 0) 
      {
        print OUTFILE1 "<td bgcolor=\"#D2FFFF\"><pre>\n"; 
        print OUTFILE1 "<font face=\"Tahoma\" size=\"2\" color=\"#000080\">&nbsp;$$w0</font></pre></td>\n"; 
      }
    if ($repcount > 1) 
      {
        print OUTFILE1 "<td bgcolor=\"#D2FFFF\"><pre>\n"; 
        print OUTFILE1 "<font face=\"Tahoma\" size=\"2\" color=\"#000080\">&nbsp;$$w1</font></pre></td>\n"; 
      }
    if ($repcount > 2) 
      {
        print OUTFILE1 "<td bgcolor=\"#D2FFFF\"><pre>\n"; 
        print OUTFILE1 "<font face=\"Tahoma\" size=\"2\" color=\"#000080\">&nbsp;$$w2</font></pre></td>\n"; 
      }
    if ($repcount > 3) 
      {
        print OUTFILE1 "<td bgcolor=\"#D2FFFF\"><pre>\n"; 
        print OUTFILE1 "<font face=\"Tahoma\" size=\"2\" color=\"#000080\">&nbsp;$$w3</font></pre></td>\n"; 
      }
    if ($repcount > 4) 
      {
        print OUTFILE1 "<td bgcolor=\"#D2FFFF\"><pre>\n"; 
        print OUTFILE1 "<font face=\"Tahoma\" size=\"2\" color=\"#000080\">&nbsp;$$w4</font></pre></td>\n"; 
      }
    if ($repcount > 5) 
      {
        print OUTFILE1 "<td bgcolor=\"#D2FFFF\"><pre>\n"; 
        print OUTFILE1 "<font face=\"Tahoma\" size=\"2\" color=\"#000080\">&nbsp;$$w5</font></pre></td>\n"; 
      }
    if ($repcount > 6) 
      {
        print OUTFILE1 "<td bgcolor=\"#D2FFFF\"><pre>\n"; 
        print OUTFILE1 "<font face=\"Tahoma\" size=\"2\" color=\"#000080\">&nbsp;$$w6</font></pre></td>\n"; 
      }
    if ($repcount > 7) 
      {
        print OUTFILE1 "<td bgcolor=\"#D2FFFF\"><pre>\n"; 
        print OUTFILE1 "<font face=\"Tahoma\" size=\"2\" color=\"#000080\">&nbsp;$$w7</font></pre></td>\n"; 
      }
    if ($repcount > 8) 
      {
        print OUTFILE1 "<td bgcolor=\"#D2FFFF\"><pre>\n"; 
        print OUTFILE1 "<font face=\"Tahoma\" size=\"2\" color=\"#000080\">&nbsp;$$w8</font></pre></td>\n"; 
      }
    if ($repcount > 9) 
      {
        print OUTFILE1 "<td bgcolor=\"#D2FFFF\"><pre>\n"; 
        print OUTFILE1 "<font face=\"Tahoma\" size=\"2\" color=\"#000080\">&nbsp;$$w9</font></pre></td>\n"; 
      }
    if ($repcount > 10) 
      {
        print OUTFILE1 "<td bgcolor=\"#D2FFFF\"><pre>\n"; 
        print OUTFILE1 "<font face=\"Tahoma\" size=\"2\" color=\"#000080\">&nbsp;$$w10</font></pre></td>\n"; 
      }
    if ($repcount > 11) 
      {
        print OUTFILE1 "<td bgcolor=\"#D2FFFF\"><pre>\n"; 
        print OUTFILE1 "<font face=\"Tahoma\" size=\"2\" color=\"#000080\">&nbsp;$$w11</font></pre></td>\n"; 
      }
    if ($repcount > 12) 
      {
        print OUTFILE1 "<td bgcolor=\"#D2FFFF\"><pre>\n"; 
        print OUTFILE1 "<font face=\"Tahoma\" size=\"2\" color=\"#000080\">&nbsp;$$w12</font></pre></td>\n"; 
      }
    if ($repcount > 13) 
      {
        print OUTFILE1 "<td bgcolor=\"#D2FFFF\"><pre>\n"; 
        print OUTFILE1 "<font face=\"Tahoma\" size=\"2\" color=\"#000080\">&nbsp;$$w13</font></pre></td>\n"; 
      }
    if ($repcount > 14) 
      {
        print OUTFILE1 "<td bgcolor=\"#D2FFFF\"><pre>\n"; 
        print OUTFILE1 "<font face=\"Tahoma\" size=\"2\" color=\"#000080\">&nbsp;$$w14</font></pre></td>\n"; 
      }
    if ($repcount > 15) 
      {
        print OUTFILE1 "<td bgcolor=\"#D2FFFF\"><pre>\n"; 
        print OUTFILE1 "<font face=\"Tahoma\" size=\"2\" color=\"#000080\">&nbsp;$$w15</font></pre></td>\n"; 
      }
    if ($repcount > 16) 
      {
        print OUTFILE1 "<td bgcolor=\"#D2FFFF\"><pre>\n"; 
        print OUTFILE1 "<font face=\"Tahoma\" size=\"2\" color=\"#000080\">&nbsp;$$w16</font></pre></td>\n"; 
      }
    if ($repcount > 17) 
      {
        print OUTFILE1 "<td bgcolor=\"#D2FFFF\"><pre>\n"; 
        print OUTFILE1 "<font face=\"Tahoma\" size=\"2\" color=\"#000080\">&nbsp;$$w17</font></pre></td>\n"; 
      }
    if ($repcount > 18) 
      {
        print OUTFILE1 "<td bgcolor=\"#D2FFFF\"><pre>\n"; 
        print OUTFILE1 "<font face=\"Tahoma\" size=\"2\" color=\"#000080\">&nbsp;$$w18</font></pre></td>\n"; 
      }
    print OUTFILE1 "<td bgcolor=\"#D2FFFF\"><pre>\n"; 
    print OUTFILE1 "<font face=\"Tahoma\" size=\"2\" color=\"#000080\">&nbsp;</font></pre></td>\n"; 

#    print OUTFILE1 "<td bgcolor=\"#D2FFFF\"><pre>\n"; 
#    print OUTFILE1 "<font face=\"Tahoma\" size=\"2\" color=\"#000080\">&nbsp;$a_duration</font></pre></td>\n"; 
#    print OUTFILE1 "<td bgcolor=\"#D2FFFF\"><pre>\n"; 
#    print OUTFILE1 "<font face=\"Tahoma\" size=\"2\" color=\"#000080\">&nbsp;$a_robot</font></pre></td>\n"; 
#    print OUTFILE1 "<td bgcolor=\"#D2FFFF\"><pre>\n"; 
#    print OUTFILE1 "<font face=\"Tahoma\" size=\"2\" color=\"#000080\">&nbsp;$a_source</font></pre></td>\n"; 
#    print OUTFILE1 "<td bgcolor=\"#D2FFFF\"><pre>\n"; 
#    print OUTFILE1 "<font face=\"Tahoma\" size=\"2\" color=\"#000080\">&nbsp;$a_subsys</font></pre></td>\n"; 
#    print OUTFILE1 "<td bgcolor=\"#D2FFFF\"><pre>\n"; 
#    print OUTFILE1 "<font face=\"Tahoma\" size=\"2\" color=\"#000080\">&nbsp;$a_prid</font></pre></td>\n"; 
#    print OUTFILE1 "<td bgcolor=\"#D2FFFF\"><pre>\n"; 
#    print OUTFILE1 "<font face=\"Tahoma\" size=\"2\" color=\"#000080\">&nbsp;$a_severity</font></pre></td>\n"; 
#    print OUTFILE1 "<td bgcolor=\"#D2FFFF\"><pre>\n";
#    if ($a_level eq '1' || $a_level eq '0')
#    {
#       print OUTFILE1 "<font face=\"Tahoma\" size=\"2\" color=\"#00FF00\">&nbsp; &#9632;</font>\n";
#    }
#    else
#    { 
#      if ($a_level eq '5' || $a_level eq '4')
#       {
#         print OUTFILE1 "<font face=\"Tahoma\" size=\"2\" color=\"#FF0000\">&nbsp; &#9632;</font>\n";
#       } 
#      else
#       {
#        if ($a_level eq ' ')
#         {
#          print OUTFILE1 "<font face=\"Tahoma\" size=\"2\" color=\"#D2FFFF\">&nbsp; &#9632;</font>\n";
#         } 
#        else
#         {
#           print OUTFILE1 "<font face=\"Tahoma\" size=\"2\" color=\"#FFFF00\">&nbsp; &#9632;</font>\n";
#         }
#       }   
#     }
#    print OUTFILE1 "<td bgcolor=\"#D2FFFF\"><pre>\n"; 
#    print OUTFILE1 "<font face=\"Tahoma\" size=\"2\" color=\"#000080\">&nbsp;$a_visible</font></pre></td>\n"; 
#    print OUTFILE1 "<td bgcolor=\"#D2FFFF\"><pre>\n"; 
#    print OUTFILE1 "<font face=\"Tahoma\" size=\"2\" color=\"#000080\">&nbsp;$d_suppcount</font></pre></td>\n"; 
#    print OUTFILE1 "<td bgcolor=\"#D2FFFF\"><pre>\n"; 
#    print OUTFILE1 "<font face=\"Tahoma\" size=\"2\" color=\"#000080\">&nbsp;$a_message</font></pre></td>\n"; 
#    print OUTFILE1 "<td bgcolor=\"#D2FFFF\"><pre>\n"; 
#    print OUTFILE1 "<font face=\"Tahoma\" size=\"2\" color=\"#000080\">&nbsp;$a_suppkey</font></pre></td>\n"; 
#    print OUTFILE1 "<td bgcolor=\"#D2FFFF\"><pre>\n"; 
#    print OUTFILE1 "<font face=\"Tahoma\" size=\"2\" color=\"#000080\">&nbsp;$a_sid</font></pre></td>\n"; 
#    print OUTFILE1 "<td bgcolor=\"#D2FFFF\"><pre>\n"; 
#    print OUTFILE1 "<font face=\"Tahoma\" size=\"2\" color=\"#000080\">&nbsp;$a_acknowledged_by</font></pre></td>\n"; 
    print OUTFILE1 "</tr>\n";
    $tel = 0;
    }
    $teller++;
 }

sub report_over_detail
 {
   if ($tel eq "0")
    { 
    print OUTFILE2 "<tr>\n";
    print OUTFILE2 "<td><pre><font face=\"Tahoma\" size=\"2\" color=\"#000080\">&nbsp;$over0</font></pre></td>\n";
    print OUTFILE2 "<td><pre><font face=\"Tahoma\" size=\"2\" color=\"#000080\">&nbsp;$over2</font></pre></td>\n";
    print OUTFILE2 "<td><pre><font face=\"Tahoma\" size=\"2\" color=\"#000080\">&nbsp;$over1</font></pre></td>\n";
    print OUTFILE2 "</tr>\n";
    $tel = 1;
    }
   else
    {
    print OUTFILE2 "<td bgcolor=\"#D2FFFF\"><pre>\n"; 
    print OUTFILE2 "<font face=\"Tahoma\" size=\"2\" color=\"#000080\">&nbsp;$over0</font></pre></td>\n"; 
    print OUTFILE2 "<td bgcolor=\"#D2FFFF\"><pre>\n"; 
    print OUTFILE2 "<font face=\"Tahoma\" size=\"2\" color=\"#000080\">&nbsp;$over2</font></pre></td>\n"; 
    print OUTFILE2 "<td bgcolor=\"#D2FFFF\"><pre>\n"; 
    print OUTFILE2 "<font face=\"Tahoma\" size=\"2\" color=\"#000080\">&nbsp;$over1</font></pre></td>\n"; 
    print OUTFILE2 "</tr>\n";
    $tel = 0;
    }
    $teller++;
 }

sub calcdate
 {
  $timeval = time();
  @timelist = gmtime ($timeval);
  if (@timelist[6] eq '0') { $dag = "Sunday" }
  if (@timelist[6] eq '1') { $dag = "Monday" }
  if (@timelist[6] eq '2') { $dag = "Tuesday" }
  if (@timelist[6] eq '3') { $dag = "Wednesday" }
  if (@timelist[6] eq '4') { $dag = "Thursday" }
  if (@timelist[6] eq '5') { $dag = "Friday" }
  if (@timelist[6] eq '6') { $dag = "Saturday" }
  if (@timelist[4] eq '0') { $maand = "January" }
  if (@timelist[4] eq '1') { $maand = "February" }
  if (@timelist[4] eq '2') { $maand = "March" }
  if (@timelist[4] eq '3') { $maand = "April" }
  if (@timelist[4] eq '4') { $maand = "May" }
  if (@timelist[4] eq '5') { $maand = "June" }
  if (@timelist[4] eq '6') { $maand = "July" }
  if (@timelist[4] eq '7') { $maand = "Augustus" }
  if (@timelist[4] eq '8') { $maand = "September" }
  if (@timelist[4] eq '9') { $maand = "October" }
  if (@timelist[4] eq '10') { $maand = "November" }
  if (@timelist[4] eq '11') { $maand = "December" }
  $maandnum = @timelist[4] +1;
  if ($maandnum < 10) {$maandnum="0$maandnum";}
  $dagnum = @timelist[3] ;
  if ($dagnum < 10) {$dagnum="0$dagnum";}
  $jaar = @timelist[5] + 1900;
#
# get time 
#
  $intim1 = time();
  $intim2 = localtime($intim1);
  @intime = split(/ +/ , $intim2);
  $intim = $intime[3]; 
  $sql_today="$jaar-$maandnum-$dagnum"; 


  print ("Today its $dag $dagnum $maand ($maandnum) $jaar $intim \n");
 }

sub report_fieldex
{
# Field explanation
print OUTFILE1 "</table>\n";
print OUTFILE1 "<table>\n";
print OUTFILE1 "<table border=\"0\" cellpadding=\"0\" cellspacing=\"0\" style=\"border-collapse: collapse\" bordercolor=\"#111111\" width=\"100%\" id=\"AutoNumber1\">\n";
print OUTFILE1 "<tr>\n";
print OUTFILE1 "<td width=\"100%\" bgcolor=\"#009AD6\" colspan=\"$colspan\" align=\"left\" height=\"30\"><b>\n";
print OUTFILE1 "<pre><font color=\"#D2FFFF\" face=\"MS Sans Serif\" size=\"2\">&nbsp;Field explanation";
print OUTFILE1 "</b></pre></td>\n";
print OUTFILE1 "</tr>\n";

# --- start with blue empty line
$tel=1;
&report_emptyline;
$tel=0;
&report_emptyline; 

# --- Filters
print OUTFILE1 "<tr><td bgcolor=\"#D2FFFF\"><pre>\n"; 
print OUTFILE1 "<font face=\"Tahoma\" size=\"2\" color=\"#000080\">Filters:</font></pre></td>\n";
print OUTFILE1 "<td bgcolor=\"#D2FFFF\" colspan=\"$colspan\" align=\"left\"><pre>\n"; 
print OUTFILE1 "<font face=\"Tahoma\" size=\"2\" color=\"#000080\">Message Inc: $s_filter DB: CA_UIM SQL Table: $sqltable (based on the -al option). nas_transaction_log: every separate alarm record is written where the type column can be used to indicate what action was taken (1: alarm created, 2: alarm updated, 4: acknowledged, 8: assigned/unassigned). nas_transaction_summary: for every alarm only 1 record is written </font></pre></td></tr>\n"; 
#print OUTFILE1 "<font face=\"Tahoma\" size=\"2\" color=\"#000080\">Message Inc: $s_message_inc Message Exc: $s_message_exc Hostname Inc: $s_hostname_inc Hostname Exc: $s_hostname_exc Origin Inc: $s_origin_inc Origin exc: $s_origin_exc Hub Inc: $s_hub_inc Hub exc: $s_hub_exc Robot Inc: $s_robot_inc Robot Exc: $s_robot_exc Probe Inc: $s_probe_inc Probe Exc: $s_probe_exc AsgBy Inc: $s_assignedby_inc AsgBy Exc: $s_assignedby_exc AsgTo Inc: $s_assignedto_inc AsgTo Exc: $s_assignedto_exc Ack Inc: $s_acknowledged_inc Ack Exc: $s_acknowledged_exc LevelInc: $s_level_inc LevelExc: $s_level_exc Visible: $s_visible Suppcount: $s_count_value Starttime: $sql_starthh Endtime: $sql_endhh DB: CA_UIM SQL Table: $sqltable (based on the -al option). nas_transaction_log: every separate alarm record is written where the type column can be used to indicate what action was taken (1: alarm created, 2: alarm updated, 4: acknowledged, 8: assigned/unassigned). nas_transaction_summary: for every alarm only 1 record is written </font></pre></td></tr>\n"; 

print OUTFILE1 "<tr><td width=\"10%\"><pre><font face=\"Tahoma\" size=\"2\" color=\"#000080\">Columns: </font></pre></td>\n";
print OUTFILE1 "<td width=\"90%\"><pre><font face=\"Tahoma\" size=\"2\" color=\"#000080\">Can be defined via the -co option: duration,origin,hub,robot,source,subsys,probe,severity,level,visible,suppcount,message,suppkey,sid,acknowledged,user_tag1,user_tag2,type,assigned_by,assigned_to (if you select as source the nas_transaction_summary you can also use columns: user1, user2, user3, user4 and user5)</font></pre></td></tr>\n";

print OUTFILE1 "<tr><td bgcolor=\"#D2FFFF\"><pre>\n"; 
print OUTFILE1 "<font face=\"Tahoma\" size=\"2\" color=\"#000080\">Count:</font></pre></td>\n";
print OUTFILE1 "<td bgcolor=\"#D2FFFF\" colspan=\"$colspan\" align=\"left\"><pre>\n"; 
print OUTFILE1 "<font face=\"Tahoma\" size=\"2\" color=\"#000080\">Simple counter for the generated alarms</font></pre></td></tr>\n"; 
  
print OUTFILE1 "<tr><td width=\"10%\"><pre><font face=\"Tahoma\" size=\"2\" color=\"#000080\">LastCreated: </font></pre></td>\n";
print OUTFILE1 "<td width=\"90%\"><pre><font face=\"Tahoma\" size=\"2\" color=\"#000080\">Date\\Time the alarms was generated</font></pre></td></tr>\n";

print OUTFILE1 "<tr><td bgcolor=\"#D2FFFF\"><pre>\n"; 
print OUTFILE1 "<font face=\"Tahoma\" size=\"2\" color=\"#000080\">Duration:</font></pre></td>\n";
print OUTFILE1 "<td bgcolor=\"#D2FFFF\" colspan=\"$colspan\" align=\"left\"><pre>\n"; 
print OUTFILE1 "<font face=\"Tahoma\" size=\"2\" color=\"#000080\">dd.hh:mm:ss since initial message was created (based on nimts field) (Last Created - Initial Created)</font></pre></td></tr>\n"; 
  
print OUTFILE1 "<tr><td width=\"10%\"><pre><font face=\"Tahoma\" size=\"2\" color=\"#000080\">Nimts:</font></pre></td>\n";
print OUTFILE1 "<td width=\"90%\"><pre><font face=\"Tahoma\" size=\"2\" color=\"#000080\">Creation date and time of the unique nimid that is common for all exact same messages (counted in field #)</font></pre></td></tr>\n";

print OUTFILE1 "<tr><td bgcolor=\"#D2FFFF\"><pre>\n"; 
print OUTFILE1 "<font face=\"Tahoma\" size=\"2\" color=\"#000080\">Robot:</font></pre></td>\n";
print OUTFILE1 "<td bgcolor=\"#D2FFFF\" colspan=\"$colspan\" align=\"left\"><pre>\n"; 
print OUTFILE1 "<font face=\"Tahoma\" size=\"2\" color=\"#000080\">Machine name where the probe is installed</font></pre></td></tr>\n"; 
  
print OUTFILE1 "<tr><td width=\"10%\"><pre><font face=\"Tahoma\" size=\"2\" color=\"#000080\">Subsys: </font></pre></td>\n";
print OUTFILE1 "<td width=\"90%\"><pre><font face=\"Tahoma\" size=\"2\" color=\"#000080\">Subsys name</font></pre></td></tr>\n";

print OUTFILE1 "<tr><td bgcolor=\"#D2FFFF\"><pre>\n"; 
print OUTFILE1 "<font face=\"Tahoma\" size=\"2\" color=\"#000080\">Probe:</font></pre></td>\n";
print OUTFILE1 "<td bgcolor=\"#D2FFFF\" colspan=\"$colspan\" align=\"left\"><pre>\n"; 
print OUTFILE1 "<font face=\"Tahoma\" size=\"2\" color=\"#000080\">Probe name\\id</font></pre></td></tr>\n"; 
  
print OUTFILE1 "<tr><td width=\"10%\"><pre><font face=\"Tahoma\" size=\"2\" color=\"#000080\">Severity: </font></pre></td>\n";
print OUTFILE1 "<td width=\"90%\"><pre><font face=\"Tahoma\" size=\"2\" color=\"#000080\">Severity of the generated alarm</font></pre></td></tr>\n";

print OUTFILE1 "<tr><td bgcolor=\"#D2FFFF\"><pre>\n"; 
print OUTFILE1 "<font face=\"Tahoma\" size=\"2\" color=\"#000080\">V:</font></pre></td>\n";
print OUTFILE1 "<td bgcolor=\"#D2FFFF\" colspan=\"$colspan\" align=\"left\"><pre>\n"; 
print OUTFILE1 "<font face=\"Tahoma\" size=\"2\" color=\"#000080\">Visible (1) or not viible (0) message</font></pre></td></tr>\n"; 

print OUTFILE1 "<tr><td width=\"10%\"><pre><font face=\"Tahoma\" size=\"2\" color=\"#000080\">#: </font></pre></td>\n";
print OUTFILE1 "<td width=\"90%\"><pre><font face=\"Tahoma\" size=\"2\" color=\"#000080\">Suppress count</font></pre></td></tr>\n";

print OUTFILE1 "<tr><td bgcolor=\"#D2FFFF\"><pre>\n"; 
print OUTFILE1 "<font face=\"Tahoma\" size=\"2\" color=\"#000080\">Message:</font></pre></td>\n";
print OUTFILE1 "<td bgcolor=\"#D2FFFF\" colspan=\"$colspan\" align=\"left\"><pre>\n"; 
print OUTFILE1 "<font face=\"Tahoma\" size=\"2\" color=\"#000080\">Message text</font></pre></td></tr>\n"; 

print OUTFILE1 "<tr><td width=\"10%\"><pre><font face=\"Tahoma\" size=\"2\" color=\"#000080\">Suppression Key: </font></pre></td>\n";
print OUTFILE1 "<td width=\"90%\"><pre><font face=\"Tahoma\" size=\"2\" color=\"#000080\">When an alarm is generated at can optionally receive a suppression key.  When later a clear alarm is received with the same suppression key, this alarm will be removed from the active alarm list.  The alarm text must even not be the same, as long as the suppression key is the same the alarm will be removed.  This can also be used with user generated alarms.</font></pre></td></tr>\n";

# --- start with blue empty line
$tel=1;
&report_emptyline;
$tel=0;
&report_emptyline;   
}


sub getparams
{
  usage() unless @ARGV gt 0;
  $argCount=$#ARGV + 1;

  for ($i=0 ; $i < $argCount ; $i++) 
  {
	if ($ARGV[$i] =~ /^-\?/)
	{ 
		usage();
	} 

   	if ($ARGV[$i] =~ /^-di/ ) 
	{ 
		$outdir =  substr($ARGV[$i],3);
        }
	if ($ARGV[$i] =~ /^-fi/ )   
	{ 
		$outfile =  substr($ARGV[$i],3);
        }
	if ($ARGV[$i] =~ /^-x/ )   
	{ 
		$do_csv =  substr($ARGV[$i],2);
        }
	if ($ARGV[$i] =~ /^-rs/ )   
	{ 
		$reportseconds =  substr($ARGV[$i],3);
        }
	if ($ARGV[$i] =~ /^-su/ )   
	{ 
		$sqluser =  substr($ARGV[$i],3);
        }
	if ($ARGV[$i] =~ /^-sp/ )   
	{ 
		$sqlpass =  substr($ARGV[$i],3);
        }
	if ($ARGV[$i] =~ /^-sr/ )   
	{ 
		$sqlserver =  substr($ARGV[$i],3);
        }
	if ($ARGV[$i] =~ /^-wr/ )   
	{ 
		$s_overview =  substr($ARGV[$i],3);
        }
	if ($ARGV[$i] =~ /^-ty/ )   
	{ 
		$sqlopt =  substr($ARGV[$i],3);
        }
	if ($ARGV[$i] =~ /^-al/ )   
	{ 
		$sqlcollection =  substr($ARGV[$i],3);
        }
	if ($ARGV[$i] =~ /^-nb/ )   
	{ 
		$numlines =  substr($ARGV[$i],3);
        }
	if ($ARGV[$i] =~ /^-8/ )   
	{ 
		$database =  substr($ARGV[$i],2);
        }
	if ($ARGV[$i] =~ /^-1/ )   
	{ 
		$sql_start =  substr($ARGV[$i],2);
        }
	if ($ARGV[$i] =~ /^-3/ )   
	{ 
		$sql_end =  substr($ARGV[$i],2);
        }
	if ($ARGV[$i] =~ /^-2/ )   
	{ 
		$sql_starthh =  substr($ARGV[$i],2);
        }
	if ($ARGV[$i] =~ /^-4/ )   
	{ 
		$sql_endhh =  substr($ARGV[$i],2);
        }
	if ($ARGV[$i] =~ /^-9/ )   
	{ 
		$repview =  substr($ARGV[$i],2);
        }
	if ($ARGV[$i] =~ /^-mi/ )   
	{ 
		$s_message_inc =  substr($ARGV[$i],3);
        }
	if ($ARGV[$i] =~ /^-me/ )   
	{ 
		$s_message_exc =  substr($ARGV[$i],3);
        }
	if ($ARGV[$i] =~ /^-hi/ )   
	{ 
		$s_hostname_inc =  substr($ARGV[$i],3);
        }
	if ($ARGV[$i] =~ /^-he/ )   
	{ 
		$s_hostname_exc =  substr($ARGV[$i],3);
        }
	if ($ARGV[$i] =~ /^-ti/ )   
	{ 
		$s_robot_inc =  substr($ARGV[$i],3);
        }
	if ($ARGV[$i] =~ /^-te/ )   
	{ 
		$s_robot_exc =  substr($ARGV[$i],3);
        }
	if ($ARGV[$i] =~ /^-gi/ )   
	{ 
		$s_origin_inc =  substr($ARGV[$i],3);
        }
	if ($ARGV[$i] =~ /^-ge/ )   
	{ 
		$s_origin_exc =  substr($ARGV[$i],3);
        }
	if ($ARGV[$i] =~ /^-ui/ )   
	{ 
		$s_hub_inc =  substr($ARGV[$i],3);
        }
	if ($ARGV[$i] =~ /^-ue/ )   
	{ 
		$s_hub_exc =  substr($ARGV[$i],3);
        }
	if ($ARGV[$i] =~ /^-ai/ )   
	{ 
		$s_assignedby_inc =  substr($ARGV[$i],3);
        }
	if ($ARGV[$i] =~ /^-ae/ )   
	{ 
		$s_assignedby_exc =  substr($ARGV[$i],3);
        }
	if ($ARGV[$i] =~ /^-ki/ )   
	{ 
		$s_acknowledged_inc =  substr($ARGV[$i],3);
        }
	if ($ARGV[$i] =~ /^-ke/ )   
	{ 
		$s_acknowledged_exc =  substr($ARGV[$i],3);
        }
	if ($ARGV[$i] =~ /^-li/ )   
	{ 
		$s_level_inc =  substr($ARGV[$i],3);
        }
	if ($ARGV[$i] =~ /^-le/ )   
	{ 
		$s_level_exc =  substr($ARGV[$i],3);
        }
	if ($ARGV[$i] =~ /^-si/ )   
	{ 
		$s_assignedto_inc =  substr($ARGV[$i],3);
        }
	if ($ARGV[$i] =~ /^-se/ )   
	{ 
		$s_assignedto_exc =  substr($ARGV[$i],3);
        }
	if ($ARGV[$i] =~ /^-oi/ )   
	{ 
		$s_probe_inc =  substr($ARGV[$i],3);
        }
	if ($ARGV[$i] =~ /^-oe/ )   
	{ 
		$s_probe_exc =  substr($ARGV[$i],3);
        }
	if ($ARGV[$i] =~ /^-cv/ )   
	{ 
		$s_count_value =  substr($ARGV[$i],3);
        }
	if ($ARGV[$i] =~ /^-co/ )   
	{ 
		$report =  substr($ARGV[$i],3);
        }
	if ($ARGV[$i] =~ /^-vi/)   
	{ 
		$s_visible =  substr($ARGV[$i],3);
        }
	if ($ARGV[$i] =~ /^-vb/ )   
	{ 
		$verbose =  substr($ARGV[$i],3);
        }
	if ($ARGV[$i] =~ /^-mm/ )   
	{ 
		$s_month =  substr($ARGV[$i],3);
        }
#	if ($ARGV[$i] =~ /^-b/ )   
#	{ 
#		$s_days =  substr($ARGV[$i],2);
#        }
	if ($ARGV[$i] =~ /^-bh/ )   
	{ 
		$s_days =  substr($ARGV[$i],3);
        }
	if ($ARGV[$i] =~ /^-bm/ )   
	{ 
		$s_minutes =  substr($ARGV[$i],3);
                $s_days="";  
        }
	if ($ARGV[$i] =~ /^-bh/ )   
	{ 
		$s_hours =  substr($ARGV[$i],3);
        }
	if ($ARGV[$i] =~ /^-gu/ )   
	{ 
		$gui =  substr($ARGV[$i],3);
        }
	if ($ARGV[$i] =~ /^-db/ )   
	{ 
		$debug =  substr($ARGV[$i],3);
        }
	if ($ARGV[$i] =~ /^-lh/ ) 
	{ 
		$sql_hostname =  substr($ARGV[$i],3);
        }
	if ($ARGV[$i] =~ /^-lg/ ) 
	{ 
		$sql_origin =  substr($ARGV[$i],3);
        }
	if ($ARGV[$i] =~ /^-lm/ ) 
	{ 
		$sql_message =  substr($ARGV[$i],3);
        }
	if ($ARGV[$i] =~ /^-lo/ ) 
	{ 
		$sql_probe =  substr($ARGV[$i],3);
        }
	if ($ARGV[$i] =~ /^-lu/ ) 
	{ 
		$sql_hub =  substr($ARGV[$i],3);
        }
	if ($ARGV[$i] =~ /^-lt/ ) 
	{ 
		$sql_robot =  substr($ARGV[$i],3);
        }
	if ($ARGV[$i] =~ /^-fq/ ) 
	{ 
		$fqdn_strip =  substr($ARGV[$i],3);
        }


  }

# --- get credentials and settings from common file (nimsoft_generic.dat)
($Server_With_Rest_Interface, $RestServicePort, $UIM_username, $UIM_password,$UIM_Domain,$UIM_Hub,$UIM_Robot,$sql_server,$sql_user,$sql_password,$sql_db,$sql_type,$use_https,$sql_driver,$sql_dsn) = nimsoft_generic::getRestData();

# print "nimsoft_generic: $sql_server $sql_user $sql_password $sql_db $sql_type\n";

# -- de-crypt (need $encoded)
$decoded = decode_base64($UIM_password);
$UIM_password = RC4($crkey, $decoded);

$decoded = decode_base64($sql_password);
$sql_password = RC4($crkey, $decoded);

# --- if not used in parameters, use defaults from generic
if ($sqluser ne '')  {$sql_user=$sqluser; }
if ($sqlpass ne '') { $sql_password=$sqlpass; }
if ($sqlserver ne '') { $sql_server=$sqlserver; }
if ($database ne '') { $database=$sql_db; }
if ($sql_driver ne '') {$sql_driver="SQL Server"}

# --- set defaults
  if ($sqlopt eq '1' || $sqlopt eq '')
   {
      $sqltype="ms";
   }
  if ($sqlopt eq '2')
   {
      $sqltype="my";
   }
  if ($sql_type eq 'mssql')
   {
      $sqltype="ms";
   } 
  if ($sql_type eq 'mysql')
   {
      $sqltype="my";
   } 
  if ($sqlcollection eq '')
   {
     $sqlcollection="t";
   }
  if ($sqlcollection eq 't') { $sqltable="nas_transaction_log";}
  if ($sqlcollection eq 's') { $sqltable="nas_transaction_summary";}
  if ($sqlcollection eq 'y') { $sqltable="AlarmTransactionLog";}
  if ($database eq '')
   {
     $database="CA_UIM";
   }

  if ($sql_user eq '')
   {
     if ($sqltype eq 'my')
      {
        $sqluser = "root";
      }
     if ($sqltype eq 'ms')
      {
        if ($sql_pass eq '')
         {
           $sql_user="trusted";
         }
        else
         {
           $sql_user = "sa";
         }
      }
   }

  if ($sql_user eq 'trusted')
      {
         $trusted=yes;
      }

  if ($sql_password eq '')
   {
     if ($sql_user eq 'trusted')
      {
         $trusted=yes;
      }
     else
      {
        print "By default we try to access the local mysql with user: $sql_user\n";
        print "BUT you need at least the -sp parameter to give a password\n";
        exit 12;
      }
   }
  if ($sql_server eq '')
   {
     $sql_server=$hostn;
     if ($sqltype eq 'my')
      {
        print "We will connect the the mysql DB on server: $sql_server\n";
      }
     if ($sqltype eq 'ms')
      {
        print "We will connect the the mssql DB on server: $sql_server and userid: $sql_user\n";
      }
   }
  if ($outdir eq '')
   {
     if ($myos eq 'MSWin32')
      {
        $outdir = "c:\\temp";
      }
     else
      {
        $outdir = "/opt/temp";
      }
   }
  if ($fqdn_strip eq '')
   {
     $fqdn_strip = "y";
   }
  if ($repview eq '')
   {
     $repview = "n";
   }
  if ($do_csv eq '')
   {
     $do_csv = "n";
   }
  if ($outfile eq '')
   {
     $outfile = "report\_nimsoft\_alarm\_reporter";
   }
  if ($s_visible eq '')
   {
     $s_visible="y";
#     print "Select visible and non-visible messages\n";
   }
  if ($s_visible ne 'y' && $s_visible ne 'n' && $s_visible ne 'o')
   {
     print "The report visible option -vi can only have y, n or o as parameter\n";
     exit 12;
   } 
  else
   {
     print "Reporting option -vi visible: $s_visible (show visible and unvisible)\n"; 
   }
  if ($s_overview eq '')
   {
     $s_overview=Y;
   }
  if ($reportseconds eq '')
   {
     $reportloop=N;
   }
  else
   {
     $reportloop=Y;
   }
  if ($s_message_inc ne '')
   {
      print (color("black on_green "),"Using message include filter: $s_message_inc",color("reset"),"\n");
   } 
  if ($s_message_exc ne '')
   {
     print (color("black on_green "),"Using message exclude filter: $s_message_exc",color("reset"),"\n");
   } 
  if ($s_hostname_inc ne '')
   {
     print (color("black on_green "),"Using hostname include filter: $s_hostname_inc",color("reset"),"\n");
   } 
  if ($s_hostname_exc ne '')
   {
     print (color("black on_green "),"Using hostname exclude filter: $s_hostname_exc",color("reset"),"\n");
   } 
  if ($s_origin_inc ne '')
   {
     print (color("black on_green "),"Using origin include filter: $s_origin_inc",color("reset"),"\n");
   } 
  if ($s_origin_exc ne '')
   {
     print (color("black on_green "),"Using origin exclude filter: $s_origin_exc",color("reset"),"\n");
   } 
  if ($s_hub_inc ne '')
   {
     print (color("black on_green "),"Using hub include filter: $s_hub_inc",color("reset"),"\n");
   } 
  if ($s_hub_exc ne '')
   {
     print (color("black on_green "),"Using hub exclude filter: $s_hub_exc",color("reset"),"\n");
   } 
  if ($s_assignedby_inc ne '')
   {
     print (color("black on_green "),"Using assigned_by include filter: $s_assignedby_inc",color("reset"),"\n");
   } 
  if ($s_assignedby_exc ne '')
   {
     print (color("black on_green "),"Using assigned_by exclude filter: $s_assignedby_exc",color("reset"),"\n");
   } 
  if ($s_acknowledged_inc ne '')
   {
     print (color("black on_green "),"Using acknowledged_by include filter: $s_acknowledged_inc",color("reset"),"\n");
   } 
  if ($s_acknowledged_exc ne '')
   {
     print (color("black on_green "),"Using acknowledged_by exclude filter: $s_acknowledged_exc",color("reset"),"\n");
   } 
  if ($s_level_inc ne '')
   {
     print (color("black on_green "),"Using level include filter: $s_level_inc",color("reset"),"\n");
   } 
  if ($s_level_exc ne '')
   {
     print (color("black on_green "),"Using level exclude filter: $s_level_exc",color("reset"),"\n");
   } 
  if ($s_assignedto_inc ne '')
   {
     print (color("black on_green "),"Using assigned_to include filter: $s_assignedto_inc",color("reset"),"\n");
   } 
  if ($s_assignedto_exc ne '')
   {
     print (color("black on_green "),"Using assigned_to exclude filter: $s_assignedto_exc",color("reset"),"\n");
   } 
  if ($s_robot_inc ne '')
   {
     print (color("black on_green "),"Using robot include filter: $s_robot_inc",color("reset"),"\n");
   } 
  if ($s_robot_exc ne '')
   {
     print (color("black on_green "),"Using robot exclude filter: $s_robot_exc",color("reset"),"\n");
   } 
  if ($s_probe_inc ne '')
   {
     print (color("black on_green "),"Using probe include filter: $s_probe_inc",color("reset"),"\n");
   } 
  if ($s_probe_exc ne '')
   {
     print (color("black on_green "),"Using probe exclude filter: $s_probe_exc",color("reset"),"\n");
   } 
  if ($s_count_value ne '')
   {
     print (color("black on_green "),"Using suppression count filter: $s_count_value",color("reset"),"\n");
   } 
  if ($sql_hub ne '')
   {
     print (color("black on_green "),"Using SQL hub filter: $sql_hub",color("reset"),"\n");
   } 
  if ($sql_origin ne '')
   {
     print (color("black on_green "),"Using SQL origin filter: $sql_origin",color("reset"),"\n");
   } 
  if ($sql_hostname ne '')
   {
     print (color("black on_green "),"Using SQL hostname filter: $sql_hostname",color("reset"),"\n");
   } 
  if ($sql_robot ne '')
   {
     print (color("black on_green "),"Using SQL robot filter: $sql_robot",color("reset"),"\n");
   } 
  if ($sql_probe ne '')
   {
     print (color("black on_green "),"Using SQL probe filter: $sql_probe",color("reset"),"\n");
   } 
  if ($sql_message ne '')
   {
     print (color("black on_green "),"Using SQL message filter: $sql_message",color("reset"),"\n");
   } 
# --- date/time selection
  if ($s_days eq '' && $s_minutes eq '' && $sql_start eq '' && $sql_stop eq '' && $s_month eq '' && $sql_starthh eq '' && $sql_endhh eq '')
   {
     $s_days=24;
   }
#  else 
#   {
#     $s_days="";
#   }   
  if ($sql_start ne '')
   {
     if ($sql_start =~ /^(\d\d\d\d)-(\d\d)-(\d\d)\s(\d\d):(\d\d)/)
      {
#        print "$sql_start is a valid date\n";
      }
     else
      {
        print "Parameter -1 is NOT a valid start date (format: \"yyyy-mm-dd hh:mm\")\n";
        exit;
      }
   }
  if ($sql_starthh ne '')
   {
     if ($sql_starthh =~ /^(\d\d):(\d\d)/)
      {
#        print "$sql_starthh is a valid start hour\n";
      }
     else
      {
        print "Parameter -2 is NOT a valid start hour format (format: \"hh:mm\")\n";
        exit;
      }
   }

  if ($sql_end ne '' && $sql_start eq '')
   {
     print " When using the report end date/time via param -3 you must also use the start date/time parameter -1\n";
     exit;
   }

  if ($sql_end ne '')
   {
     if ($sql_end =~ /^(\d\d\d\d)-(\d\d)-(\d\d)\s(\d\d):(\d\d)/)
      {
#        print "$sql_end is a valid date\n";
      }
     else
      {
        print "Parameter -3 is NOT a valid end date (format: \"yyyy-mm-dd hh:mm\")\n";
        exit;
      }
   }
  if ($sql_endhh ne '')
   {
     if ($sql_endhh =~ /^(\d\d):(\d\d)/)
      {
#        print "$sql_endhh is a valid date\n";
      }
     else
      {
        print "Parameter -4 is NOT a valid end hour format (format: \"hh:mm\")\n";
        exit;
      }
   }

  

# --- end date/time

  if ($numlines eq '')
   {
     $numlines=5000;
   }

# --- if no extra columns are defined via -co, use some defaults
   if ($report eq '' && $sqlcollection ne 's')
    {
#     $report="duration,hub,robot,probe,severity,level,visible,suppcount,message,suppkey,sid,acknowledged,user_tag1,user_tag2";
     $report="duration,origin,hub,robot,hostname,probe,severity,level,visible,suppcount,message,suppkey,sid,acknowledged,user_tag1,user_tag2";
    }
   if ($report eq '' && $sqlcollection eq 's')
    {
     $report="origin,hub,robot,probe,severity,level,visible,suppcount,message,suppkey,sid,acknowledged,user_tag1,user_tag2,custom1,custom2,custom3,custom4,custom5";
    }

# --- create dir + name
if ($myos eq 'MSWin32')
 {      
   $outfil = "$outdir\\$outfile";
 }
else
 {      
   $outfil = "$outdir/$outfile";
 }

}

sub splitreport
{
  @rep = split(/,/ , $report);
  $repcount=@rep;
  $repcc=0;
# --- create headers and vars
  if ($verbose eq 'Y') 
   {
     while ($rep[$repcc] ne '')
      {             
#        print "$rep[$repcc]\n";
        $repcc++;
      }
   } 

}

sub detail_line
{
     &report_detail;
     $sqlcount++;
# - repeat title every xx lines
     $totcol++;
     if ($totcol > 30)
      {
       &report_header2();
       $totcol=0;
      }
}

sub sql_filters
{
if ($sql_hostname ne '' || $sql_origin ne '' || $sql_message ne '' || $sql_probe ne '' || $sql_hub ne '' || $sql_robot ne '')
 {
    $sql_filt="y";
    $sql_f1="-";
    $sql_f2="-";
    if ($sql_hostname ne '')
     {
        $sql_f1="hostname like '%$sql_hostname%'";
        $sql_f2="l.hostname like '%$sql_hostname%'";  
     }
    if ($sql_origin ne '')
     {
        if ($sql_f1 eq '-')
         {
            $sql_f1="origin like '%$sql_origin%'";
            $sql_f2="l.origin like '%$sql_origin%'";  
         }
       else
         {
            $sql_f1="$sql_f1 and origin like '%$sql_origin%'";
            $sql_f2="$sql_f2 and l.origin like '%$sql_origin%'";  
         }
     }
    if ($sql_message ne '')
     {
        if ($sql_f1 eq '-')
         {
            $sql_f1="message like '%$sql_message%'";
            $sql_f2="l.message like '%$sql_message%'";  
         }
       else
         {
            $sql_f1="$sql_f1 and message like '%$sql_message%'";
            $sql_f2="$sql_f2 and l.message like '%$sql_message%'";  
         }
     }
    if ($sql_probe ne '')
     {
        if ($sql_f1 eq '-')
         {
            $sql_f1="prid like '%$sql_probe%'";
            $sql_f2="l.prid like '%$sql_probe%'";  
         }
       else
         {
            $sql_f1="$sql_f1 and prid like '%$sql_probe%'";
            $sql_f2="$sql_f2 and l.prid like '%$sql_probe%'";  
         }
     }
    if ($sql_hub ne '')
     {
        if ($sql_f1 eq '-')
         {
            $sql_f1="hub like '%$sql_hub%'";
            $sql_f2="l.hub like '%$sql_hub%'";  
         }
       else
         {
            $sql_f1="$sql_f1 and hub like '%$sql_hub%'";
            $sql_f2="$sql_f2 and l.hub like '%$sql_hub%'";  
         }
     }
    if ($sql_robot ne '')
     {
        if ($sql_f1 eq '-')
         {
            $sql_f1="robot like '%$sql_robot%'";
            $sql_f2="l.robot like '%$sql_robot%'";  
         }
       else
         {
            $sql_f1="$sql_f1 and robot like '%$sql_robot%'";
            $sql_f2="$sql_f2 and l.robot like '%$sql_robot%'";  
         }
     }
 }
else
 {
    $sql_filt="n";
 }  
}

sub usage
 {
   print "      \n";
   print "Create reports on the nas logging sql tables, version: $version\n";
   print (color("black on_yellow "),"Example: nimsoft_alarm_reporter -bh\"# hours to report on\"",color("reset"),"\n");
   print "      \n";
   print "Note: Default access parameters must be defined in nimsoft_generic.dat\n";
   print "      (You can use -u for the sql userid and -p for the not crypted password)\n";
#   print "      \n";
   print (color("black on_green "),"  Input parameters:",color("reset"),"\n");
   print "  -rs: (optional) number of seconds between report + html updates\n";
#   print "  -su: sql userid (default: mysql: root mssql: sa) (trusted: for windows auth.)\n";
#   print "  -sp: sql password *** required ***\n";
#   print "  -sr: repository/server name: default local host\n";
   print "  -wr: create overvieW report (Y, N) Default Y\n";
#   print "  -ty: 1: mssql 2: mysql (default: 1)\n";
   print "  -al: t: use nas_transaction_log (default) s: nas_transaction_summary y: alarmtransactionlog\n";
#   print "  -8: database name, default: CA_UIM\n"; 
#   print "      \n";
   print (color("black on_green "),"  Date Selection:",color("reset"),"\n");
   print "  -bh:(optional) # of hours to go back to start report (default: 24)\n";
   print "  -bm:(optional) # of minutes to go back to start report\n";
   print "  -mm:(optional) # of months to go back (0: current month, 1: prev month)\n"; 
   print "  -1: (optional) start date-time (format: \"yyyy-mm-dd hh:mm\")\n";
   print "  -2: (optional) start hour filter (format hh:mm) on todays date\n";
   print "  -3: (optional) end date-time (format: \"yyyy-mm-dd hh:mm\")\n";
   print "  -4: (optional) end hour filter (format hh:mm) on todays date\n";
#   print "      \n";
   print (color("black on_green "),"  Report Columns:",color("reset"),"\n");
   print "  -co: give the columns you want in your report:\n";
   print "  (duration,robot,source,hostname,subsys,probe,severity,visible,suppcount)\n";
   print "  (message,suppkey,sid,acknowledged,user_tag1,user_tag2,hub,origin,level)\n";
   print "  (assigned_by,assigned_to)\n";
#   print "      \n";
   print (color("black on_green"),"SQL filters (LIKE, % will be added before and after):",color("reset"),"\n");
   print "  -lm: Like Message (used as where like clause in sql)\n";
   print "  -lh: Like Hostname (used as where like clause in sql)\n";
   print "  -lg: Like oriGin (used as where like clause in sql)\n";
   print "  -lu: Like hUb (used as where like clause in sql)\n";
   print "  -lt: Like roboT (used as where like clause in sql)\n";
   print "  -lo: Like prObe (used as where like clause in sql)\n";
   print (color("black on_green "),"  Search Criteria: (regex)",color("reset"),"\n");
   print "  -mi/me: message text include/exclude\n";
   print "  -hi/he: hostname include/exclude\n";
   print "  -gi/ge: oriGin include/exclude\n";
   print "  -ui/ue: hUb include/exclude\n";
   print "  -ti/te: roboT include/exclude\n";
   print "  -oi/oe: prObe include/exclude\n";
   print "  -ai/ae: assigned_by include/exclude\n";
   print "  -si/se: assigned_to include/exclude\n";
   print "  -ki/ke: acknowledged_by include/exclude\n";
   print "  -li/le: level include/exclude (0-5), can be: \"3|4|5\"\n";
   print "  -cv: (suppress)count value regex filter (1: first occurrence)\n";
   print "  -vi: Visible y: show visible (default) n: no visible reported o: only visible\n";
#   print "      \n";
   print (color("black on_green "),"  Output:",color("reset"),"\n");
   print "  -di: output directory (default: c:\\temp)\n";
   print "  -fi: output file (default: report_nimsoft_alarm_reporter)\n";
   print "  -9: View report directly (y,n) Default: n\n";
   print "  -x: create csv file (y) or html report (n) Default: n\n";
   print "  -nb: maximum reported lines. Default: 5000\n";
   print "  -fq: fqdn strip on hostname column (y,n) default: y\n";
   print "      \n";
   exit 0;
}

