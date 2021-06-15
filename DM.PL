#!c:/perl/bin/perl.exe
#
# DataMart Browser Script
# 05/18/00 Craig Fitzgerald
#
#
# To implement this script on a new machine you must:
# 1. Put this file (dm.pl) in your CGI directory.
# 2. Look at/Update  The Database connection info inits in %Context 
# 3. If you intend to run this from the command line (for debugging)
#	  find the sub CommandLineInit and fix the paths first
#

use strict;
use DBI;
use CGI::Carp qw(fatalsToBrowser);
use DBD::MySQL;
use Win32::OLE 'in';
use GD;
use Data::Dumper;
use URI::Escape;


# Constants:
my $SubDir          = "dmdata";		 # name of subdir created in HtDocs and in cgi-bin dirs
my $ContextCount    = "last.txt";	 # name of special file used to store last context id
my $BackgroundColor = "#fff6ee";		 # screen backround color 	"#ffffff"	"linen"
my $TableColor 	  = "#ffffff";		 # table background color 	"#ffffff"	"linen"	"#eef6ff"
my $TableHeaderColor= "#cccccc";		 # table header bkg color 	"#eeeeee"	"linen"
my $TableAltRowColor= "#eeeeee";		 # every 2nd row bkg colr 	"#eeeeee"	"linen" $BackgroundColor;


# The TDistribution at the 90% confidence level, used for calculating the 
# Confidence interval in the Price/Quantity Graph I got this from Jeanese
# who generated it in SAS.
my @TDistr =
	(0,	6.31375,	2.91999,	 2.35336,  2.13185,  2.01505,
			1.94318,	1.89458,	 1.85955,  1.83311,  1.81246,
			1.79588,	1.78229,	 1.77093,  1.76131,  1.75305,
			1.74588,	1.73961,	 1.73406,  1.72913,  1.72472,
			1.72074,	1.71714,	 1.71387,  1.71088,  1.70814,
			1.70562,	1.70329,	 1.70113,  1.69913,  1.69726,
			1.69552,	1.69389,	 1.69236,  1.69092,  1.68957,
			1.6883, 	1.68709,	 1.68595,  1.68488,  1.68385,
			1.68288,	1.68195,	 1.68107,  1.68023,  1.67943,
			1.67866,	1.67793,	 1.67722,  1.67655,  1.67591,
			1.67528,	1.67469,	 1.67412,  1.67356,  1.67303,
			1.67252,	1.67203,	 1.67155,  1.67109,  1.67065,
			1.67022,	1.6698, 	 1.6694,	  1.66901,  1.66864,
			1.66827,	1.66792,	 1.66757,  1.66724,  1.66691,
			1.6666, 	1.66629,	 1.666,	  1.66571,  1.66543,
			1.66515,	1.66488,	 1.66462,  1.66437,  1.66412,
			1.66388,	1.66365,	 1.66342,  1.6632,   1.66298,
			1.66277,	1.66256,	 1.66235,  1.66216,  1.66196,
			1.66177,	1.66159,	 1.6614,	  1.66123,  1.66105,
			1.66088,	1.66071,	 1.66055,  1.66039,  1.66023,
			1.66008,	1.65993,	 1.65978,  1.65964,  1.6595,
			1.65936,	1.65922,	 1.65909,  1.65895,  1.65882,
			1.6587, 	1.65857,	 1.65845,  1.65833,  1.65821,
			1.6581, 	1.65798,	 1.65787,  1.65776,  1.65765);


# color values used for plotting.  The first color (black) is always
# used for plotting the Engineers Estimate in the LIP graph
#
my @ColorSet = ([0, 0, 0	 ], 
                [58,57,193  ],
					 [219,4,1	 ],
					 [34,104,24  ],
					 [183,137,183],
					 [0,0,112	 ],
					 [16,170,17  ],
					 [206,104,1  ],
					 [207,218,15 ],
					 [122,122,122],
					 [108,18,17  ],
					 [69,24,93	 ],
					 [83,162,159 ],
					 [181,21,126 ],
					 [64,64,255  ],
					 [0,0,0      ],
					 [64,255,64  ],
					 [255,128,128]);


# Globals:

# %Context: user context
#
# Context contains:
# > Database connection info
# > Keys and various user data through GET params
# > User changable Metrics for graph generation fns
# > Calculated Metrics for graph generation fns
#
# This hash gets loaded and saved to a context file.
# The values defined here are just the starting default values.
#
#
my %Context =
	(
	Provider    	=> "", 	 	# MSOLAP for olap,  MSDASQL or blank for ODBC
	DataSource   	=> "CRAIGF", 	 	# CRAIGF or DSSNTS for OLAP, DSN for ODBC
	Catalog     	=> "dss",		 	# dss for OLAP, blank for ODBC
	Cube         	=> "newprofile", 	# newprofile for OLAP, blank for ODBC

	LIPGraph			=> 1,             # Default is graph rather than table for LIP
	LIPPrices  		=> 1,             # Display Prices at right edge of graph

	LIPYSize 	  	=> 300, 			 	# default height of item profile image
	LIPLeftEdge		=> 80,				# space for labels at left of item profile image
	LIPRightEdge	=> 90,
	LIPTopEdge		=> 10,
	LIPBottomEdge	=> 35,				# space for labels at bottom of item profile image
	LIPPointSize  	=> 7,   			 	# diameter of points on the item profile image
	LIPMinDev 	   => 1.0,			 	# default minimum deviation for LIP graph
	LIPMaxItems    => 80,  			 	# max # of items (cols) in the LIP graph

	PQXSize 		 	=> 550, 			 	# default height of Price/Quan graph image
	PQYSize 		 	=> 350, 			 	# default height of Price/Quan graph image
	PQLeftEdge	 	=> 85,  			 	# space for labels at left of Price/Quan image
	PQRightEdge	 	=> 25,  			 	# space for   Price/Quan image
	PQTopEdge	 	=> 10,  			 	# space for   Price/Quan image
	PQBottomEdge 	=> 20, 			 	# space for   Price/Quan image
	PQPointSize  	=> 7,  			 	# diameter of points on the Price/Quan image

	PQLogPrice		=> 1,
	PQLogQuan		=> 1,

	HalfYLabel   	=> 6	  			 	# 1/2 the height of the label text on item profile image
	);


my $DataDir;	 		# DataDir (in docs tree) stores graphs and dingbats
my $ContextDir; 		# ContextDir (in cgi-bin tree) stores context data and notes files

my $ContextID;			# The user ID used to save/load the context file - from extra path info
my $Target;		 		# The page to display - from extra path info

my @ImageMapEntries;	# ImageMap entries, used when creating the graphs

#############################################################################
MAIN:

	if (!$ENV{"GATEWAY_INTERFACE"}) # running from command line
		{
		if (!$ARGV[0])
			{
			print ("Usage dm.pl ContextID [PageName]\n");
			exit (0);
			}
		$ContextID = $ARGV[0];
		$Target	  = $ARGV[1];
		CommandLineInit (); # initialize script
		}
	else # running from web server
		{
		($ContextID, $Target) = (split "/", $ENV{"PATH_INFO"})[1..2];
		WebInit ();
		}


	NewContext () if (!$ContextID || !($ContextID =~ /^[0-9]+$/));
	ReadContext ($ContextID);

	$Target =~ /^(.+)[_\?\&](.+)$/;
	$Target = $1 if $1;

   DebugPage          	() if $Target eq "Debug";
	LettingPage        	() if $Target eq "Letting" || !$Target;
	ContractPage		 	() if $Target eq "Contract";
	ItemPage		 		 	() if $Target eq "Item";
	ProfilePage			 	() if $Target eq "Profile";
	HistoricalPricePage	() if $Target eq "HistPrice";
	LettingPricePage	 	() if $Target eq "LettingPrice";
	ContractPricePage	 	() if $Target eq "ContractPrice";
	VendorPricePage	 	() if $Target eq "VendorPrice";
	ItemChoicePage		 	() if $Target eq "ItemChoice";
	SettingsPage		 	() if $Target eq "Settings";
	NotesPage		 		() if $Target eq "Notes";
	AddNotePage		 		() if $Target eq "AddNote";
	SourcePage		 		() if $Target eq "Source";

	DebugPage          	(); # the fallthrough condition

   exit (0);
#END

#############################################################################


sub SettingsPage
	{
	CommonHeader ("Settings");

	my $Action = $Context{From} || "Letting";

	print "<center>\n";
	print "<FORM METHOD=GET action=\"$Action\">\n";
	print "<table border=1>\n";
	print "<tr><td bgcolor=\"#eeeeee\"><b>Option</b></td><td bgcolor=\"#eeeeee\"><b>Value</b></td><td bgcolor=\"#eeeeee\"><b>Description</b></td></tr>\n";

	print "<tr><td align=right><b>Provider:    </b></td><td><INPUT type=text name=Provider    value=\"$Context{Provider}\"      size=12 maxlength=18></td><td>Provider: MSOLAP for olap services, MSDASQL or blank for ODBC</td></tr>\n";
	print "<tr><td align=right><b>Data Source: </b></td><td><INPUT type=text name=DataSource  value=\"$Context{DataSource}\"    size=12 maxlength=18></td><td>DataSource: CRAIGF or DSSNTS for olap, a DSN for odbc</td></tr>\n";
	print "<tr><td align=right><b>Catalog:     </b></td><td><INPUT type=text name=Catalog     value=\"$Context{Catalog}\"	    size=12 maxlength=18></td><td>Catalog: dss for olap, blank out for odbc     </td></tr>\n";
	print "<tr><td align=right><b>Cube: 		 </b></td><td><INPUT type=text name=Cube        value=\"$Context{Cube}\"		  	 size=12 maxlength=18></td><td>Cube: newprofile for olap, unused for odbc </td></tr>\n";
	print "<tr><td align=right><b>LIP Graph Item Spacing:</b></td><td><INPUT type=text name=ColSize value=\"$Context{ColSize}\" size=12 maxlength=18></td><td>Default is to autocalculate spacing			 </td></tr>\n";
	print "<tr><td align=right><b>LIP Graph Y Size:</b></td><td><INPUT type=text name=LIPYSize  value=\"$Context{LIPYSize}\"          size=12 maxlength=18></td><td>Vertical size of LIP graph in pixels			 </td></tr>\n";
	print "<tr><td align=right><b>LIP Graph Point Size:  </b></td><td><INPUT type=text name=LIPPointSize    value=\"$Context{LIPPointSize}\"   size=12 maxlength=18></td><td>Size of plot points					</td></tr>\n";
	print "<tr><td align=right><b>Qty/Price Graph X Size:</b></td><td><INPUT type=text name=PQXSize  	 value=\"$Context{PQXSize}\"     size=12 maxlength=18>		 </td><td>Horz size in pixels					</td></tr>\n";
	print "<tr><td align=right><b>Qty/Price Graph Y Size:</b></td><td><INPUT type=text name=PQYSize  	 value=\"$Context{PQYSize}\"     size=12 maxlength=18>		 </td><td>Vert size in pixels					</td></tr>\n";
	print "<tr><td align=right><b>Qty/Price Graph Point Size:  </b></td><td><INPUT type=text name=PQPointSize value=\"$Context{PQPointSize}\" size=12 maxlength=18> </td><td>Size of plot points					</td></tr>\n";

	print "</table><br>\n";
	print "<table border=1>\n";
	print "<tr><td bgcolor=\"#eeeeee\"><b> Option </b></td>".
	          "<td bgcolor=\"#eeeeee\"><b> &nbsp No &nbsp </b></td>".
	          "<td bgcolor=\"#eeeeee\"><b> &nbsp Yes &nbsp </b></td>".
	          "<td bgcolor=\"#eeeeee\"><b> Description </b></td></tr>\n";
	print "<tr><td align=right><b>Debug:	 </b></td><td><INPUT type=radio name=Debug 	  value=0 ".($Context{Debug}			  ?"":"checked")."></td><td><INPUT type=radio name=Debug 			value=1 ".($Context{Debug}			  ?"checked":"")."	></td><td>Display Debug Info in pages?</td></tr>\n";
	print "<tr><td align=right><b>UseVB:	 </b></td><td><INPUT type=radio name=UseVB 	  value=0 ".($Context{UseVB}			  ?"":"checked")."></td><td><INPUT type=radio name=UseVB 			value=1 ".($Context{UseVB}			  ?"checked":"")."	></td><td>Use VB Com object over PERL ADO where available?</td></tr>\n";
	print "<tr><td align=right><b>PreferMDX:</b></td><td><INPUT type=radio name=PreferMDX value=0 ".($Context{PreferMDX}	  	  ?"":"checked")."></td><td><INPUT type=radio name=PreferMDX 	   value=1 ".($Context{PreferMDX}	?"checked":"")."	></td><td>Prefer MDX over SQL where available?</td></tr>\n";

	print "<tr><td align=right><b>Qty/Price Log Quantity: </b></td><td><INPUT type=radio name=PQLogQuan  value=0 ".($Context{PQLogQuan}   ?"":"checked")."></td><td><INPUT type=radio name=PQLogQuan  value=1 ".($Context{PQLogQuan }	?"checked":"")."	></td><td>Use Log Quan for Best Fit Line  (no means linear)?</td></tr>\n";
	print "<tr><td align=right><b>Qty/Price Log Price:    </b></td><td><INPUT type=radio name=PQLogPrice value=0 ".($Context{PQLogPrice}  ?"":"checked")."></td><td><INPUT type=radio name=PQLogPrice value=1 ".($Context{PQLogPrice}	?"checked":"")."	></td><td>Use Log Price for Best Fit Line (no means linear) ?</td></tr>\n";

	print "</table>\n";
	print "<input type=\"submit\" value=\"Change\"><br>\n";
	print "</FORM>\n";
	print "</center>\n";
	CommonFooter ();
   exit (0);
	}


sub AddNotePage
	{
	my $From = $Context{From};

	my $NoteFile = "$ContextDir/$From.txt";
	my $TmpFile  = "$ContextDir/$From.tmp";
   my $Initials = $Context{Initials};
   my $Note		 = $Context{Note};
   $Note =~ tr/+/ /;
   $Note = uri_unescape ($Note);

	my @dt = localtime ();
	my $Time = sprintf ("%2.2d/%2.2d/%2.2d %2.2d:%2.2d", $dt[4]+1, $dt[3], $dt[5], $dt[2], $dt[1]);

	open (TMPFILE, ">$TmpFile") or MyDie ("Cannot open file $TmpFile for writing");
	print TMPFILE "$Time~$Initials~$Note\n";

	if (open (NOTEFILE, "<$NoteFile"))
		{
		while (<NOTEFILE>)
			{
			print TMPFILE $_;
			}
		close (NOTEFILE);
		}
	close (TMPFILE);
	unlink ($NoteFile);
	rename ($TmpFile, $NoteFile);

	NotesPage ();
	}


sub NotesPage
	{
	my ($From) = $Context{From};

	CommonHeader ("Notes for $From");

	print "<a href=\"$From\">Back to $From page</a><br>\n";

	print "<FORM METHOD=GET action=\"AddNote?From=$From\">\n";
	print "<center><TABLE>\n";
	print " <TR>\n";
	print " <TD bgcolor=\"#eeeeee\"><b>Date</b></TD>\n";
	print " <TD bgcolor=\"#eeeeee\" ALIGN=CENTER><b>Initials</b></TD>\n";
	print " <TD bgcolor=\"#eeeeee\"><b>Note</b></TD>\n";
	print " <TD bgcolor=\"#eeeeee\"></TD></TR>\n";
	print " <TR>\n";
	print " <TD></TD>\n";
	print " <TD><INPUT type=text name=Initials size=5 maxlength=32></TD>\n";
	print " <TD><INPUT type=text name=Note size=80 maxlength=160></TD>\n";
	print " <TD><input type=\"submit\" value=\"Add\"></td>\n";
	print " </TR>\n";

	my (@Data);
	my $NoteFile = "$ContextDir/$From.txt";
	if (open (NOTEFILE, "<$NoteFile"))
		{
		while (<NOTEFILE>)
			{
			@Data = split "~", $_;
			TableRow (\@Data, ["L<font size=1>~1</font>", "S", "S"]);
			}
		close (NOTEFILE);
		}
	print "</table>\n";
	print "<INPUT type=hidden name=From value=$From>\n";

	print "</center>\n";
	print "</form>\n";
	print "</body></html>\n";
	exit (0);
	}


sub SourcePage
	{
	my $Filename = $ENV{SCRIPT_FILENAME};

	MiniHead ();
	print "<PRE>\n";

	open (FILE, "<$Filename") || MyDie ("Cannot open $Filename");
	while (<FILE>)
		{
		print $_;
		}
	close FILE;

	print "</PRE>\n";
	print "</body></html>\n";
	exit (0);
	}


# Lists all lettings in the cube
# Make user Select a letting
#
sub LettingPage
	{
	CommonHeader ("Letting Selection");

	my $QueryOlapSQL = # SQL if using OLAP SQL 
			"Select distinct [lettings].[letting id], [lettings].[date] from $Context{Cube}";

	my $QueryODBCSQL = # SQL if using ODBC
#			"Select LettingID, LettingDate from Lettings";
			"Select distinct lettingdate from project";

	TableFromSQL (($Context{Provider} eq "MSOLAP" ? $QueryOlapSQL : $QueryODBCSQL),
	              ["L<a href=\"Contract?Letting=~1&LettingDate=~2\">~1</a>", "S"], 
	              ["Letting ID", "Letting Date"]);
	
	CommonFooter ($Context{Provider} eq "MSOLAP" ? "[PERL OLAP SQL]" : "[PERL ODBC SQL]");
	exit (0);	
	}




# Lists all contracts in the current letting
# Make user Select a contract
#
sub ContractPage		
	{
	my $Query;
		
	CommonHeader ("Contract Selection", 
	              "<font size=2><b>Letting ID: </b><a href=\"Letting\">$Context{Letting}</a> ($Context{LettingDate})</font>");

	my $QueryOlapMDX = # MDX if using OLAP MDX   # Yo, this MDX query does not filter on letting!
			"select NON EMPTY CrossJoin ({[contracts].[contract id].members}, \n".
	      "{[contractdesc].[Description].members}) on rows, \n".
			"{[amount]} on columns \n".
			"from $Context{Cube}\n";

	my $QueryODBCSQL = # SQL if using ODBC
##			" Select ContractID, ContractDesc\n".
##		   " from Contracts\n".
##		   " where LettingID = '$Context{Letting}'";
" select projectid, description from project";      


	if ($Context{Provider} eq "MSOLAP") # connected to OLAP database?
		{
		TableFromMDX ($QueryOlapMDX,
#						  ["L<a href=\"ContractChoice?Contract=~1\">~1</a>", "S", "X"], 
						  ["L<a href=\"Profile?Contract=~1\">~1</a>", "S", "X"], 
		 				  ["Contract ID", "Description"]);
		CommonFooter ("[PERL OLAP MDX]");
		}
	else # connected to ODBC database
		{
		TableFromSQL ($QueryODBCSQL,
#						  ["L<a href=\"ContractChoice?Contract=~1\">~1</a>", "S"], 
						  ["L<a href=\"Profile?Contract=~1\">~1</a>", "S"], 
		 				  ["Contract ID", "Description"]);
		CommonFooter ("[PERL ODBC SQL]");
		}
	exit (0);
	}


# Lists all items in the current contract
# Make user Select an item
#
sub ItemPage		 		
	{
	CommonHeader ("Item Selection", 
	              "<font size=2><b>Letting ID: </b><a href=\"Letting\">$Context{Letting}</a> ($Context{LettingDate})<br>".
                 "<b>Contract: </b><a href=\"Contract\">$Context{Contract}</a></font>");

	my $QueryOlapSQL = # SQL if using OLAP SQL 
			" Select distinct [Item Id]\n ".
			" from $Context{Cube}\n ".
			" where ([Contract Id] = '$Context{Contract}') ";

	my $QueryODBCSQL = # SQL if using ODBC
##			" Select ItemNumber, Quantity, Unit, ItemDesc\n ".
##			" from Items\n ".
##			" where ContractID = '$Context{Contract}'";
" select itemnumber, quantity, unit, description from item where projectid = '$Context{Contract}'";

	if ($Context{Provider} eq "MSOLAP") # connected to OLAP database?
		{
		TableFromSQL ($QueryOlapSQL,
		              ["L<a href=\"HistPrice?Item=~1\">~1</a>", "S", "S"],
						  ["Item #", "Unit", "&nbsp Description &nbsp"]);
		CommonFooter ("[PERL OLAP SQL]");
		}
	else
		{
		TableFromSQL ($QueryODBCSQL,
		              ["L<a href=\"HistPrice?Item=~1\">~1</a>", "N4", "S", "S"],
						  ["Item #", "Quantity", "Unit", "&nbsp Description &nbsp"]);
		CommonFooter ("[PERL ODBC SQL]");
		}
	exit (0);
	}


sub LettingPricePage	
	{
	my $HeaderInfo = 
		"<B><font size=+1>Letting Price Analysis</font></B><br>\n".
		"<B>Item:</B> <a href=\"Item\">$Context{Item}</a> - Item Description	<br>\n".
		"<B>Letting:</B> $Context{Letting}<br>\n".
		"For all contracts and bidders\n";
	my $Links1 =
		"<B>Further Analyses:</B><br>\n".
		"Prices for Letting $Context{Letting}<br>\n".
		"<a href=\"ContractPrice\"> Prices for Contract $Context{Contract}</a><br>\n".
		"<a href=\"VendorPrice\"  > Prices for Vendor $Context{Vendor}  </a><br>\n".
		"<a href=\"HistPrice\"	  > Historical Item Prices </a>\n";
	my $Links2 =
		"<B>To:</B><br>\n".
		"<a href=\"Profile?LIPGraph=1\" > LIP for Contract $Context{Contract}</a><br>\n".
		"<a href=\"Profile?LIPGraph=0\" > BidTabs for Contract $Context{Contract}</a><br>\n".
		"<a href=\"Letting\" > Change Letting									</a><br>\n".
		"<a href=\"Contract\"> Change Contract									</a>    \n";

	NewCommonHeader ("Letting Item Prices", $HeaderInfo, $Links1, $Links2, [40,30,30]);


	my $QueryOlapMDX = # MDX if using OLAP MDX
			" select {[Measures].[Quantity], [Measures].[Price]} on columns,\n".
			" NON EMPTY CrossJoin ({[contracts].[contract id].members}, {[vendors].[Vendor id].members})  on rows\n".
			" from $Context{Cube}\n".
			" where ([lettings].[All lettings].[$Context{Letting}], [items].[$Context{Item}])";

	my $QueryOlapSQL = # SQL if using OLAP SQL 
			" Select [Contract Id],[vendor id], vendors.Vendor, quantity, price\n".
			" from $Context{Cube}\n".
			" where ([Item Id] = '$Context{Item}') and ([letting id] = '$Context{Letting}')";

	my $QueryODBCSQL = # SQL if using ODBC
##			" Select ContractID, BidTabs.VendorID, VendorName, Quantity, UnitPrice\n".
##			" From BidTabs, Vendors\n".
##			" Where BidTabs.VendorID = Vendors.VendorID\n".
##			" and ItemNumber = '$Context{Item}'\n".
##			" and ContractID in (Select ContractID from Contracts where LettingID = '$Context{Letting}')".
##			" order by ContractID, BidTabs.VendorID";

			" select itembid.projectid, itembid.vendorid, vendor.name, item.quantity, itembid.unitprice\n".
			" from itembid, item, vendor, project\n".
##			" where project.lettingdate = '$Context{Letting}'\n".
			" where item.projectid = project.projectid\n".
			" and item.itemnumber = '$Context{Item}'\n".
			" and itembid.projectid = item.projectid\n".
			" and itembid.itemid = item.itemid\n".
			" and vendor.vendorid = itembid.vendorid";





	if ($Context{UseVB})
		{
		TableFromVBObject ("coffee.LPA", ["+S","S","S","S","S","N4","M"]);
		CommonFooter ("[VB]");
		}
	elsif ($Context{PreferMDX})
		{
		TableFromMDX ($QueryOlapMDX,
						  ["+S","S","R","R"], 
		 				  ["Contract ID", "Vendor ID", "Quantity", "Price"]);
		CommonFooter ("[PERL MDX]");
		}
	else # OLAP SQL or ODBC SQL
		{
		TableOrChartChoice ("LettingPrice");

		my $Query =	($Context{Provider} eq "MSOLAP" ? $QueryOlapSQL : $QueryODBCSQL);

		if ($Context{PQGraph}) # display data as a graph
			{
			my $PQData = LoadSQLTable ($Query, [5,4,2,1], "", "", "", "");
			my ($Name, $XSize, $YSize) = CreatePriceQuantityGraph ($PQData, "VendorPrice");
			ChartHTML ($XSize, $YSize);
			}
		else # display data in a table
			{
			TableFromSQL ($Query,
							  ["+S","S","S","N4","M"], 
			 				  ["Contract ID", "Vendor ID", "Vendor Name", "Quantity", "Price"]);
			}
		CommonFooter ($Context{Provider} eq "MSOLAP" ? "[PERL OLAP SQL]": "[PERL ODBC SQL]");
		}
	exit (0);
	}


sub ContractPricePage	
	{
	my $HeaderInfo = 
		"<B><font size=+1>Contract Price Analysis</font></B><br>\n".
		"<B>Item:</B> <a href=\"Item\">$Context{Item}</a> - Item Description	<br>\n".
		"<B>Contract:</B> $Context{Contract}<br>\n".
		"For all bidders\n";
	my $Links1 =
		"<B>Further Analyses:</B><br>\n".
		"<a href=\"LettingPrice\" > Prices for Letting $Context{Letting}  </a><br>\n".
		"Prices for Contract $Context{Contract}<br>\n".
		"<a href=\"VendorPrice\"  > Prices for Vendor $Context{Vendor}  </a><br>\n".
		"<a href=\"HistPrice\"	  > Historical Item Prices </a>\n";
	my $Links2 =
		"<B>To:</B><br>\n".
		"<a href=\"Profile?LIPGraph=1\" > LIP for Contract $Context{Contract}</a><br>\n".
		"<a href=\"Profile?LIPGraph=0\" > BidTabs for Contract $Context{Contract}</a><br>\n".
		"<a href=\"Letting\" > Change Letting									</a><br>\n".
		"<a href=\"Contract\"> Change Contract									</a>    \n";

	NewCommonHeader ("Contract Item Prices", $HeaderInfo, $Links1, $Links2, [40,30,30]);

	my $QueryOlapMDX = # MDX if using OLAP MDX
			" select {[Measures].[Price]} on columns,\n".
			" NON EMPTY {[vendors].[Vendor id].members} on rows\n".
			" from $Context{Cube}\n".
			" where ([contracts].[contract id].[$Context{Contract}], [items].[$Context{Item}])";

	my $QueryOlapSQL = # SQL if using OLAP SQL 
			" Select [vendor id], vendors.Vendor, price\n".
			" from $Context{Cube}\n".
			" where ([Item Id] = '$Context{Item}') and ([contract id] = '$Context{Contract}')";

	my $QueryODBCSQL = # SQL if using ODBC
##			" Select BidTabs.VendorID, VendorName, UnitPrice\n".
##			" From BidTabs, Vendors\n".
##			" Where BidTabs.VendorID = Vendors.VendorID\n".
##			" and ItemNumber = '$Context{Item}'\n".
##			" and ContractID = '$Context{Contract}'";

			" select itembid.vendorid, vendor.name, itembid.unitprice\n".
			" from itembid, item, vendor\n".
			" and itembid.projectid = '$Context{Contract}'\n".
			" and item.projectid = itembid.projectid\n".
			" and item.itemnumber = '$Context{Item}'\n".
			" and itembid.itemid = item.itemid\n".
			" and vendor.vendorid = itembid.vendorid";


	if ($Context{UseVB})
		{
		TableFromVBObject ("coffee.CPA", ["+S","S","R","R"]);
		CommonFooter ("[VB]");
		}
	elsif ($Context{PreferMDX})
		{
		TableFromMDX ($QueryOlapMDX, ["S","R"], ["Vendor ID", "Price"]);
		CommonFooter ("[PERL MDX]");
		}
	else
		{
		my $Query =	($Context{Provider} eq "MSOLAP" ? $QueryOlapSQL : $QueryODBCSQL);
		TableFromSQL ($Query,
						  ["S","S","M"], 
		 				  ["Vendor ID", "Vendor Name", "Price"]);
		CommonFooter ($Context{Provider} eq "MSOLAP" ? "[PERL OLAP SQL]": "[PERL ODBC SQL]");
		}
	exit (0);
	}


sub VendorPricePage
	{
	my ($VenName) = SingleRowSQL ($Context{Provider} eq "MSOLAP" ?
	                               "select distinct [vendors].[Vendor] from $Context{Cube} \n".
					    					 "where ([vendors].[Vendor Id] = '$Context{Vendor}')"
					    					 :
											 "Select VendorName from Vendors where VendorID = '$Context{Vendor}'"
   				    					 );

	my $HeaderInfo = 
		"<B><font size=+1>Vendor Historical Price Analysis</font></B>			<br>\n".
		"<B>Item:</B> <a href=\"Item\">$Context{Item}</a> - Item Description	<br>\n".
		"<B>Vendor:</B> $Context{Vendor} - $VenName<br>\n".
		"For past 6 months [<a href=\"DateRange\">Change</a>]\n";
	my $Links1 =
		"<B>Further Analyses:</B><br>\n".
		"<a href=\"LettingPrice\" > Prices for Letting $Context{Letting}  </a><br>\n".
		"<a href=\"ContractPrice\"> Prices for Contract $Context{Contract}</a><br>\n".
		"Prices for Vendor $Context{Vendor}<br>\n".
		"<a href=\"HistPrice\"	  > Historical Item Prices </a>\n";
	my $Links2 =
		"<B>To:</B><br>\n".
		"<a href=\"Profile?LIPGraph=1\" > LIP for Contract $Context{Contract}</a><br>\n".
		"<a href=\"Profile?LIPGraph=0\" > BidTabs for Contract $Context{Contract}</a><br>\n".
		"<a href=\"Letting\" > Change Letting									</a><br>\n".
		"<a href=\"Contract\"> Change Contract									</a>    \n";

	NewCommonHeader ("Vendor Item Prices", $HeaderInfo, $Links1, $Links2, [40,30,30]);

	my $QueryOlapMDX = # MDX if using OLAP MDX
			" select {[Measures].[Quantity], [Measures].[Price]} on columns,\n".
			" NON EMPTY CrossJoin ({[lettings].[letting id].members}, {[contracts].[contract id].members}) on rows\n".
			" from $Context{Cube}\n ".
			" where ([vendors].[All vendors].[$Context{Vendor}], [items].[$Context{Item}])";

	my $QueryOlapSQL = # SQL if using OLAP SQL 
			" Select [letting id], [Contract Id], quantity, price\n".
			" from $Context{Cube}\n".
			" where ([Item Id] = '$Context{Item}') and ([Vendor Id] = '$Context{Vendor}')";

	my $QueryODBCSQL = # SQL if using ODBC
			" Select LettingID, BidTabs.ContractID, Quantity, UnitPrice\n".
			" From BidTabs, Contracts\n".
			" Where BidTabs.ContractID = Contracts.ContractID\n".
			" and ItemNumber = '$Context{Item}'\n".
			" and VendorID = '$Context{Vendor}'";

##select lettingdate, projectid, quantity, unitprice
##from itembid, item, project
##where item.itemnumber = '$Context{Item}'
##and itembid.vendorid =
##


	if (!$Context{Vendor})
		{
		print "<font size = +2> You must first select a vendor from the <a href=\"Profile?LIPGraph=1\" >Line Item Profile</a>".
				" or <a href=\"Profile?LIPGraph=0\">Bid Tab</a> page.</font>";
		CommonFooter ($Context{UseVB} ? "[VB]":"[PERL]");
		exit (0);
		}

	if ($Context{UseVB})
		{
		TableFromVBObject ("coffee.VPA", ["S","S","S","S","S","S"]);
		CommonFooter ("[VB]");
		}
	elsif ($Context{PreferMDX})
		{
		TableFromMDX ($QueryOlapMDX,
						  ["+S", "S", "R", "R"],
		 				  ["Letting ID", "Contract ID", "Quantity", "Price"]);
		CommonFooter ("[PERL MDX]");
		}
	else
		{
		my $Query =	($Context{Provider} eq "MSOLAP" ? $QueryOlapSQL : $QueryODBCSQL);

		TableOrChartChoice ("VendorPrice");

		if ($Context{PQGraph}) # display data as a graph
			{
			my $PQData = LoadSQLTable ($Query, [4,3,0,2], "", "", "$Context{Vendor}", "");
			my ($Name, $XSize, $YSize) = CreatePriceQuantityGraph ($PQData, "HistPrice");
			ChartHTML ($XSize, $YSize);
			}
		else # display data in a table
			{
			TableFromSQL ($Query,
						  ["+S", "S", "N4", "M"], 
		 				  ["Letting ID", "Contract ID", "Quantity", "Price"]);
			}
		CommonFooter ($Context{Provider} eq "MSOLAP" ? "[PERL OLAP SQL]": "[PERL ODBC SQL]");
		}
	exit (0);
	}



sub HistoricalPricePage
	{
	my $HeaderInfo = 
		"<B><font size=+1>Historical Price Analysis</font></B><br>\n".
		"<B>Item:</B> <a href=\"Item\">$Context{Item}</a> - Item Description<br>\n".
		"For past 6 months [<a href=\"DateRange\"><font size=2>Change</font></a>]<br>\n".
		"For all contracts and bidders\n";
	my $Links1 =
		"<B>Further Analyses:</B><br>\n".
		"<a href=\"LettingPrice\" > Prices for Letting $Context{Letting}  </a><br>\n".
		"<a href=\"ContractPrice\"> Prices for Contract $Context{Contract}</a><br>\n".
		"<a href=\"VendorPrice\"  > Prices for Vendor $Context{Vendor}  </a><br>\n".
		"Historical Item Prices\n";
	my $Links2 =
		"<B>To:</B><br>\n".
		"<a href=\"Profile?LIPGraph=1\" > LIP for Contract $Context{Contract}</a><br>\n".
		"<a href=\"Profile?LIPGraph=0\" > BidTabs for Contract $Context{Contract}</a><br>\n".
		"<a href=\"Letting\" > Change Letting									</a><br>\n".
		"<a href=\"Contract\"> Change Contract									</a>    \n";

	NewCommonHeader ("Vendor Item Prices", $HeaderInfo, $Links1, $Links2, [40,30,30]);

	my $QueryOlapMDX = # MDX if using OLAP MDX
			" with set [t1] as 'CrossJoin({[all lettings].children}, {[all contracts].children})'\n".
			"        set [t2] as 'CrossJoin({[all].children}, {[all vendors].children})' \n".
			" select NON EMPTY CrossJoin({[t1]}, {Order([t2], price, asc)}) on rows, \n".
			" {quantity, price} on columns\n".
			" from $Context{Cube}\n".
			" where [items].[$Context{Item}]";

	my $QueryOlapSQL = # SQL if using OLAP SQL 
			" Select [letting id],[Contract Id],[vendor id],[vendors].[Vendor], quantity, price\n".
			" from $Context{Cube}\n".
			" where ([Item Id] = '$Context{Item}')";

	my $QueryODBCSQL = # SQL if using ODBC
			" Select LettingID, BidTabs.ContractID, BidTabs.VendorID, VendorName, Quantity, UnitPrice\n".
			" From BidTabs, Contracts, Vendors\n".
			" Where BidTabs.ContractID = Contracts.ContractID\n".
			" and BidTabs.VendorID = Vendors.VendorID\n".
			" and ItemNumber = '$Context{Item}'";

	if ($Context{UseVB})
		{
		TableFromVBObject ("coffee.HPA", ["+S","+S","S","S","S","N4","M"]);
		CommonFooter ("[VB]");
		}
	elsif ($Context{PreferMDX})
		{
		TableFromMDX ($QueryOlapMDX, 
		              ["+S","+S","S","S","R","R"], 
						  ["Letting ID", "Contract ID", "Vendor Name", "Vendor ID", "Quantity", "Price"]);
		CommonFooter ("[PERL MDX]");
		}
	else
		{
		my $Query =	($Context{Provider} eq "MSOLAP" ? $QueryOlapSQL : $QueryODBCSQL);

		TableOrChartChoice ("HistPrice");

		if ($Context{PQGraph}) # display data as a graph
			{
			my $PQData = LoadSQLTable ($Query, [6,5,3,2], "", "", "", "");
			my ($Name, $XSize, $YSize) = CreatePriceQuantityGraph ($PQData, "VendorPrice");
			ChartHTML ($XSize, $YSize);
			}
		else # display data in a table
			{
			TableFromSQL ($Query, 
			              ["+S","+S","S","S","N4","M"], 
			              ["Letting ID", "Contract ID", "Vendor ID", "Vendor Name", "Quantity", "Price"] );
			}
		CommonFooter ($Context{Provider} eq "MSOLAP" ? "[PERL OLAP SQL]": "[PERL ODBC SQL]");
		}
	exit (0);
	}


sub ItemChoicePage		
	{
	CommonHeader ("Item Analysis Choice", "<b>Item:</b><a href=\"Item\">$Context{Item}</a>");

	print "<br>\n";
	print "<h2>Choose Item Analysis Display</h2>\n";
	print "<li><a href=\"HistPrice\"	 	> Historical Price Analysis</a>\n";
	print "<li><a href=\"LettingPrice\"	> Letting Price Analysis	</a>\n";
	print "<li><a href=\"ContractPrice\"> Contract Price Analysis	</a>\n";
	print "<li><a href=\"VendorPrice\"	> Vendor Price Analysis		</a>\n";

	CommonFooter ();
	exit (0);
	}


sub ProfilePage			
	{
   srand (time () ^ ($$ + ($$ << 15)));
   my $Rand = sprintf ("%5.5d", rand (32000));

	CommonHeader ("Line Item Profiles and BidTabs", 
	              "<font size=2><b>Letting ID: </b><a href=\"Letting\">$Context{Letting}</a> ($Context{LettingDate})<br>".
                 "<b>Contract: </b><a href=\"Contract\">$Context{Contract}</a> ".
                 "&nbsp &nbsp &nbsp &nbsp <-<a href=\"\">Prev Contract</a> | <a href=\"\">Next Contract</a>-> ".
                 "</font>");

	my $Data = LoadBidtabData ();

	ConvertBaseBid ($Data);

	DumpProfileData ($Data) if ($Context{Debug});
	ProfileVendors ($Data);
	LIPTableOrChartChoice ("Profile");

	if ($Context{LIPGraph}) # display data as a graph
		{
		my ($Name, $XSize, $YSize) = CreateProfileGraph ($Data);

		print "<center>\n";
		print "<img src=\"/$SubDir/LIP_$ContextID.png?$Rand\" BORDER=0 USEMAP=\"#profilemap\" WIDTH=$XSize HEIGHT=$YSize>\n";
		print "</center>\n";
		print "<MAP NAME=\"profilemap\">\n";
		map {print "$_\n"} @ImageMapEntries;
		print "</MAP>\n";
		}
	else
		{
		CreateBidTabTable ($Data);
		print "(%Est - The Vendors Item Bid divided by the Baseline Total Bid expressed as a percent)<br>\n";
		print "(%Dev - The difference between Vendors Item Bid and the Baseline Bidders Item Bid / Baseline Bid Total)<br>\n";
		}
	CommonFooter ();
	exit (0);
	}


sub CreateBidTabTable 
	{
	my ($Data) = @_;

	# Table Header
	BidTabTableHeader ();

	my @VenList = GetVendorList ($Data);

	my ($Item, $VendorID);
	foreach $Item (@$Data)
		{
		BidTabTableItemRow ($Item->{ID}, $Item->{Desc});

		my $VendorIdx = 0;
		foreach $VendorID (@VenList)
			{
			$VendorIdx++;
			my $ref = sprintf("V%2.2d", $VendorIdx);
			next if (!$Context{$ref});	 # user excluded this vendor

			my $Vendor = $Item->{Bids}->{$VendorID};
			my $VendorName = $Data->[0]->{Bids}->{$VendorID}->{Name};

			BidTabTableBidRow ($Item->{ID}, $VendorID, 
			                   $VendorName, 			$Vendor->{Price}, 
									 $Item->{Quantity},	$Vendor->{Amount},
									 $Vendor->{PctEst},	$Vendor->{Dev});
			}
		}
	BidTabTableEnd ();
	}

sub BidTabTableHeader
	{
	print "<CENTER>\n";

	BorderStart ("98%");

	print "<TABLE BORDER=0 WIDTH=\"100%\" CELLPADDING=3 CELLSPACING=3>\n";
	print "<TR><TD COLSPAN=10></TD></TR>\n";
	print "<TR ALIGN=CENTER BGCOLOR=\"$TableHeaderColor\">\n";
	print "	<TD><B>Line</B></TD>\n";
	print "	<TD><B>Item</B></TD>\n";
	print "	<TD ALIGN=CENTER COLSPAN=2><B>Description</B></TD>\n";
	print "	<TD COLSPAN=7 ><B>Bidder Data</B></TD>\n";
	print "</TR>\n";
	}
sub BidTabTableItemRow
	{
	my ($ItemID, $ItemDesc) = @_;

	my $ItemLink = "<a href=\"HistPrice?Item=$ItemID\">$ItemID</a>";

	print "<TR BGCOLOR=\"$TableAltRowColor\">\n";
	print "	<TD WIDTH=\"4%\"  >????</TD>   \n";
	print "	<TD WIDTH=\"7%\" VALIGN=TOP ALIGN=LEFT NOWRAP>$ItemLink</TD>\n";
	print "	<TD WIDTH=\"58%\"VALIGN=TOP ALIGN=LEFT COLSPAN=2>&nbsp;$ItemDesc</TD>\n";
	print "	<TD WIDTH=\"7%\" ALIGN=CENTER>BidPrice</TD>\n";
	print "	<TD WIDTH=\"7%\" ALIGN=CENTER COLSPAN=3>Quantity</TD>\n";
	print "	<TD WIDTH=\"7%\" ALIGN=CENTER>ExtAmount</TD>\n";
	print "	<TD WIDTH=\"5%\" ALIGN=CENTER><nobr>%Est</nobr></TD>\n";
	print "	<TD WIDTH=\"5%\" ALIGN=CENTER><nobr>%Dev</nobr></TD>\n";
	print "</TR>\n";
	}
sub BidTabTableBidRow
	{
	my ($ItemID, $VendorID, $VendorName, $UnitPrice, $Quantity, $Ext, $PctEst, $Dev) = @_;

	my $VenLink = "<a href=\"VendorPrice?Item=$ItemID&Vendor=$VendorID\">$VendorName</a>";

	print "<TR ALIGN=RIGHT>\n";
	print "	<TD WIDTH=\"93%\" COLSPAN=4>$VenLink</TD>\n";
	print "	<TD WIDTH=\"1%\" >". &MoneyFormat ($UnitPrice) ."&nbsp;</TD>\n";
	print "	<TD WIDTH=\"1%\" >". &NumberFormat ($Quantity, "N3"). "</TD>\n";
	print "	<TD WIDTH=\"1%\" ALIGN=LEFT>LS</TD>\n";
	print "	<TD WIDTH=\"1%\" >&nbsp;</TD>\n";
	print "	<TD WIDTH=\"1%\" >". &MoneyFormat ($Ext) ."</TD>\n";
	print "	<TD WIDTH=\"1%\" >". &NumberFormat ($PctEst, "N2") ."</TD>\n";
	print "	<TD WIDTH=\"1%\" ><nobr>". &NumberFormat ($Dev,    "N2") ."</nobr></TD>\n";
	print "</TR>\n";
	}

sub BidTabTableEnd
	{
	print "</TABLE>\n";
	BorderEnd ();
	print "</CENTER>\n";
	}


sub ProfileVendors
	{
	my ($Data) = @_;
	my ($ID, $i, $VendorID, $chk);

	print "<center>\n";
	print "<FORM METHOD=GET>\n";
	
	TableStart("98%");
	TableHeader (["Show?", "Baseline", "Vendor ID", "Vendor Name", "&nbsp;Bid Total&nbsp;", "Display Total"]);
	$i = 0;

	my $BaselineVendor = GetBaselineVendorID ($Data);

	foreach $VendorID (GetVendorList ($Data))
		{
		my $VendorName   = $Data->[0]->{Bids}->{$VendorID}->{Name};
		my $BidTotal	  = $Data->[0]->{Bids}->{$VendorID}->{BidTotal};
		my $DisplayTotal = $Data->[0]->{Bids}->{$VendorID}->{DisplayTotal};

		my $ref = sprintf ("V%2.2d", $i++ +1);
		my $chk = ($Context{$ref} ? "checked" : "");
		my $sel = ($VendorID eq $BaselineVendor ? "checked" : "");

		print "<TR " . (!($i % 2) ? "bgcolor=\"$TableAltRowColor\"" : "") . ">\n";
		print " <td><input type=checkbox $chk name=\"$ref\" value=\"1\">";
		print " <img src=\"/$SubDir/$ref.jpg\"></td>";
		print " <td align=center><input type=radio name=\"BaseBid\" value=\"$VendorID\" $sel></td>";
		print " <td>&nbsp;$VendorID</td>";
		print " <td>&nbsp;$VendorName</td>";
		print " <td align=right>&nbsp;&nbsp;". &MoneyFormat ($BidTotal)."</td>";
		print " <td align=right>&nbsp;&nbsp;". &MoneyFormat ($DisplayTotal)."</td>";
		print "</TR>\n";
		}
	print "<TR>\n";
	print "<TD align=center colspan=5>\n";
	print " <B>Minimum Deviation: &nbsp;</B><INPUT type=text name=LIPMinDev value=\"$Context{LIPMinDev}\"size=4 maxlength=8>\n";
#	print " <B>Grid Height:</B><INPUT type=text name=YSize value=\"$Context{YSize}\"size=4 maxlength=8>\n";
	print " &nbsp; <input type=\"submit\" value=\"Change\"></td>\n";
	print "</TR>\n";
	print "<input type=hidden name=\"c00\" value=\"1\">\n";

	TableEnd ();
	print "</FORM>\n";
	print "</center>\n";
	}


#############################################################################
#
# Item Profile Graph generation
#
#############################################################################

#
# Creates a LIP graph in PNG format
#
sub CreateProfileGraph
	{
	my ($Data) = @_;
	my ($i, $ColorSpec);

	LIPCalcMetrics ($Data);

	my $Image = new GD::Image($Context{LIPXSize},$Context{LIPYSize});

	# colors used in the graph frame
	my $Linen = $Image->colorAllocate(255,246,236);	# actually light linen
	my $white = $Image->colorAllocate(255,255,255);
	my $gray  = $Image->colorAllocate(192,192,192);
	my $black = $Image->colorAllocate(0,0,0);

	# colors used when plotting
	my @Colors;
	foreach $ColorSpec (@ColorSet)
		{
		push @Colors, $Image->colorAllocate (@$ColorSpec);
		}
	$Image->transparent($Linen);

	# show colors used if in debug mode
	if ($Context{Debug})
		{
		$Context{LIPLeftEdge}+= 15;
		for ($i=0; $i<17; $i++)
			{
	      map $Image->arc(10,$i*15+10, $_, $_, 0, 360, $Colors[$i]), (10,8,6,4,2);
			}
		}
	DisplayGraphTemplate ($Image, $Context{ItemCount}, $black, $gray);
	DisplayXLabels ($Image, $Data, $black);
	DisplayYLabels ($Image, $Data, $black, $gray);
	PlotBids ($Image, \@Colors, $Data);

	my @Bounds = $Image->getBounds ();

	my $ImageFile = "$DataDir/LIP_$ContextID.png";
	open (PIC, ">$ImageFile") or MyDie ("Cannot create image file $ImageFile");
	binmode PIC;
	print PIC $Image->png();
	close PIC;
	return ($ImageFile, @Bounds);
	}


#
# Calculates graph metrics that are needed to create the LIP graph
#
# calculates the following %Context values:
#	  ItemCount	 - The number of items to plot on the graph
#	  YMin		 -	Minimum Data Value (Y Dimension)
#	  YMax		 -	Maximum Data Value (Y Dimension)
#	  YRange		 -	Range of values (Y Dimension)
#	  XGridMin	 - Pixel X location of left edge of graph area
#	  YGridMin	 - Pixel Y location of minimum Y Value
#	  XGridMax	 -	Pixel X location of right edge of graph
#	  YGridMax	 -	Pixel Y location of maximum Y Value
#	  XGridSize	 -	X Size of grid display area in pixels
#	  YGridSize	 -	Y Size of grid display area in pixels
#	  YGridCtr	 -	Center Y location of grid display area in pixels
#	  LIPXSize	 -	X Size of entire Graph image (not just X grid display)
#	  ColSize	 -	Space between columns
#	  Font		 -	Font to use for the labels
#	  HalfXLabel -	1/2 the size of an average item# label in pixels in X dim
#
# Looks at:
#	  FixedColSize	 -	Used for ColSize if non zero
#	  CenterBaseBid -	Used to center base bid in graph if non zero
#
sub LIPCalcMetrics
	{
	my ($Data) = @_;
	my ($i, $Key, $Item, $Val, $VendorID, $VendorIdx, @VenList);

	$Context{ItemCount} = @$Data;
	$Context{ItemCount} = $Context{LIPMaxItems} if $Context{ItemCount} > $Context{LIPMaxItems};

	$i = $Context{YMin} = $Context{YMax} = 0;

	@VenList = GetVendorList ($Data);

	foreach $Item (@$Data)
		{
		next if ($i++ > $Context{LIPMaxItems});

		$VendorIdx = 0;
		foreach $VendorID (@VenList)
			{
			$VendorIdx++;
			my $ref = sprintf ("V%2.2d", $VendorIdx); #	explicitly 
			next if (!$Context{$ref});			   		#	excluded
																	#
			$Val = $Item->{Bids}->{$VendorID}->{Dev};	#	Get the deviation for this item bid

			$Context{YMin} = ($Context{YMin} < $Val ? $Context{YMin} : $Val);
			$Context{YMax} = ($Context{YMax} > $Val ? $Context{YMax} : $Val);
			}
		}
	if ($Context{YMax} == 0 && $Context{YMin} == 0)
		{
		$Context{YMin} = "-5.00"; # in case there are no vendors
		$Context{YMax} =  "5.00";
		}
	$Context{YGridMin}  = $Context{LIPTopEdge};
	$Context{YGridMax}  = $Context{LIPYSize} - $Context{LIPBottomEdge};
	$Context{YGridSize} = $Context{YGridMax} - $Context{YGridMin} + 1;


	# select item spacing and font based on the # of items to display in graph
	$Context{Font} 		 = gdMediumBoldFont;
	$Context{HalfXLabel}  = 30;
	$Context{ColSize} = 70;
	$Context{ColSize} = 60  if ($Context{ItemCount} == 9);
	$Context{ColSize} = 50  if ($Context{ItemCount} == 10);
	$Context{ColSize} = 47  if ($Context{ItemCount} == 11);
	$Context{ColSize} = 45  if ($Context{ItemCount} == 12);
	$Context{ColSize} = 40  if ($Context{ItemCount} >  12);

	if ($Context{ItemCount} > 15)
		{
		$Context{ColSize} 	= 30;
		$Context{Font} 		= gdTinyFont;
		$Context{HalfXLabel} = 22;
		}
	$Context{ColSize} = $Context{FixedColSize} if ($Context{FixedColSize});

	$Context{LIPXSize} = ($Context{ItemCount}+1) * $Context{ColSize} + 
	                      $Context{LIPLeftEdge} + $Context{LIPRightEdge};

	$Context{XGridMin}  = $Context{LIPLeftEdge};
	$Context{XGridMax}  = $Context{LIPXSize} - $Context{LIPRightEdge};
	$Context{XGridSize} = $Context{XGridMax} - $Context{XGridMin} + 1;

	$Context{YRange}  = $Context{YMax} - $Context{YMin};
	}



# Draws the general appearance of the graph without
# the protted points or labels
#
sub DisplayGraphTemplate
	{
	my ($Image, $Items, $Color1, $Color2) = @_;
	my $i;

	#draw frame
	for ($i=0; $i<2; $i++)
		{
	   $Image->line ($Context{XGridMin}-$i, $Context{YGridMin}, 
	                 $Context{XGridMin}-$i, $Context{YGridMax}, $Color1); # vert
	   $Image->line ($Context{XGridMin}-1, $Context{YGridMax}+$i, 
	                 $Context{XGridMax}, $Context{YGridMax}+$i, $Color1); #horz
		}
	$Image->line ($Context{XGridMin}, YPos(0), $Context{XGridMax}, YPos(0), $Color2);

	#draw vertical item col locations
	for ($i=0; $i<$Items; $i++)
		{
		my $XPos = $Context{XGridMin}+($i+1) * $Context{ColSize};
		$Image->line ($XPos, $Context{YGridMin}, $XPos, $Context{YGridMax}, $Color2);
		my $StaggeredLabelOffset = $i % 2 * 2 * $Context{HalfYLabel};
		$Image->line ($XPos, $Context{YGridMax}, $XPos, $Context{YGridMax}+ 5 + $StaggeredLabelOffset, $Color1);
		}
	$Image->line ($Context{XGridMax}, $Context{YGridMin}, 
	              $Context{XGridMax}, $Context{YGridMax}-1, $Color2);
	}


# Draws the Horiz labels
# Adds an ImageMap entry for each
#
sub DisplayXLabels
	{
	my ($Image, $Data, $Color) = @_;
	my ($i, $Str, $Quan, $Unit, $Desc);

	my $YLabel = $Context{YGridMax} + 8;
	for ($i=0; $Data->[$i]; $i++)
		{
		next if ($i >= $Context{LIPMaxItems});

		my $XPos = $Context{XGridMin}+ ($i+1) * $Context{ColSize} - $Context{HalfXLabel};
		my $StaggeredLabelOffset = $i % 2 * 2 * $Context{HalfYLabel};
		my $YPos = $YLabel + $StaggeredLabelOffset;

		$Str 	   = $Data->[$i]->{ID};
		$Quan    = $Data->[$i]->{Quantity};
		$Unit    = $Data->[$i]->{Unit};
		$Desc    = $Data->[$i]->{Desc};
		$Image->string($Context{Font}, $XPos, $YPos, $Str, $Color);

		AddMapRect ($XPos,$YPos,$XPos+$Context{HalfXLabel}*2,$YPos+$Context{HalfYLabel}*2-1,
		            "ContractPrice?Item=$Str", "Quantity: $Quan\n".
		                                       "Unit: $Unit\n".
		                                       "Description: $Desc");
		}
	}

#
#
#
sub DisplayYLabels
	{
	my ($Image, $Data, $Color1, $Color2) = @_;
	my ($i, $Last, $End);

	my $LE 	 = $Context{XGridMin};
	my $RE 	 = $Context{XGridMax};
	my $HY 	 = $Context{HalfYLabel};
	my $YPMin = YPos ($Context{YMin});
	my $YPMax = YPos ($Context{YMax});

	my $BaseLine = GetBaselineVendorID ($Data);
	my $OnePct = $Data->[0]->{Bids}->{$BaseLine}->{BidTotal} / 100;

	$Image->line ($LE-5, $YPMin, $LE, $YPMin, $Color1);
	if (SufficientGap (0, $Context{YMin}))
		{
		my $Val = MoneyFormat ($OnePct * $Context{YMin});
		$Image->string(gdMediumBoldFont, $LE-57, $YPMin-$HY, $Context{YMin}."%", $Color1);
		$Image->string(gdMediumBoldFont, $RE+3,  $YPMin-$HY, $Val, $Color1) if ($Context{LIPPrices});
		}

	$Image->line ($LE-5, YPos(0), $LE, YPos(0), $Color1);
	my $XOffset = (length $BaseLine) * 7 + 9;
	$Image->string(gdMediumBoldFont, $LE-$XOffset, YPos(0)-$HY, $BaseLine, $Color1);
	$Image->string(gdMediumBoldFont, $RE+3,        YPos(0)-$HY, "\$0.00", $Color1);

	$Image->line ($LE-5, $YPMax, $LE, $YPMax, $Color1);
	if (SufficientGap ($Context{YMax}, 0))
		{
		my $Val = MoneyFormat ($OnePct * $Context{YMax});
		$Image->string(gdMediumBoldFont, $LE-50, $YPMax-$HY, $Context{YMax}."%", $Color1);
		$Image->string(gdMediumBoldFont, $RE+3,  $YPMax-$HY, $Val, $Color1) if ($Context{LIPPrices});
		}

	# draw horizontal markers
	for ($i = int $Context{YMin}; $i <= int $Context{YMax}; $i++)
		{
		next if (!$i);
		$Image->line ($LE+1, YPos($i), $Context{XGridMax}, YPos($i), $Color2);
		$Image->line ($LE-4, YPos($i), $LE-2, YPos($i), $Color1);
		}
	$Image->line ($LE+1, $YPMax, $Context{XGridMax}, $YPMax, $Color2);

	# draw labels for grid markers
	for ($Last = 0, $End = $Context{YMax}, $i = 1; $i <= int $End; $i++)
		{
		if (SufficientGap ($i, $Last) && SufficientGap ($i, $End))
			{
			my $Val = MoneyFormat ($OnePct * $i);
			$Image->string(gdMediumBoldFont, $LE-25, YPos($i)-$HY, $i."%", $Color1);
			$Image->string(gdMediumBoldFont, $RE+3, YPos($i)-$HY, $Val, $Color1) if ($Context{LIPPrices});
			$Last = $i;
			}
		}
	for ($Last = 0, $End  = $Context{YMin}, $i = -1; $i >= int $End; $i--)
		{
		if (SufficientGap ($i, $Last) && SufficientGap ($i, $End))
			{
			my $Val = MoneyFormat ($OnePct * $i);
			$Image->string(gdMediumBoldFont, $LE-38, YPos($i)-$HY, $i."%", $Color1);
			$Image->string(gdMediumBoldFont, $RE+3, YPos($i)-$HY, $Val, $Color1) if ($Context{LIPPrices});
			$Last = $i;
			}
		}
	}

#	Draws the points and lines
# Adds an ImageMap entry for each plotted point
#
sub PlotBids
	{
	my ($Image, $Colors, $Data) = @_;
	my ($Item, $VendorID, $VendorIdx, $Dia, $Color, $XLast, $YLast, $ref, $ItemIndex);

	$Dia = $Context{LIPPointSize};

	my @VenList = GetVendorList ($Data);
	$VendorIdx = scalar @VenList - 1;

	for ($VendorIdx = @VenList; $VendorIdx; $VendorIdx--)
		{
		$VendorID = $VenList[$VendorIdx-1];

		$ref = sprintf("V%2.2d",$VendorIdx);
		next if (!$Context{$ref});					 # user excluded this vendor

		$Color = $$Colors[$VendorIdx-1];
		$XLast = 0;
		$YLast = 0;

		$ItemIndex = 0;
		foreach $Item (@$Data)
			{
			next if ($ItemIndex >= $Context{LIPMaxItems});

			my $Vendor 	= $Item->{Bids}->{$VendorID};
			my $Val   	= $Vendor->{Dev};
			my $Price 	= $Vendor->{Price};
			my $Quantity= $Vendor->{Quantity};
			my $Amount	= $Vendor->{Amount};
		   my $x = $Context{XGridMin}+($ItemIndex+1) * $Context{ColSize};
		   my $y = YPos($Val);

         map $Image->arc($x,$y,$_,$_, 0, 360, $Color), ($Dia,$Dia-2,$Dia-4);

#			AddMapCircle ($x,$y, $Dia, "ItemChoice?Vendor=$VendorID&Item=$Item->{ID}", 
			AddMapCircle ($x,$y, $Dia, "VendorPrice?Vendor=$VendorID&Item=$Item->{ID}", 
			                           "Vendor: $VendorID\n".
			                           "Unit Price: $Price\n".
			                           "Quantity: $Quantity\n".
			                           "Amount: \$$Amount\n".
			                           "Deviation: $Val%");

			$Image->line($XLast, $YLast, $x, $y, $Color) if ($XLast);
			($XLast, $YLast) = ($x, $y);
			$ItemIndex++;
			}
		}
	}

#
# Make the selected vendor the zero deviation value
# Convert the deviations of the other vendors to be
# relative to the selected vendor
#
sub ConvertBaseBid
	{
	my ($Data) = @_;
	my ($BaseBid, $Item, $VendorID, $Vendor, $NewDeviation, $Tmp);

	$BaseBid = $Context{BaseBid};
	if (!$BaseBid || !$Data->[0]->{Bids}->{$BaseBid} || $BaseBid eq "Estimate")
		{
		$Context{BaseBid} = "Estimate";
		return;
		}
	my @VenList = GetVendorList ($Data);

	my $BaseBidTotal = $Data->[0]->{Bids}->{$BaseBid}->{BidTotal};
	foreach $Item (@$Data)
		{
		my $BaseVendor = $Item->{Bids}->{$BaseBid};

		foreach $VendorID (@VenList)
			{
			$Vendor = $Item->{Bids}->{$VendorID};
			my $Tmp = (($Vendor->{Amount} - $BaseVendor->{Amount})/$BaseBidTotal) * 100;

			$Vendor->{Dev} = sprintf ("%.2f", $Tmp);
			}
		}
	}

#############################################################################
#
# Item vs Quantity Graph Generation
#
#############################################################################
#
# The Price -vs- Quantity graph was added agter the Item profile graph was 
# already created.  Several support functions written for the IP graph should
# be re-written to be generic so that all graphs can use them.  For now this
# graph is just kind-of hacked in.
#
sub CreatePriceQuantityGraph
	{
	my ($PQData, $LinkName) = @_;
	my ($i, $ColorSpec);

	PQCalcMetrics ($PQData);

	my $Image = new GD::Image($Context{PQXSize},$Context{PQYSize});

	# colors used in the graph frame
	my $Linen = $Image->colorAllocate (250,240,230);
	my $white = $Image->colorAllocate (255,255,255);
	my $gray  = $Image->colorAllocate (192,192,192);
	my $black = $Image->colorAllocate (0,0,0);
	my $gray2 = $Image->colorAllocate (192,222,192);

	# colors used when plotting
	my @Colors;
	foreach $ColorSpec (@ColorSet)
		{
		push @Colors, $Image->colorAllocate (@$ColorSpec);
		}
	$Image->transparent($Linen);

	PQDisplayGraphTemplate ($Image, $black, $gray);

	PQPlotBestFitLine ($Image, $PQData, $black, $gray2);

	PQPlotBids ($Image, \@Colors, $PQData, $LinkName);

	my @Bounds = $Image->getBounds ();

	my $ImageFile = "$DataDir/PQ_$ContextID.png";
	open (PIC, ">$ImageFile") or MyDie ("Cannot create image file $ImageFile");
	binmode PIC;
	print PIC $Image->png();
	close PIC;
	return ($ImageFile, @Bounds);
	}

#
# Calculates graph metrics that are needed to create the PQ graph
#
# calculates the following %Context values:
#
# XMin		 -	Minimum Data Value 
# XMax		 -	Maximum Data Value 
# YMin		 -	Minimum Data Value 
# YMax		 -	Maximum Data Value 
# XTrueMin	 -	Minimum Data Value (not adjusted to get a border)
# XTrueMax	 -	Maximum Data Value (not adjusted to get a border)
# XRange		 -	Range in X Values
# YRange		 -	Range in Y Values
# XMarkers	 -	X Labels
# YMarkers	 -	Y Labels
# XGridMin	 -	Pixel X location of left edge of graph area
# YGridMin	 -	Pixel Y location of minimum Y Value
# XGridMax	 -	Pixel X location of right edge of graph
# YGridMax	 -	Pixel Y location of maximum Y Value
# XGridSize	 -	X Size of grid display area in pixels
# YGridSize	 -	Y Size of grid display area in pixels
# Font		 -	Font to use for the labels
#
sub PQCalcMetrics
	{
	my ($PQData) = @_;

	$Context{XMin} = 1000000;
	$Context{XMax} = 0;
	$Context{YMin} = 1000000;
	$Context{YMax} = 0;

	my ($Bid, $Price, $Quan);

	foreach $Bid (@$PQData)
		{
		($Price, $Quan) = @$Bid;
	
		$Context{XMin} = ($Context{XMin} < $Quan ?  $Context{XMin} : $Quan); # not using this
		$Context{XMax} = ($Context{XMax} > $Quan ?  $Context{XMax} : $Quan);
		$Context{YMin} = ($Context{YMin} < $Price ? $Context{YMin} : $Price); # not using this
		$Context{YMax} = ($Context{YMax} > $Price ? $Context{YMax} : $Price);
		}
	$Context{XTrueMin} = $Context{XMin}; #	Used for best fit line
	$Context{XTrueMax} = $Context{XMax}; #

	my @XMarkers = CreateReasonableMarkers ($Context{XMax}, 5);
	$Context{XMin}		 = 0;
	$Context{XMax}		 = $XMarkers[-1];
	$Context{XRange}	 = $Context{XMax} - $Context{XMin};
	$Context{XMarkers} = [@XMarkers];

	my @YMarkers = CreateReasonableMarkers ($Context{YMax}, 10);
	$Context{YMin}		 = 0;
	$Context{YMax}		 = $YMarkers[-1];
	$Context{YRange}	 = $Context{YMax} - $Context{YMin};
	$Context{YMarkers} = [@YMarkers];

	$Context{XGridMin} = $Context{PQLeftEdge};
	$Context{XGridMax} = $Context{PQXSize} - $Context{PQRightEdge};
	$Context{XGridSize} = $Context{XGridMax} - $Context{XGridMin} + 1;

	$Context{YGridMin} = $Context{PQTopEdge};
	$Context{YGridMax} = $Context{PQYSize} - $Context{PQBottomEdge};
	$Context{YGridSize} = $Context{YGridMax} - $Context{YGridMin} + 1;

	$Context{Font} = gdMediumBoldFont;
	}


sub PQDisplayGraphTemplate
	{
	my ($Image, $Color1, $Color2) = @_;
	my ($i, $x, $y);

	#draw frame
	for ($i=0; $i<2; $i++)
		{
	   $Image->line ($Context{XGridMin}-$i, $Context{YGridMin}, 
	                 $Context{XGridMin}-$i, $Context{YGridMax}, $Color1); # vert
	   $Image->line ($Context{XGridMin}, $Context{YGridMax}+$i, 
	                 $Context{XGridMax}+1, $Context{YGridMax}+$i, $Color1); #horz
		}
	
	# Price Labels and gridlines
	my $YMarkers = $Context{YMarkers};
	my $XMarkers = $Context{XMarkers};
	my $HY = $Context{HalfYLabel};
	my $LE = $Context{XGridMin};
	my $BE = $Context{YGridMax};
	my ($Value, $Label, $xLabel);

	foreach $Value (@$YMarkers)
		{
		$y = YPos ($Value);
		$Image->line ($LE-4, $y, $LE, $y, $Color1);
		$Image->line ($LE+1, $y, $Context{XGridMax}, $y, $Color2) if ($y < $BE);
		$Label = MoneyFormat ($Value, "");

		$x = $LE - length ($Label) * 7 - 7;
		$Image->string(gdMediumBoldFont, $x, $y-$HY, $Label, $Color1);
		}

	# Quantity Labels and gridlines
	foreach $Value (@$XMarkers)
		{
		$x = XPos ($Value);

		$Image->line ($x, $BE+2, $x, $BE+5, $Color1);
		$Image->line ($x, $BE-1, $x, $Context{YGridMin}, $Color2) if ($x > $LE);
		$Label = NumberFormat ($Value, ($Context{XMax} < 10 ? ($Context{XMax} < 1 ? "N4" : "N2") : "N0"));

		$xLabel = $x - (length ($Label) * 7/2);
		$Image->string(gdMediumBoldFont, $xLabel, $BE+7, $Label, $Color1);
		}
	}


sub PQPlotBids
	{
	my ($Image, $Colors, $PQData, $LinkName) = @_;
	my ($Dia, $x, $y, $Price, $Quan, $VendorID, $ContID, $Bid, $Color);

	$Dia = $Context{PQPointSize};
	foreach $Bid (@$PQData)
		{
		($Price, $Quan, $VendorID, $ContID) = @$Bid;

		$Color = ($VendorID eq "Estimate" ? $Colors->[2] : $Colors->[1]);
		$x = XPos ($Quan);
		$y	= YPos	($Price);
      map $Image->arc($x,$y,$_,$_, 0, 360, $Color), ($Dia,$Dia-2,$Dia-4);

		AddMapCircle ($x,$y, $Dia, "$LinkName?Vendor=$VendorID&Contract=$ContID", 
		                           "Vendor: $VendorID\n".
		                           "Unit Price: $Price\n".
		                           "Quantity: $Quan\n".
		                           "Deviation: ???");
		}
	}


sub PQPlotBestFitLine
	{
	my ($Image, $PQData, $Color1, $Color2) = @_;
	my ($x, $y, $y1, $y2, $yA, $oldX, $oldY, $xVal, $yVal, $yVal1, $yVal2, $Increment, $Steps);

	my $Coeff = {};
	return if (!PQCalcCoefficients ($PQData, $Coeff));

	$Steps = $Context{PQSteps} || 250;
	$Increment = ($Context{XTrueMax} - $Context{XTrueMin}) / $Steps;
	$oldX = $oldY = 0;

	return if (!$Increment);

	for ($xVal = $Context{XTrueMin}; $xVal <= $Context{XTrueMax}; $xVal += $Increment)
		{
		$yA =   $Coeff->{Lambda} + $Coeff->{Beta} * ($Context{PQLogQuan} ? log ($xVal)  : $xVal);
		$yVal = ($Context{PQLogPrice} ? exp ($yA) : $yA);

		#--- plot confidence interval lines ---#
		($yVal1, $yVal2) = PQCalcConfidenceInterval ($xVal, $Coeff);
		$yVal1 = min ($Context{YMax}, max ($Context{YMin}, $yVal1));
		$yVal2 = min ($Context{YMax}, max ($Context{YMin}, $yVal2));
		($y1, $y2) = (YPos($yVal1), YPos($yVal2));

		#--- plot curve ---#
		$x = XPos ($xVal);
		$y	= YPos	($yVal);
		$Image->line ($oldX, $oldY, $x, $y, $Color1) if ($oldX);
		$Image->line ($oldX, $oldY+1, $x, $y+1, $Color1) if ($oldX);
		$oldX = $x;
		$oldY = $y;

		$Image->line ($x, $y1, $x, $y2, $Color2);
		}
	}



sub PQCalcCoefficients
	{
	my ($PQData, $Coeff) = @_;
	my ($n, $Beta, $Lambda, $SumSqDif, $Sigma, $Bid, $Price, $Quan, $AvgX, $AvgY, $Num, $x, $y);

	$Coeff->{OK} = 0;

	$n = @$PQData;	# n is number of points

	return 0 if ($n < 3);
	
	$AvgX = $AvgY = 0;
	foreach $Bid (@$PQData)
		{
		($Price, $Quan) = @$Bid;
		$AvgX += ($Context{PQLogQuan}  ?  log ($Quan)  : $Quan );
		$AvgY += ($Context{PQLogPrice} ?  log ($Price) : $Price);
		}
	$AvgX /= $n;
	$AvgY /= $n;

	$Num = $SumSqDif = 0;
	foreach $Bid (@$PQData)
		{
		($Price, $Quan) = @$Bid;
		$x = ($Context{PQLogQuan}  ?  log ($Quan)  : $Quan );
		$y = ($Context{PQLogPrice} ?  log ($Price) : $Price);
		$Num += ($x - $AvgX)*($y - $AvgY);
		$SumSqDif += ($x - $AvgX)*($x  - $AvgX);
		}
	return 0 if (!$SumSqDif);

	$Beta =$Num/$SumSqDif;
	$Lambda = $AvgY - $Beta * $AvgX;

	$Num = 0;
	foreach $Bid (@$PQData)
		{
		($Price, $Quan) = @$Bid;
		$x = ($Context{PQLogQuan}  ? log ($Quan)  : $Quan );
		$y = ($Context{PQLogPrice} ? log ($Price) : $Price);
		
		$Num += ($y - $Lambda - $Beta * $x) * ($y - $Lambda - $Beta * $x);
		}
	$Sigma = sqrt (1/($n-2) * $Num);

	$Coeff->{OK} 		 =  1;
	$Coeff->{n} 		 =  $n;
	$Coeff->{Beta} 	 =  $Beta;
	$Coeff->{Lambda} 	 =  $Lambda;
	$Coeff->{SumSqDif} =  $SumSqDif;
	$Coeff->{Sigma} 	 =  $Sigma;
	$Coeff->{AvgX} 	 =  $AvgX;
	$Coeff->{AvgY} 	 =  $AvgY;
	$Coeff->{TDist} 	 =  TDistribution ($Coeff->{n} - 2);

	return 1;
	}

sub PQCalcConfidenceInterval
	{
	my ($x, $Coeff) = @_;

	my $LogX = ($Context{PQLogQuan}  ?  log ($x) : $x);
	my $a = $Coeff->{Lambda} + $Coeff->{Beta} * $LogX;
	my $b = $Coeff->{TDist};
	my $c = (($LogX-$Coeff->{AvgX})*($LogX-$Coeff->{AvgX}))/$Coeff->{SumSqDif};
	my $d = $Coeff->{Sigma} * sqrt (1/$Coeff->{n} + $c);

	my $y0 =	($Context{PQLogPrice} ?  exp($a + $b * $d) : $a + $b * $d);
	my $y1 =	($Context{PQLogPrice} ?  exp($a - $b * $d)	: $a - $b * $d);
	return ($y0, $y1);
	}



sub TDistribution
	{
	my ($n) = @_;

	$n = int (($n > 120) ? 120 : $n);
	return $TDistr[$n];
	}

#############################################################################
#
# Graphing Util Fn's
#
#############################################################################
# Converts a Deviation Range Value into a Graph Y Location
# Uset to get the Y location of all points and labels on the graph
#

sub YPos
	{
	my ($Val) = @_;

	my $y = $Context {YGridMax} - ($Val - $Context{YMin}) / $Context{YRange} * $Context{YGridSize};
	return int ($y + 0.5);
	}

sub XPos
	{
	my ($Val) = @_;

	my $x = $Context {XGridMin} + ($Val - $Context{XMin}) / $Context{XRange} * $Context{XGridSize};
	return int ($x + 0.5);
	}


# determines if there is enough distance between the two Y distances to display the labels for both
# returns 0 if no, or the distance in pixels if true
sub SufficientGap
	{
	my ($Val1, $Val2) = @_;

	my $Dist = abs (YPos($Val1) - YPos ($Val2));
	return ($Dist >= $Context{HalfYLabel}*2 ? $Dist : 0);
	}



sub CreateReasonableMarkers
	{
	my ($V, $MaxMarkers)  = @_;
	my ($i, @Markers);

	my $lg= Log10($V);
	my $a = (int $lg) + ($lg > 0 ? -1 : -2);
	my $b	= $V * Exp10 (-$a);
	my $c = int ($b + 0.001);
	my $d = ($c+1) * Exp10($a);
#	my $e = int ($c/10 + 0.5);
	my $g = (($c+1)/$MaxMarkers);
	my $f = ($g < 1 ? int ($g*10)/10 : int ($g+0.5));
	$f    = ($f>6?10:($f>4?5:($f>2.75?3:($f>2.25?2.5:($f>1.5?2:($f>1?1:$f))))));

	push @Markers, sprintf ("%f", 0);
	for ($i=0.0; $i < $d; )
		{
		$i = $i + $f * Exp10 ($a);
		push @Markers, sprintf ("%f", $i);
		}
	return @Markers;
	}


sub Log10
	{
	return log($_[0])/log(10);
	}

sub Exp10
	{
	return exp ($_[0] * log (10));	
	}



# If the colored dots do not exist on the server, they are created here.
# These are needed to display the vendor table beneath the graph.
#
sub CreateDingbats
	{
	my ($i);

	for ($i=0; $ColorSet[$i]; $i++)
		{
		my $Image = new GD::Image(11,11);
		my $Linen = $Image->colorAllocate(250,240,230);
		$Image->transparent($Linen);
		my $DotColor = $Image->colorAllocate (@{$ColorSet[$i]});
	   map $Image->arc(5,5,$_,$_,0,360, $DotColor), (11,9,7,5,3);
#		my $PicName = sprintf ("V%2.2d.png", $i+1);
		my $PicName = sprintf ("V%2.2d.jpg", $i+1); # lying about the filetype!
		open (PIC, ">$DataDir/$PicName") or MyDie ("Cannot create image file $DataDir/$PicName");
		binmode PIC;
		print PIC $Image->png();
		close PIC;

		undef $Image;
		}

	}

#############################################################################
#
# Data gathering - Specific to getting BidTab Data
#
#############################################################################

# This loads the data used by the LIP graph
# This contains item/vendor/bid info.
#
sub LoadBidtabData
	{
	my ($Data);
	
	$Data = LoadBidtabDataFromOLAP (@_) if ($Context{Provider} eq "MSOLAP");

	$Data = LoadBidtabDataFromODBC (@_) if ($Context{Provider} ne "MSOLAP");

###
print "Data1 = $Data<br>\n";
	LoadBidTotalsAndVendorNames ($Data);

###
print "Data2 = $Data<br>\n";
	BidtabCalc_BidTotal		($Data);

###
print "Data3 = $Data<br>\n";
	BidtabCalc_Dev				($Data);

###
print "Data4 = $Data<br>\n";

	BidtabCalc_FilterItems	($Data);
 	BidtabCalc_PctEst			($Data);
	BidtabCalc_DisplayTotal	($Data);
	return $Data;
	}

sub LoadBidtabDataFromOLAP
	{
	my $Conn = Connect();

	my $Threshold = $Context{LIPMinDev};

#	# --- tries (incorrectly) to filter on threshold ---
	my $Query = 
		"select {[Measures].[Wtd Per Vs Est], [Measures].[Price], [Measures].[Quantity], [Measures].[Amount]} on columns, \n".
		"NON EMPTY  CrossJoin (Filter ({[items].[item id].members}, \n".
	   "([Wtd Per Vs Est]) >= $Threshold), [vendors].[vendor id].members) on rows \n".
		"from $Context{Cube}\n".
		"where [contracts].[contract id].[$Context{Contract}] ";

# Why is this so slow running locally?
#
#	my $Query = 
#		"select {[Measures].[Wtd Per Vs Est], [Measures].[Price], [Measures].[Quantity], [Measures].[Amount]} on columns, \n".
#		"NON EMPTY  CrossJoin ({[items].[item id].members}, [vendors].[vendor id].members) on rows \n".
#	   "from $Context{Cube}\n".
#		"where [contracts].[contract id].[$Context{Contract}] \n";
#

# A test case
#	my $Query = 
#"select {[Measures].[Wtd Per Vs Est], [Measures].[Price], [Measures].[Quantity], [Measures].[Amount]} on columns, ".
#"NON EMPTY  CrossJoin ({[items].[item id].members}, [vendors].[vendor id].members}) on rows ".
#"from newprofile ".
#"where [contracts].[contract id].[CALL6086] ";
#



	print "<B>Query:</B><pre>$Query</pre><hr>\n" if ($Context{Debug});

	my $CellSet = new Win32::OLE ("ADOMD.Cellset") or die "Oops, cannot create ADODB CellSet";

	$CellSet->Open ($Query, $Conn);

	ConnError ($Conn, $Query) if ($Conn->Errors->{Count});
			 
	my $Data = ();	# the entire data set - a ref to a list of $Item

	my ($j);
	my ($Item, $ItemID, $LastItemID);
	my ($Vendor, $VendorID, $Deviation);

	for ($j=0; $j < $CellSet->Axes(1)->{Positions}->{Count} && $j < 2000; $j++)
		{
		$ItemID = $CellSet->Axes(1)->Positions($j)->Members(0)->{Caption};

		if ($ItemID ne $LastItemID) 
			{
			$Item = {};			  # create a new item hash 
			push @$Data, $Item; # add new item to the @$Data list
			$Item->{ID} = $LastItemID = $ItemID;
			$Item->{Desc} 		= "Item Description Goes Here";
			$Item->{Unit} 		= "Unit!";
			$Item->{Quantity} = "Quan!";
			}

		$Vendor = {};	# create a new vendor hash 
		$VendorID = $CellSet->Axes(1)->Positions($j)->Members(1)->{Caption};
		$Vendor->{Dev}	  		= $CellSet->Item (0, $j)->FormattedValue();
		$Vendor->{Price} 		= StripNumberStuff ($CellSet->Item (1, $j)->FormattedValue());
		$Vendor->{Quantity} 	= StripNumberStuff ($CellSet->Item (2, $j)->FormattedValue());
		$Vendor->{Amount}   	= StripNumberStuff ($CellSet->Item (3, $j)->FormattedValue());
		$Item->{Quantity} 	= $Vendor->{Quantity};

		$Item->{Bids}->{$VendorID} = $Vendor;	# add new Vendor to the $Item->{Bids} set

		}
	$CellSet->Close ();
	DisConnect ($Conn);
	return $Data;
	}



sub LoadBidtabDataFromODBC
	{
	my $Conn = Connect();

##	my $Query =	" Select ItemNumber, VendorID, UnitPrice, Quantity, Extension\n".
##					" From BidTabs".
##					" Where ContractID = '$Context{Contract}'\n".
##					" order by ItemNumber";

my $Query =
" select item.itemnumber, vendorid, itembid.unitprice, quantity, itembid.unitprice * quantity as extension\n".
" from itembid, item\n".
" where itembid.projectid = '$Context{Contract}'\n".
" and item.projectid = itembid.projectid\n".
" and item.itemid = itembid.itemid\n";
" order by item.itemnumber";

	print "<B>Query:</B><pre>$Query</pre><hr>\n" if ($Context{Debug});

##	my $Result = $Conn->Execute ($Query) || ConnError ($Conn, $Query);
   my $sth = $Conn->prepare ($Query);
   $sth->execute ();

	my $Data = ();	# the entire data set - a ref to a list of $Item

	# First we load it all
	my ($i, $LastItemID, $Item, $Vendor, $array);
	for ($i = 0; ($array = $sth->fetchrow_arrayref ()) && $i < 5000; $i++) 
		{
###
###print "Getting a row!!!<br>\n";
		my ($ItemID, $VendorID, $UnitPrice, $Quan, $Ext) = @{$array};

		if ($ItemID ne $LastItemID) 
			{
			$Item = {};			  # create a new item hash 
			push @$Data, $Item; # add new item to the @$Data list
			$Item->{ID} = $LastItemID = $ItemID;
			$Item->{Desc} 		= "Item Description Goes Here";
			$Item->{Unit} 		= "Unit!";
			$Item->{Quantity} = sprintf ("%.3f", $Quan);
			}
		$Vendor = {};	# create a new vendor hash 
		$Vendor->{Dev}	  		= 0;
		$Vendor->{Quantity} 	= sprintf ("%.3f", $Quan);
		$Vendor->{Price} 		= sprintf ("%.2f", $UnitPrice);
		$Vendor->{Amount}   	= sprintf ("%.2f", $Ext);
		$Item->{Bids}->{$VendorID} = $Vendor;	# add new Vendor to the $Item->{Bids} set

##	   $Result->MoveNext;
		}
##	$Result->Close;
   $sth->finish();
##	DisConnect ($Conn);

	return $Data;
	}


# Loads additional data to $Data
# Prereq:
# 	LoadBidtabData()
#
sub LoadBidTotalsAndVendorNames
	{
	my ($Data) = @_;

	my $QueryOlapSQL = # SQL if using OLAP SQL 
			" Select [vendors].[vendor id], [vendors].[vendor], sum(amount) \n".
			" from $Context{Cube}\n".
			" where [contract id] = '$Context{Contract}'\n".
			" group by [vendors].[vendor id], [vendors].[Vendor]";

	my $QueryODBCSQL = # SQL if using ODBC
##			" Select Bidders.VendorID, VendorName, Total\n".
##			" From Bidders, Vendors\n".
##			" Where Bidders.VendorID = Vendors.VendorID\n".
##			" and Bidders.ContractID = '$Context{Contract}'";
" select bid.vendorid, name, itemtotal\n".
" from bid, vendor\n".
" where bid.projectid = '$Context{Contract}'\n".
" and vendor.vendorid = bid.vendorid";


	my $Query =	($Context{Provider} eq "MSOLAP" ? $QueryOlapSQL : $QueryODBCSQL);
	print "<B>Query:</B><pre>$Query</pre><hr>\n" if ($Context{Debug});

	my $Conn = Connect();
##	my $Result = $Conn->Execute ($Query) || ConnError ($Conn, $Query);
   my $sth = $Conn->prepare ($Query);
   $sth->execute ();

	my $Bids = $Data->[0]->{Bids};
	my ($i, $array);
	for ($i = 0; ($array = $sth->fetchrow_arrayref ()) && $i < 100; $i++) 
		{
##		my ($VendorID, $VendorName, $BidTotal) = map $_->{Value}, in $Result->Fields;
		my ($VendorID, $VendorName, $BidTotal) = @{$array};
		$Bids->{$VendorID}->{Name}		 =	$VendorName;
		$Bids->{$VendorID}->{BidTotal} =	$BidTotal;

##	   $Result->MoveNext;
		}
##	$Result->Close;
   $sth->finish();
##	DisConnect ($Conn);
	}



# Calculates % Deviation for each bid
# Prereq:
# 	LoadBidtabData()
#	LoadBidTotalsAndVendorNames () and/or BidtabCalc_BidTotal
#
sub BidtabCalc_Dev
	{
	my ($Data) = @_;

	my ($Item0, $VendorID, $Item, $ItemEst, $BidEst, $Vendor);
	$Item0 = $Data->[0];
	foreach $VendorID (GetVendorList ($Data))
		{
		$BidEst = $Item0->{Bids}->{Estimate}->{BidTotal};
		foreach $Item (@$Data)
			{
			$ItemEst = $Item->{Bids}->{Estimate}->{Amount};
			$Vendor = $Item->{Bids}->{$VendorID};

#debug
#my $a = ($Vendor->{Amount} - $ItemEst)/$BidEst * 100;
#$a = sprintf ("%2.2f", $a);
#print "DEV: $Vendor->{Dev} -vs- $a  ($Vendor->{Amount} - $ItemEst)/$BidEst * 100)<br>\n";
#debug
#debug
   $BidEst ||= 1;
#debug
#debug



			$Vendor->{Dev} = ($Vendor->{Amount} - $ItemEst)/$BidEst * 100;
			$Vendor->{Dev} = sprintf ("%.2f", $Vendor->{Dev});
			}
		}
	}


# filters out items that don't have a significant deviation
# Prereq:
# 	LoadBidtabData()
#	LoadBidTotalsAndVendorNames () and/or BidtabCalc_BidTotal()
#	BidtabCalc_Dev ()
#
sub BidtabCalc_FilterItems
	{
	my ($Data) = @_;

	my $Threshold = $Context{LIPMinDev};
	my @VenList = GetVendorList ($Data);

	my ($i, $Item, $KeepItem, $VendorID);
	for ($i=0; $i < @$Data; )
		{
		$Item = $Data->[$i];
		$KeepItem = 0;
		foreach $VendorID (@VenList)
			{
			$KeepItem = 1 if (abs $Item->{Bids}->{$VendorID}->{Dev} >= $Threshold);
			}
		splice @$Data, $i, 1 if (!$KeepItem);
		$i++ if ($KeepItem);
		}
	}


# Calculates PctEst (percent of estimate) for each bid
# Prereq:
# 	LoadBidtabData()
# 	LoadBidTotalsAndVendorNames () and/or  BidtabCalc_BidTotal()
#
sub BidtabCalc_PctEst
	{
	my ($Data) = @_;

	# get baseline bid total
	my $BaselineVendor = GetBaselineVendorID ($Data);
	my $BaselineTotal = $Data->[0]->{Bids}->{$BaselineVendor}->{BidTotal};

	my ($Item, $VendorID);
	# calculate percent of baseline bid for each bid item.
	foreach $Item (@$Data)
		{
		foreach $VendorID (GetVendorList ($Data))
			{
			my $Vendor =  $Item->{Bids}->{$VendorID};
			$Vendor->{PctEst} = $Vendor->{Amount} / $BaselineTotal * 100  if ($BaselineTotal);
			}
		}
	}


# Totals items
# used by
# 	BidtabCalc_DisplayTotal
# 	BidtabCalc_BidTotal
# Don't use directly
#
sub _BidtabCalc_Total
	{
	my ($Data, $HashKey) = @_;

	# calculate total for displayed (loaded) items
	my ($Item0, $Vendor0, $VendorID, $Vendor, $Item);
	$Item0 = $Data->[0];
	foreach $VendorID (GetVendorList ($Data))
		{
		my $Vendor0 = $Item0->{Bids}->{$VendorID};
		foreach $Item (@$Data)
			{
			$Vendor = $Item->{Bids}->{$VendorID};
			$Vendor0->{$HashKey} += $Vendor->{Amount};
			}
		}
	}


# Adds calculated Bid Total to $Data
# Prereq:
# 	LoadBidtabData()
#
# Assumes all items in proposal are loaded
#
sub BidtabCalc_BidTotal
	{
	my ($Data) = @_;
	_BidtabCalc_Total ($Data, "BidTotal");
	}


# Adds calculated Total of loaded items to $Data
# Prereq:
# 	LoadBidtabData()
#	LoadBidTotalsAndVendorNames () and/or BidtabCalc_BidTotal
# 	BidtabCalc_Dev()
# 	BidtabCalc_FilterItems()
#
# Assumes items are already filtered
#
sub BidtabCalc_DisplayTotal
	{
	my ($Data) = @_;
	_BidtabCalc_Total ($Data, "DisplayTotal");
	}



#############################################################################
#
# Data gathering - general
#
#############################################################################

sub GetBaselineVendorID
	{
	my ($Data) = @_;

	my $BaselineVendor = $Context{BaseBid} || "Estimate";
	$BaselineVendor = (keys %{$Data->[0]->{Bids}})[0] if (!$Data->[0]->{Bids}->{$BaselineVendor});

	return $BaselineVendor;
	}




sub GetVendorList
	{
	my ($Data) = @_;
	my (@VenList, $Vendor);

	foreach $Vendor (sort keys %{$Data->[0]->{Bids}})
		{
		if ($Vendor eq "Estimate")
			{
			unshift @VenList, $Vendor;
			}
		else
			{
			push @VenList, $Vendor;
			}
		}
	return @VenList;
	}



sub LoadSQLTable
	{
	my ($Query, $ColumnMap, @FillIns) = @_;

	my $Conn = Connect();
	print "<B>Query:</B><pre>$Query</pre><hr>\n" if ($Context{Debug});

##	my $Result = $Conn->Execute ($Query) || ConnError ($Conn, $Query);
   my $sth = $Conn->prepare ($Query);
   $sth->execute ();


	my ($i, @SQLRow, @SQLEntry, $SQLData, $MaxMapEntry);
	$MaxMapEntry = @$ColumnMap - 1;

	my ($array);
	for ($i = 0; ($array = $sth->fetchrow_arrayref ())&& $i < 5000; $i++) 
		{
##		@SQLRow = map $_->{Value}, in $Result->Fields;
		@SQLRow = @{$array};
		@SQLEntry = ();
		if ($MaxMapEntry == -1) # no mapping defined
			{
			@SQLEntry = @SQLRow;
			}
		else
			{
			map {push @SQLEntry, ($ColumnMap->[$_] ? $SQLRow[$ColumnMap->[$_]-1] : $FillIns[$_])} (0..$MaxMapEntry);
			}
		push @$SQLData, [@SQLEntry];
##	   $Result->MoveNext;
		}
##	$Result->Close;
   $sth->finish();
	DisConnect ($Conn);

	return $SQLData;
	}

#############################################################################
#
#	Database Utilities Section
#
#############################################################################

my $PersistentConnection = 0;

sub Connect
	{
##	return ($PersistentConnection) if $PersistentConnection;
##
##	my $Provider	= $Context{Provider};   # "MSOLAP";
##	my $DataSource	= $Context{DataSource}; # "CRAIGF";
##	my $Catalog   = $Context{Catalog};   # "dss";
##
##	my $Conn = new Win32::OLE ("ADODB.Connection") or MyDie ("Cannot create ADODB object");
##
##	my $ConnectStr = "DSN=$DataSource";
##	$ConnectStr   .= ";Provider=$Provider" if ($Provider);
##	$ConnectStr   .= ";initial catalog=$Catalog" if ($Catalog);
##
##	print "<B>ConnectStr:</B><pre>$ConnectStr</pre><hr>\n" if ($Context{Debug} > 1);
##
##	$Conn->Open($ConnectStr);
##	MyDie ("Cannot connect using $ConnectStr", $ConnectStr) if (!$Conn);
##
###DumpHash ($Conn);
###DumpHash ($Conn->{Errors});
##
##	$PersistentConnection = $Conn;
##
##	return $Conn;

   return DBI->connect('DBI:mysql:database=devel_preconst;host=localhost') or SysError("Could not connect to database");
	}


sub DisConnect
	{
#	$PersistentConnection->Close();
	}

sub SingleRowSQL
	{
	my ($Query) = @_;

	my $Conn = Connect();
	print &MiniHead() . "<B>Query:</B><pre>$Query</pre><hr>\n" if ($Context{Debug});

   return $Conn->selectrow_array ($Query);


##	my $Result = $Conn->Execute ($Query) || ConnError ($Conn, $Query);
##	my @Row = ();
##	if (!$Result->EOF)
##		{
##		@Row = map $_->{Value}, in $Result->Fields;
##		}
##	$Result->Close;
##	DisConnect ($Conn);
##
##	return @Row;
	}


sub SingleRowMDX
	{
	}



sub TableFromSQL
	{
	my ($Query, $Formats, $Labels) = @_;
	my ($i, $j, @Row);

	my $Conn = Connect();
	print "<B>Query:</B><pre>$Query</pre><hr>\n" if ($Context{Debug});

##	my $Result = $Conn->Execute ($Query) || ConnError ($Conn, $Query);
   my $sth = $Conn->prepare ($Query);
   $sth->execute ();
   my @results = ();

	TableStart();
	TableHeader ($Labels);

   while (my $array = $sth->fetchrow_arrayref ())
##	for ($i = 0;!$Result->EOF && $i < 2000; $i++) 
		{
##		@Row = map $_->{Value}, in $Result->Fields;
##		TableRow (\@Row, $Formats);
		TableRow ($array, $Formats);
##	   $Result->MoveNext;
		}
	TableEnd ();
	print "<center><font size=1>$i matching Rows<font></center><br>\n";
##	$Result->Close;
   $sth->finish();

	DisConnect ($Conn);
	}


sub TableFromMDX
	{
   die "No mdx!";

##	my ($Query, $Formats, $Labels) = @_;
##
##	my $Conn = Connect();
##	print "<B>Query:</B><pre>$Query</pre><hr>\n" if ($Context{Debug});
##
##	my $CellSet = new Win32::OLE ("ADOMD.Cellset") or die "Oops, cannot create ADODB CellSet";
##	$CellSet->Open ($Query, $Conn);
##
##	ConnError ($Conn, $Query) if ($Conn->Errors->{Count});
##
##	TableStart();
##	TableHeader ($Labels);
##
##	my ($j, $i, $NewLettingID, @Row);
##	for ($i=0; $i < $CellSet->Axes(1)->Positions->{Count}; $i++)
##		{
##		@Row = ();
##		for ($j=0; $j < $CellSet->Axes(1)->{DimensionCount}; $j++)
##			{
##			push @Row, $CellSet->Axes(1)->Positions($i)->Members($j)->{Caption};
##			}
##		for ($j=0; $j < $CellSet->Axes(0)->Positions->{Count}; $j++)
##			{
##			push @Row, $CellSet->Item($j, $i)->{FormattedValue};
##			}
##		TableRow (\@Row, $Formats);
##		}
##	TableEnd ();
##
##	print "<center><font size=1>$i matching Rows<font></center><br>\n";
##	$CellSet->Close ();
##	DisConnect ($Conn);
	}

sub TableFromVBObject
	{
   die "No vb objects!";
##
##	my ($OLEName, $Formats) = @_;
##
##	my $obj;
##	if (!($obj = Win32::OLE->GetActiveObject($OLEName)))
##		{
##		$obj = Win32::OLE->new($OLEName) or MyDie ("Oops, cannot create OLE object $OLEName");
##		}
##	map $obj->SetKey ($_, $Context{$_}), keys %Context;
##
##	$obj->SetKey ("Letting"  , "1997" ); # This is dummied up for new
##	$obj->SetKey ("ItemClass", "ASPH" ); # This is dummied up for new
##	$obj->Connect();
##
##	my $ErrStr = $obj->Exec();
##	MyDie ("Error from $OLEName :" . $ErrStr) if ($ErrStr);
##
##	print "<B>Query:</B><pre>$obj->{strQuery}</pre><hr>\n" if ($Context{Debug});
##
##	TableStart();
##	my @Labels = split ("\t", $obj->GetLabel());
##	TableHeader (\@Labels);
##	my ($i, $RowSet, @Row);
##	for ($i=0; ($RowSet = $obj->GetRow()) && $i<1000; $i++)
##			{
##			last if ($RowSet eq "Cellset is Null");
##			@Row = split "\t", $RowSet;
##			TableRow (\@Row, $Formats);
##			}
##	$obj->Disconnect();
##	TableEnd();
##	print "<center><font size=1>$i matching Rows<font></center><br>\n";
	}


#############################################################################
#
#	HTML Utilities Section
#
#############################################################################

sub MiniHead
	{
   print "Pragma: no-cache\n";
   print "Content-type: text/html\n\n";
	}

sub NewCommonHeader
	{
	my ($Title, $Header, $Links1, $Links2, $Sizes) = @_;

	MiniHead ();

   print "<HTML>\n";
   print "<head><title>$Title</title></head>\n";
   print "<body bgcolor='$BackgroundColor'>\n";
   print "<TABLE width='100%' border=0 cellpadding=0 cellspacing=0>\n";
   print " <TR>\n";
   print "  <TD width='$Sizes->[0]%' valign=top>$Header</TD>\n";
   print "  <TD width='$Sizes->[1]%' valign=top>$Links1</TD>\n";
   print "  <TD width='$Sizes->[2]%' valign=top>$Links2</TD>\n";
   print " </TR>\n";
   print " <TR>\n";
   print "   <td colspan=3><hr noshade size=2></td>\n";
   print " </TR>\n";
   print "</TABLE>\n";
	}

sub CommonHeader
	{
	my ($Msg, $Msg2, $Links) = @_;

	MiniHead ();

   print "<html>\n";
   print "  <head><title>Letting Analysis: $Msg</title></head>\n";
   print "  <body bgcolor='$BackgroundColor'>\n";
   print "<TABLE width='100%' border=0 cellpadding=0 cellspacing=0>\n";
   print "	<TR>\n";
   print "		<TD>\n";
   print "		<FONT size=6><b>Letting Analysis</b></FONT>\n";
   print "		</TD>\n";

   if ($Links)
		{
	   print "		<td align='right' valign='bottom' rowspan=2>\n";
	   print "		 $Links\n";
		}
	else # the logo
		{
   	print "		<td align='center' valign='bottom' rowspan=2>\n";
   	print "		<a href=\"/cgi-bin/dm.pl\"><img src='/images/xh_linen.gif' border=0></a>\n";
		}
   print "		</TD>\n";
   print "	</TR>\n";
   print "		<TD>\n";
   print "			   <FONT size=4><b>$Msg</b></FONT>\n";
   print "			   <br>$Msg2\n" if ($Msg2);
   print "		</TD>\n";
   print "		<TD>\n";
   print "		</TD>\n";
   print "	<TR>\n";
   print "	</TR>\n";
   print "  <tr>\n";
   print "    <td colspan=2><hr noshade size=2></td>\n";
   print "  </tr>\n";
   print "</TABLE>\n";
	}


sub CommonFooter
	{
	my ($Msg1) = @_;
	
	printf "<br><br>\n";
	print "<font size=1>$Msg1</font><br>\n";
	printf "<font size=1><a href=\"Notes?From=$Target\">Prototype Notes</a></font><br>\n";
	printf "<font size=1><a href=\"Settings?From=$Target\">Change Settings</a></font><br>\n";
	printf "<font size=1><a href=\"Debug\">Debug Info</a></font><br>\n" if (exists $Context{Debug});
	print "</body></html>\n";
	} 

sub TableOrChartChoice
	{
	my ($LinkName) = @_;

	if ($Context{PQGraph})
		{
		print "<center><font size=2><b><a href=\"$LinkName?PQGraph=0\">Display Table</a> | ";
		print "Display Price/Quantity Graph</b></font></center>\n";
		}
	else
		{
		print "<center><font size=2><b>Display Table | ";
		print "<a href=\"$LinkName?PQGraph=1\">Display Price/Quantity Graph</a></b></font></center>\n";
		}
	}

sub LIPTableOrChartChoice
	{
	my ($LinkName) = @_;

	if ($Context{LIPGraph})
		{
		print "<center><font size=2><b><a href=\"$LinkName?LIPGraph=0\">Display BidTabs</a> | ";
		print "Display Line Item Profile Graph</b></font></center>\n";
		}
	else
		{
		print "<center><font size=2><b>Display BidTabs | ";
		print "<a href=\"$LinkName?LIPGraph=1\">Display Line Item Profile Graph</a></b></font></center>\n";
		}
	}




sub ChartHTML
	{
	my ($XSize, $YSize) = @_;

	my $Rand = sprintf ("%5.5d", rand (32000));

	print "<center>\n";
	print "<img src=\"/$SubDir/PQ_$ContextID.png?$Rand\" BORDER=0 USEMAP=\"#pqmap\" WIDTH=$XSize HEIGHT=$YSize>\n";
	print "</center>\n";
	print "<MAP NAME=\"pqmap\">\n";
	map {print "$_\n"} @ImageMapEntries;
	print "</MAP>\n";
	}

sub BorderStart
	{
	my ($Width) = @_;

	my $Wd = (defined $Width ? "WIDTH=\"$Width\"" : "");
	
	print "<TABLE BORDER=0 CELLPADDING=1 CELLSPACING=0 BGCOLOR=\"#000000\" $Wd><TR><TD>\n";
	print "<TABLE WIDTH=\"100%\" BORDER=\"0\" CELLSPACING=\"1\" CELLPADDING=\"3\">\n";
	print "<TR><TD BGCOLOR=\"$TableColor\">\n";
	}

sub BorderEnd
	{
	print "</TD></TR></TABLE></TD></TR></TABLE>\n";
	}

sub TableStart
	{
	print "<center>\n";
	BorderStart (@_);
	print "<Table BORDER=0 cellpadding=0 cellspacing=1 Width=\"100%\">\n";
	}


sub TableEnd
	{
	print "</Table>\n";
	BorderEnd ();
	print "</center>\n";
	}


sub TableHeader
	{
	my ($Fields) = @_;
	my ($Field);

	print "<tr bgcolor=\"$TableHeaderColor\">\n";
	foreach $Field (@$Fields)
		{
		print "<th align=center nowrap>".
	         "<b>&nbsp;$Field&nbsp;</b></font>".
		      "</th>\n";
		}
	print "</tr>\n";
	}


my @NewRow;
my @OldRow;
my $OddRow = 0;

# $Fields is a list ref of the strings to print
# $Formats is a listref of the format of the strings 
# a field format may be:
#
# 'S'	 	for a string
# 'R'	 	for a right justified string
# 'N#' 	for a number, the # is the number of places after the decimal
# 'M'	 	for money
# 'Lstr' for a link, str is the links string, use %s in the string to use the field data
#
# any of the formats may be preceeded by a '+' character.  This means don't print the field
# if it has the same value as the value in the previous row
#
# an example $Formats would be: ["+S","+S","S","S","N4","M"]
#
sub TableRow
	{
	my ($Fields, $Formats, $AltRow) = @_;
	my ($Field, $Fld, $Format, $OldField);

	my @LocalFormats = @$Formats;
	@NewRow = @$Fields;

	$OddRow = 1 - $OddRow;
	$AltRow = defined $AltRow ? $AltRow : $OddRow;
	 
	print "<TR " . ($AltRow ? "bgcolor=\"$TableAltRowColor\"" : "") . ">\n";
	foreach $Fld (@$Fields)
		{
		$Field = $Fld;	# make copy, don't change the actual data
		$Format 	 = shift @LocalFormats;
		$OldField = shift @OldRow;

		if ($Format =~ /^\+(.*)/)
			{
			$Field = "&nbsp" if ($Field eq $OldField);
#			$Field = "" if ($Field eq $OldField);
			$Format = $1;
			}
		if ($Format =~ /^L/) { print "<td>&nbsp;". &LinkFormat  ($Field, $Fields, $Format) . 		"&nbsp;</td>";}
		if ($Format =~ /^S/) { print "<td>&nbsp;". &StringFormat ($Field, $Format) . 			 		"&nbsp;</td>";}
		if ($Format =~ /^R/) { print "<td align=right>&nbsp;". &StringFormat ($Field, $Format) . 	"&nbsp;</td>";}
		if ($Format =~ /^M/) { print "<td align=right>&nbsp;". &MoneyFormat  ($Field, $Format) . 	"&nbsp;</td>";}
		if ($Format =~ /^N/) { print "<td align=right>&nbsp;". &NumberFormat ($Field, $Format) . 	"&nbsp;</td>";}
		if ($Format =~ /^X/) { }	# skip this column
		if (!$Format 		 ) { print "<td align=right>&nbsp;". &StringFormat ($Field, $Format) . 	"&nbsp;</td>";}
		}
	@OldRow = @NewRow;
	print "</TR>\n";
	}


sub MoneyFormat 
	{
	my ($Field, $Format) = @_;
	return "\$" . NumberFormat ($Field, "N2");
	}

sub NumberFormat
	{
	my ($Field, $Format) = @_;
	my $Decimals = substr ($Format, 1, 1) || '0';

	my $Str1 = sprintf ("%.".$Decimals."f", $Field);

   my $i = reverse $Str1;
   $i =~ s<(\d\d\d)(?=\d)(?!\d*\.)><$1,>g;
   return scalar reverse $i;
	}

sub LinkFormat
	{
	my ($Fld, $Fields, $Format) = @_;
	my $Data = substr ($Format, 1);

	return $Fld if (!$Fld || $Fld eq "&nbsp");

	my $i;
	for ($i=1; $i <= @{$Fields}; $i++)
		{
		$Data =~ s/\~$i/$Fields->[$i-1]/g;
		}
   return $Data;
	}


sub StringFormat
	{
	my ($Field, $Format) = @_;
	return $Field;
	}


sub StripNumberStuff 
	{
	my ($Num) = @_;

	$Num =~ tr/$,//d; # remove any existing commas

	return $Num;
	}

sub AddMapCircle
	{
	my ($x,$y, $PointSize, $Url, $Alt) = @_;
	my $Rad = $PointSize/2;
#	push @ImageMapEntries, "<AREA SHAPE=CIRCLE COORDS=\"$x,$y,$Rad\" HREF=\"$Url\" ALT=\"$Alt\">";
	unshift @ImageMapEntries, "<AREA SHAPE=CIRCLE COORDS=\"$x,$y,$Rad\" HREF=\"$Url\" ALT=\"$Alt\">";

	#The circle doesn't seem to support the popup alt tag if using Netscape
	#my $x1 = $x - $PointSize/2;
	#my $y1 = $y - $PointSize/2;
	#my $x2 = $x + $PointSize/2;
	#my $y2 = $y + $PointSize/2;
	#push @ImageMapEntries, "<AREA SHAPE=RECT COORDS=\"$x1,$y1,$x2,$y2\" HREF=\"$Url\" ALT=\"$Alt\">";
	}

sub AddMapRect
	{
	my ($x1,$y1, $x2, $y2, $Url, $Alt) = @_;
#	push @ImageMapEntries, "<AREA SHAPE=RECT COORDS=\"$x1,$y1,$x2,$y2\" HREF=\"$Url\" ALT=\"$Alt\">";
	unshift @ImageMapEntries, "<AREA SHAPE=RECT COORDS=\"$x1,$y1,$x2,$y2\" HREF=\"$Url\" ALT=\"$Alt\">";
	}

#############################################################################
#
#	General Utilities Section
#
#############################################################################

sub MyDie
	{
	my ($Msg, $Msg2) = @_;

   print "Content-type: text/html\n\n";
   print "<h1>ERROR</h1>\n";
   print "<hr><h3>$Msg</h3><hr>\n";
   print "$Msg2<hr>\n" if ($Msg2);
	CommonFooter ("ERROR");
   exit (0);
	}


sub ConnError
	{
	my ($Conn, $Str2) = @_;
	
	MyDie ("Errors: ". $Conn->Errors(0)->{Description} . "\n", $Str2) if ($Conn->Errors->{Count});
	MyDie ("Errors: Unknown connection error\n", $Str2);
	}


sub WebInit
	{
	# DataDir (in docs tree) stores graphs and dingbats
	$DataDir   = "$ENV{DOCUMENT_ROOT}/$SubDir";

	# ContextDir (in cgi-bin tree) stores context data and notes files
   ($ContextDir) =  $ENV{SCRIPT_FILENAME} =~ /^(.+\/)[^\/]+$/;
	$ContextDir .= $SubDir;
	CommonInit ();
	}

sub CommandLineInit
	{
	# We cannot use web server environment settings to determine these vars if
	# we are running from the cmd line.  So here I make them up
	$DataDir    = "c:/apache/htdocs/$SubDir";
	$ContextDir = "c:/apache/cgi-bin/$SubDir";
	CommonInit ();
	}


sub CommonInit
	{
	# create Dirs if needed
	(mkdir $DataDir    || MyDie ($!)) if (! -d $DataDir);
	(mkdir $ContextDir || MyDie ($!)) if (! -d $ContextDir);

	# create the dingbats if they don't exist
	CreateDingbats () if (! -e "$DataDir/V01.jpg");
	}



# This sub is called if the script is called without any
# extra path info.  This creates a new context and does a reloc
#
sub NewContext
	{
	my $ContextID = 0;

#	mkdir $ContextDir."/" || MyDie ($!);
	mkdir $ContextDir if (! -d $ContextDir);

	if (open (LASTCX, "<$ContextDir/$ContextCount"))
		{
		$ContextID = <LASTCX>;
		close LASTCX;
		}
	$ContextID += 1;
	open (LASTCX, ">$ContextDir/$ContextCount") or MyDie ("Cannot write Context controller file.\n");
	print LASTCX $ContextID;
	close LASTCX;

	# start with default.dat if it exists
	if (open (DEF, "<$ContextDir/default.dat") && open (NEW, ">$ContextDir/$ContextID.dat"))
		{
		while (<DEF>)
			{
			print NEW $_;
			}
		close (DEF);
		close (NEW);
		}

	my $NewLoc = "http://" . $ENV{"SERVER_NAME"} . $ENV{"SCRIPT_NAME"} . "/" . "$ContextID/";
   print "Location: $NewLoc\n\n";
	exit (0);
	}

# ReadContext
# This fn reads the context file into %Context
#
# This file also adds any GET parameters and then writes the Context 
# back to disk.  Therefore, any changes made to %Context after this
# fn call do not get saved.
#
sub ReadContext
	{
	my $ContextID = shift;

	if (open (CX, "<$ContextDir/$ContextID.dat"))
		{
		while (<CX>)
			{
			chomp;
			@_ = split "=";
			$Context{$_[0]} = $_[1];
			}
		close CX;
		}

	# also get the GET parameters
	my @Params = split ("&", $ENV{"QUERY_STRING"});

	# the next dozen lines or so are to handle the "include vendors" form 
	# from the item profile page.  It's looks like a hack, and it is.
	my $i = 0;
	my $j = 0;
	foreach (@Params)
		{
		@_ = split "=";
		$i += ($_[0] eq "c00");
		$j += ($_[0] =~ /c[0-9][0-9]/);
		}
	if ($i || !$j)
		{
		for ($i=0; $i<20;$i++)
			{
			my $ref = sprintf("V%2.2d",$i);
			$Context{$ref} = ($j ? "" : 1);
			}
		}

	return if ! @Params; 
	foreach (@Params)
		{
		@_ = split "=";
		$Context{$_[0]} = $_[1];
		}
	WriteContext ($ContextID);
	}


sub WriteContext 
	{
	my $ContextID = shift;
   my $key;

	open (CX, ">$ContextDir/$ContextID.dat") or MyDie ("Cannot open context file $ContextDir/$ContextID.dat\n");

   foreach $key (sort keys %Context) 
      {
      print CX "$key=$Context{$key}\n" if $key;
      }
	close CX;
	}

sub max
	{
	return ($_[0] > $_[1] ? $_[0] : $_[1]);
	}

sub min
	{
	return ($_[0] < $_[1] ? $_[0] : $_[1]);
	}


sub DebugPage
   {
   my $key;

	CommonHeader ("Debug Dump Page");

	print "<H2>User Settable (GET) Parameters:</H2>";
	print "DataSource=str (CRAIGF or DSSNTS, default is CRAIGF) [currently set to $Context{DataSource}]<br>\n";
	print "Catalog=str	(dss is only choice) [currently set to $Context{Catalog}]<br>\n";
	print "Cube=str	(newprofile is only choice) [currently set to $Context{Cube}]<br>\n";
	print "Debug=0 or 1 (default is 0) [currently set to $Context{Debug}]<br>\n";
	print "UseVB=0 or 1 (default is 0) [currently set to $Context{UseVB}]<br>\n";
	print "CenterBaseBid=0 or 1 (default is 0) [currently set to $Context{CenterBaseBid}]<br>\n";
	print "ColSize=# (default is to auto calc based on # of items) [currently set to $Context{ColSize}]<br>\n";
	print "LIPPointSize=#	 (default is 7 pixels wide) [currently set to $Context{LIPPointSize}]<br>\n";
	print "PreferMDX=0 or 1 (default is 0) [currently set to $Context{PreferMDX}]<br>\n";
   print "<hr>\n";

	my ($Context, $Target) = (split "/", $ENV{"PATH_INFO"})[1..2];
	print "<H2>Current Context Control</H2>";
	print "Context ID = $Context<BR>\n";
	print "Target = $Target<BR>\n";
   print "<hr>\n";

	print "<H2>Context</H2>";
   foreach $key (sort keys %Context) 
      {
      print "<B>$key</B> ->$Context{$key}<br>\n";
      }
   print "<hr>\n";

	print "<H2>Environment</H2>";
   foreach $key (sort keys %ENV) 
      {
      print "<b>$key</b> -> $ENV{$key}<br>\n";
      }
   print "<hr>\n";
	CommonFooter ();
   exit (0);
   }


# for debugging, only called if Debug is true
sub DumpProfileData
	{
	my ($Data) = @_;
	my ($Item, $BidderID);

	open (DAT, ">c:\\profile.txt") || MyDie ("Cannot open profile dump file c:\\profile.txt");

	foreach $Item (@$Data)
		{
		print DAT "Item: $Item->{ID}\n";
		foreach $BidderID (sort keys %{$Item->{Bids}})
			{
			print DAT "   Bid: $BidderID = $Item->{Bids}->{$BidderID}->{Dev}\n";
			}
		print DAT "   Est: Estimate = $Item->{Estimate}->{Dev}\n";
		}
	close DAT;
	}


sub DumpHash
	{
	my ($HRef) = @_;
	my $Key;

	print "=============\n";
	foreach $Key (sort keys %$HRef)
		{
		print "$Key => $HRef->{$Key}\n";
		}
	print "=============\n";
	}

#############################################################################
