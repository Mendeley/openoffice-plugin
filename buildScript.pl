# BEGIN LICENSE ECL 1.0
# This Educational Community License (the "License") applies
# to any original work of authorship (the "Original Work") whose owner
# (the "Licensor") has placed the following notice immediately following
# the copyright notice for the Original Work:
# 
# Copyright (c) 2009 Mendeley Ltd.
# 
# Licensed under the Educational Community License version 1.0
# 
# This Original Work, including software, source code, documents,
# or other related items, is being provided by the copyright holder(s)
# subject to the terms of the Educational Community License. By
# obtaining, using and/or copying this Original Work, you agree that you
# have read, understand, and will comply with the following terms and
# conditions of the Educational Community License:
# 
# Permission to use, copy, modify, merge, publish, distribute, and
# sublicense this Original Work and its documentation, with or without
# modification, for any purpose, and without fee or royalty to the
# copyright holder(s) is hereby granted, provided that you include the
# following on ALL copies of the Original Work or portions thereof,
# including modifications or derivatives, that you make:
# 
# 
# The full text of the Educational Community License in a location viewable to
# users of the redistributed or derivative work.
# 
# 
# Any pre-existing intellectual property disclaimers, notices, or terms and
# conditions.
# 
# 
# Notice of any changes or modifications to the Original Work, including the
# date the changes were made.
# 
# 
# Any modifications of the Original Work must be distributed in such a manner as
# to avoid any confusion with the Original Work of the copyright holders.
# 
# THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND,
# EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF
# MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT.
# IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY
# CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT,
# TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE
# SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
# 
# The name and trademarks of copyright holder(s) may NOT be used
# in advertising or publicity pertaining to the Original or Derivative
# Works without specific, written prior permission. Title to copyright in
# the Original Work and any associated documentation will at all times
# remain with the copyright holders.

# prerequisite tools:
# 	perl (obviously)
# 	zip command line tool 7za.exe (Windows) or zip (Linux)
# END LICENSE ECL

# Other files/directories that the Script assumes that are available
#   mendeleyEmptyExtension.oxt/ (saved as a file for SVN tracking, would be a .oxt file that it's a zip)

# input: 
#   mendeleyMain.vb
#   mendeleyLib.vb
#   mendeleyDataTypes.vb
#   mendeleyUnitTests.vb

# output:
#   Mendeley-$PLUGIN_VERSION.oxt  # Plugin for OpenOffice (Windows, Mac, Linux)

use File::Copy;
use File::Path;
use Config;

mkdir "temp";

$SEVENZIP_LOCATION="7za.exe";

$PLUGIN_VERSION = $ARGV[0];
$DEBUG_MODE = $ARGV[1];

print "OpenOffice Plugin Version: $PLUGIN_VERSION\n";

processSourceFile("mendeleyMain.vb", "mendeleyMain-OpenOffice.vb", "Mendeley");
processSourceFile("mendeleyLib.vb", "mendeleyLib-OpenOffice.vb", "MendeleyLib");
processSourceFile("mendeleyDataTypes.vb", "mendeleyDataTypes-OpenOffice.vb", "MendeleyDataTypes");
processSourceFile("mendeleyUnitTests.vb", "mendeleyUnitTests-OpenOffice.vb", "MendeleyUnitTests");
processSourceFile("shared/sharedLib.vb", "sharedLib-OpenOffice.vb", "SharedLib");

# create OpenOffice mendeleyPlugin.oxt
# (which is actually a zip archive)
mkdir "Mendeley";
copy("temp/mendeleyMain-OpenOffice.vb", "Mendeley/mendeleyMain.xba")
	or die "copy failed: $!";
copy("temp/mendeleyLib-OpenOffice.vb", "Mendeley/mendeleyLib.xba")
	or die "copy failed: $!";
copy("temp/mendeleyDataTypes-OpenOffice.vb", "Mendeley/mendeleyDataTypes.xba")
	or die "copy failed: $!";
copy("temp/mendeleyUnitTests-OpenOffice.vb", "Mendeley/mendeleyUnitTests.xba")
	or die "copy failed: $!";
copy("temp/sharedLib-OpenOffice.vb", "Mendeley/sharedLib.xba")
	or die "copy failed: $!";
	
# TODO: refactor the copy commands into a function or use the Perl ones
# system ("svn export MendeleyEmptyExtension.oxt MendeleyEmptyExtensionTemp.oxt");
if($Config{osname} eq "MSWin32")
{
	system ("xcopy /E /I /Y MendeleyEmptyExtension.oxt MendeleyEmptyExtensionTemp.oxt");
}
else
{
	system ("cp -r MendeleyEmptyExtension.oxt MendeleyEmptyExtensionTemp.oxt");
}

if (not chdir("MendeleyEmptyExtensionTemp.oxt/"))
{
	if($Config{osname} eq "MSWin32")
	{
		system ("xcopy /E /I /Y MendeleyEmptyExtension.oxt MendeleyEmptyExtensionTemp.oxt");
	}
	else
	{
		system ("cp -r MendeleyEmptyExtension.oxt MendeleyEmptyExtensionTemp.oxt");
	}
	
	chdir("MendeleyEmptyExtensionTemp.oxt/") or	die ("Couldn't svn export or copy MendeleyEmptyExtension.oxt");
}

open(FP_DESCRIPTION_ORIG,"description.xml") || die ("Could not open description.xml");
open(FP_DESCRIPTION_NEW, ">description.xml.new") ||die ("Could not open description.xml.new");

while ($line = <FP_DESCRIPTION_ORIG>)
{
	$line =~ s/%PLUGIN_VERSION%/$PLUGIN_VERSION/;
	print FP_DESCRIPTION_NEW $line;
}
close (FP_DESCRIPTION_ORIG);
close (FP_DESCRIPTION_NEW);
move("description.xml.new", "description.xml");

if ((substr($Config{osname},0,5) eq "linux") || (substr($Config{osname},0,6) eq "darwin"))
{
  system("zip -r -q MendeleyPlugin.oxt .") == 0 or die("call to zip failed"); # TODO: exclude .svn files/directories
}
else
{
  system("\"$SEVENZIP_LOCATION\" a -tzip MendeleyPlugin.oxt . -x!icons\\.svn -x!.svn -x!Mendeley\\.svn -x!META-INF\\.svn -x!Office\\.svn -x!Office\\UI\\.svn -x!pkg-desc\\.svn -x!Scripts\\.svn");
}
move("MendeleyPlugin.oxt", "..");
chdir("..");

if ((substr($Config{osname},0,5) eq "linux") || (substr($Config{osname},0,6) eq "darwin"))
{
  system("zip -r -q MendeleyPlugin.oxt Mendeley") == 0 or die("call to zip failed");
}
else
{
  system("\"$SEVENZIP_LOCATION\" a -tzip MendeleyPlugin.oxt Mendeley\\");
}
rmtree("Mendeley") or die "rmtree failed: $!";
rmtree("MendeleyEmptyExtensionTemp.oxt") or die "rmtree failed: $!";
move("MendeleyPlugin.oxt","Mendeley-$PLUGIN_VERSION.oxt") or die "move failed: $!";

close(OPEN_OFFICE_MAIN);
close(OPEN_OFFICE_LIB);
close(OPEN_OFFICE_DATATYPES);
close(OPEN_OFFICE_UNIT_TESTS);

rmtree("temp") or die "rmtree failed: $!";

print "Mendeley OpenOffice plugin built successfully: version $PLUGIN_VERSION";
exit;

sub processSourceFile
{
	#arguments:
	my $source = $_[0];
	my $openOffice = $_[1];
	my $moduleName = $_[2];

	open(SOURCE_FILE, $source) or die "Couldn't open file $source";
	open(OPEN_OFFICE, ">temp/$openOffice") or die "Couldn't open file $openOffice";

	# open office header
	print OPEN_OFFICE "<?xml version=\"1.0\" encoding=\"UTF-8\"?>\n";
	print OPEN_OFFICE "<!DOCTYPE script:module PUBLIC \"-//OpenOffice.org//DTD OfficeDocument 1.0//EN\" \"module.dtd\">\n";
	print OPEN_OFFICE "<script:module xmlns:script=\"http://openoffice.org/2000/script\" script:name=\"$moduleName\" script:language=\"StarBasic\">\n";
	
	$line = <SOURCE_FILE>;

	while($line)
	{
		$line =~ s/\${DEBUG_MODE}/$DEBUG_MODE/g;
		
		# replace characters for openoffice (not sure why, but this is no longer required)
		$line =~ s/&/&amp;/g;
		$line =~ s/</&lt;/g;
		$line =~ s/>/&gt;/g;
		
		print OPEN_OFFICE $line;
		$line = <SOURCE_FILE>;
	}

	# open office footer
	print OPEN_OFFICE "\n\n</script:module>\n";
	close OPEN_OFFICE;
}