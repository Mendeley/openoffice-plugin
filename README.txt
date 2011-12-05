The Mendeley OpenOffice extension.

Prerequisites to build:

- Perl
- 7za (Windows) or zip (Linux or OSX)

Instructions to build:

1. Open a command prompt within the directory containing buildScript.pl

2. Run buildScript.pl with arguments:
	argv[0] = version (string)
	argv[1] = use debug mode (boolean)

	e.g. "buildScript.pl 1.5 true" to make version 1.5 as a debug build

(NOTE: debug builds don't use custom error handling code, but instead OpenOffice/LibreOffice will open the debugger)
