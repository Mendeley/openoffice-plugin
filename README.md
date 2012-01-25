*__Please Note:__  This won't work with the stable version of Mendeley Desktop. We're working on it, so please check back later.*

*We are in the process of opening up the code and API for integration between Mendeley Desktop and word processors. The stable version of the plugin is bundled with Mendeley Desktop and can be installed via Tools -> Install OpenOffice Plugin.*

# The Mendeley OpenOffice Extension

This extension provides integration between Mendeley Desktop and OpenOffice/LibreOffice,
providing the ability to insert citations from your Mendeley library into OpenOffice documents
and generated a bibliography automatically.

## Build prerequisites:

 * Perl
 * 7za (Windows) or zip (Linux or OSX)

## Building the extension:

 1. Open a command prompt within the directory containing `buildScript.pl`
 2. Run `buildScript.pl <version> <debug mode>`
   * `<version>`: Version number to use for this plugin build.
   * `<debug mode>`: Boolean which specifies whether the debugger should be enabled in OpenOffice.

	e.g. Run `buildScript.pl 1.5 true` to make version *1.5* as a *debug build*

(Note: Debug builds don't use custom error handling code, but instead OpenOffice/LibreOffice will open the debugger)

## Installing the extension:

 1. Build the .oxt file using the steps above
 2. Start OpenOffice Writer, go to Tools -> Extension Manager, click the 'Add' button and select the generated .oxt file.

## Running unit tests in OpenOffice:

 1. Build and install the .oxt file using the steps above
 2. Set the environment variable `MENDELEY_OO_TEST_FILES` to the full path of the `testFiles/` directory
 3. Copy the `tests/testDatabase@test.com@local.sqlite` file into your Mendeley data directory (see http://www.mendeley.com/faq/#locate-database).
 4. Run Mendeley Desktop with options `--account testDatabase@test.com --server local`
 5. Run OpenOffice Writer and select Tools->Macros->Run Macro... 
 6. From the tree choose My Macros->Mendeley->mendeleyUnitTests, and then choose the macro named `runUnitTests()`
 7. Click Run (currently it will only show a message if something goes wrong)

(Tip: it's handy to set up a "Run Unit Tests" toolbar button using Tools->Customise if you run the tests frequently)

## Running other unit tests:

 Prerequisite: Python 2.6 or 2.7

### Non-interactive tests

 1. Ensure Mendeley Desktop is running
 2. Run "python src/MendeleyHttpClient\_test.py"
 3. Run "python src/MendeleyDesktopAPI\_test.py"

### Interactive tests

 1. Ensure Mendeley Desktop is running
 2. Run "python src/MendeleyDesktopAPI\_test\_interactive.py"
    (this will prompt you for manual input and produce output based on your actions)

