# The Mendeley OpenOffice Extension

This extension provides integration between Mendeley Desktop and OpenOffice/LibreOffice,
providing the ability to insert citations from your Mendeley library into OpenOffice documents
and generated a bibliography automatically.

## Prerequisites to build:

 * Perl
 * 7za (Windows) or zip (Linux or OSX)

## Instructions to build the extension:

 1. Open a command prompt within the directory containing `buildScript.pl`
 2. Run `buildScript.pl <version> <debug mode>`
   * `<version>`: Version number to use for this plugin build.
   * `<debug mode>`: Boolean which specifies whether the debugger should be enabled in OpenOffice.

	e.g. `buildScript.pl 1.5 true` to make version *1.5* as a *debug build*

(Note: Debug builds don't use custom error handling code, but instead OpenOffice/LibreOffice will open the debugger)

## Instructions to run unit tests:

 1. Build the .oxt file using the steps above
 2. Install the extension in OpenOffice Writer (via Tools -> Extension Manager)
 3. Set the environment variable `MENDELEY_OO_TEST_PATH` to the full path of the `tests/` directory
 4. Copy the `tests/testDatabase@test.com@local.sqlite` file into your Mendeley data directory (see http://www.mendeley.com/faq/#locate-database).
 5. Run Mendeley Desktop with options `--account testDatabase@test.com --server local`
 6. Run OpenOffice Writer and select Tools->Macros->Run Macro... 
 7. From the tree choose My Macros->Mendeley->mendeleyUnitTests, and then choose the macro named `runUnitTests()`
 8. Click Run (currently it will only show a message if something goes wrong)

(Tip: it's handy to set up a toolbar button using Tools->Customise if you run the tests frequently)
