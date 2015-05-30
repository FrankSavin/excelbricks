<table border='1'>
<a href='Hidden comment: 
<tr><td>	[version number here]	

Unknown end tag for </td>

<td><ul><li>*[feature here]*<li>*[feature here]
'></a><br>
<tr><td>	v7.7,1</td><td><ul>
<li>!Form - Fixed for untrimmed strings that can be converted to numbers<br>
<tr><td>	v7.7</td><td><ul>
<li>Documentation - !Set applies to range<br>
<li>Licensing - Changed license to MIT license<br>
<li>!Form - Use "Escape" key to close form<br>
<li>General - Changed :r: to <code>[r]</code>
<li>General - ExcelBricks internal functions can be used externally<tr><td>	v7.6</td><td><ul>
<li>Documentation - Remove document properties<br>
<li>Documentation - Command index should have sub-items<br>
<li>Documentation - Change "Extracted data" references to "Sample Sheet"<br>
<li>Documentation - Command Syntax section link formatting changed<br>
<li>Documentation - Fixed table formatting in Configuration1<br>
<li>Usability - Changed sample module error handlers to one-liners<br>
<li>Design - Moved user error numbers to user range 513+<br>
<li>General - Last run time set properly in standard template<br>
<li>!Clear - Fix for filtered sheet<br>
<li>!Extract - Fix for .xslx files<br>
<li>!Extract - Fix for junk characters at the end when input > 255 characters from excel<br>
<li>!Extract - Support for MS Access databases<br>
<li>!Form - Fix for Label wrap rows alignment to form bottom<br>
<li>!Form - Fix for Pulldown source input shift when row <> 1<br>
<li>!Form - Shows proper error message for unrecognized control type<br>
<li>!Set - Added "No screen refresh" configuration variable<tr><td>	v7.5	</td><td><ul><li>Renamed to remove space between Excel and Bricks<li>Changed demo form name to ExcelBricks demo<li>Updated formatting to avoid loss of fidelity warning while saving in Excel 2007<li>!form - form parameters don't show with form name in summary log<li>!form - long labels are wrapped instead of being clipped off<br>
<tr><td>	v7.4	</td><td><ul><li>Moved version history to site<li>Documentation - Removed references to old module name<li>Documentation - Updated for changes, and corrected typos, reformatted Change History, updated content for clarity<li>Documentation - Changed examples to refer to Sheet1 (existing) as much as possible<li>Documentation - Corrected hyperlinks (offset)<li>Documentation - Added label text clip-off as known issue<li>!Form - Destination not mandatory if there are no controls that get a value<li>!Form - User friendly error message for missing destination<li>!Form - Blank lines are interpreted as blank lines<li>Added a Demo as the default configuration<li>Code comments - corrected outdated comments<li>Configuration - updated labels for Configuration1 to match those in documentation<li>Documentation - Program parameters section updated to be more clear<li>Documentation - Included copyright and licensing<li>Documentation - reformatted change history and Syntax<li>Documentation - Add a note for macro security in Excel 2007<li>Documentation - Add a note explaining how to add a button and add onclick event code<br>
<tr><td>	v7.3	</td><td><ul>	<li>!Extract - SQL Server instances work properly now<li>!Extract - Clean up does not error out on connection failure<br>
<tr><td>	v7.2	</td><td><ul>	<li>Name changed to Excel Bricks<br>
<tr><td>	v7.1	</td><td><ul>	<li>!Debug - pause seconds now accepted<li>!Extract - worked around memory leak on Excel self-reference<li>Screen updates enabled during message box displays<li>Code cleanup<br>
<tr><td>	v7.0	</td><td><ul>	<li>!Debug - Status bar cleared only if !Debug was used<li>!Form - major syntax changes<li>!Form - code cleanup<li>!Form - radio and pulldown list locations are part of the label<li>!Form - mandatory fields are indicated by a <code>*</code> at the end of the label text<li>!Form - data target location is now at form level<li>!Form - default values are pulled from the target automatically<li>!Set > Run Message - Fixed capitalization<li>!Set > Run Timestamp Address - Fixed worksheet names with quotes<li>!Set - can now set a cell to a value<li>Template - updated to build a new module more easily<li>Syntax - Added Introduction and Contents<li>Syntax - Added index for commands<li>Syntax - Content update for clarity<br>
<tr><td>	v6.2	</td><td><ul>	<li>Changed If_True_Skip_n_Rows to OnTrueMove<li>OnTrueMove traces last executed real command properly<li>Negative steps for OnTrueMove is level 1 indexed<li>Swapped parameters for !Msgbox<li>Icon for StopIfBlank changed<li>!Form now supports mandatory fields<li>Fixed error when an excel extract comes between 2 SQL !Extracts to the same database<li>Fixed error for :r: value evaluation in !Form when there are 2 buttons<li>Option to self adjust height in !Form<li>Option to set width of label and control section in !Form<li>!ConvertFormulaToValues now works immediately after !Filter<li>Fixed clearing of sheet for !Extracts (now based on SQL, not header row)<li>Corrected message position on closing !Form<li>Added !Debug statement<li>!Set sets configuration variables<li>Change excel !Extract command format<li>!Extract suports multiple database:user:password sets<li>Renamed Configuration 1 sheet to Configuration1 to keep it simple<li>Added note about being able to use A1 notation for formulas to Syntax<li>Updated formatting on Syntax sheet<li>Removed native support for MS Access to SQL Server conversion<li>Don't show checkbox for password box on external call<li>Code cleanup<li>:r: conversion is now at the range level<br>
<tr><td>	v6.1	</td><td><ul>	<li>Fixed !Clear function not working for specific ranges<li>Fixed !ConvertFormulaToValues function not working for specific ranges<li>Add Template Function and Template Sub<li>Removed Sample module<li>Added Don't ask me again this run to database password prompt screen and removed retain data parameter<li>Corrected form caption parameter functionality<li>Label width with and without controls<li>Default value for radio button set is now the first one<br>
<tr><td>	v6.0	</td><td><ul>	<li>Move version history to different worksheet<li>Add sheet for Syntax<li>Change !Validate to StopIfBlank<li>Removed redundant calculate before ConvertFormulaToValues<li>Removed multiple messages for ConvertFormulaToValues<li>Generic form is resizable<li>No limit on number of form controls<li>Pulldown need only have a start cell instead of a range.<li>Add checkbox to custom form<li>Add radio button to custom form<li>Removed skip run summary dialog option for StopIfBlank<li>Default values for custom form fields<li>Fixed !Clear method for worksheet level<li>Removed row number for StopIfBlank failure message<br>
<tr><td>	v5.12	</td><td><ul>	<li>Extract excel checks for potential data truncation<li>Moved version history to separate table<li>!Run function should support relative references<li>Option to change name of button on custom form<li>Self reference for same workbook<li>Fixed mixed type error for import from Excel<li>Filter out excel extract self reference to avoid password prompt.<li>Changed password<br>
<tr><td>	v5.11	</td><td><ul>	<li>Add extract from excel sheet<li>Add :r+n: or :r-n: or :row: to denote reference relative to current row<li>Add pulldown to custom form<li>Add dialog box with Yes and No for program flow<li>Changed !Skip_n_Rows_if_True keyword for consistency - swapped parameters<li>Changed called subroutine name to sRunProgram<li>Make commands case insensitive<li>Fix !Extract defect for specific location with sheet name within single quotes<li>Show line number for errors<li>Changed default sheet to have no configured database password<li>!Version command<li>Changed code password to opensesame<br>
<tr><td>	v5.10	</td><td><ul>	<li>Update function and subroutine syntax<li>Fix program flow when it encounters 2 skip rows and only the 2nd fails<li>Add message for skip row<li>Unrecognized command gives row number<br>
<tr><td>	v5.9	</td><td><ul>	<li>Add !Filter statement<li>Renamed intQueryColumn in code<li>Updated !Run function to support return values<li>Merged Wrapper module into Sample module<br>
<tr><td>	v5.8	</td><td><ul>	<li>Add check for if sheet is shared<li>Rename Control Panel and Configuration and Button to be more related to each other<li>Issue with retention when a preceding run errors out because target range is invalid<li>Add !Skip_n_Rows_if_True statement<br>
<tr><td>	v5.7	</td><td><ul>	<li>Updated configuration sheet labels<br>
<tr><td>	v5.6	</td><td><ul>	<li>Make Change History maintenance formatting easier by including dummy row<li>Included wrapper for sample function call<br>
<tr><td>	v5.5	</td><td><ul>	<li>Clear out dynamic control values to prevent retention<li>Fixed ConvertFormulaToValues error when worksheet name has spaces<li>Rolled back copyfromrecordset to fix calculation of formulas<li>Added Extract keyword<li>Removed Calculate keyword<br>
<tr><td>	v5.4	</td><td><ul>	<li>Error out for unrecognized commands<li>Add version to code for Forms<li>Force form repaint after hiding<li>Updated file properties<br>
<tr><td>	v5.3	</td><td><ul>	<li>Fix for sheet names with spaces<li>Error message for validate statement and skip temporary status<li>Added versioning for syntax<li>Added cleanup for worksheet objects<li>Added dummy statement<li>Fixed dragging window screen painting issue<br>
<tr><td>	v5.2	</td><td><ul>	<li>Fall through error handling for called subroutines<br>
<tr><td>	v5.1	</td><td><ul>	<li>Initialize global variables to avoid a run from interfering with subsequent runs<br>
<tr><td>	v5.0	</td><td><ul>	<li>Moved base code to module<li>Retain user configurations in control sheet<li>Add Run statement<li>Make final status configurable<li>Make connection paramenter fallthrough configurable<li>Changed name of template<li>Add Form statement<li>     Label<li>     Textbox<li>     Password<li>Fixed control focus<li>Added Validate statement<br>
<tr><td>	v4.8	</td><td><ul>	<li>Add clear statement<br>
<tr><td>	v4.7.1	</td><td><ul>	<li>Code review and cleanup<li>     Optimized clear data code<li>     Cleaned up comment formula code<li>     Removed user configuration data from code<li>Added logging for all command types<li>Changed message box styles<li>Fixed stop clock code<br>
<tr><td>	v4.7	</td><td><ul>	<li>Repaint before and after showing forms<li>Update status message titles<li>Exits if user name is blank<li>Compensate for data entry time in run time<br>
<tr><td>	v4.6	</td><td><ul>	<li>Set formulas using range<li>Set format for comment cascaded formula cells<li>Updated input form<br>
<tr><td>	v4.5	</td><td><ul>	<li>Speed up calculated values<br>
<tr><td>	v4.4	</td><td><ul>	<li>Keyword to convert formulas to values (at worksheet and range level)<li>Keyword to refresh pivot tables in a sheet<li>Include run time in final message<br>
<tr><td>	v4.3	</td><td><ul>	<li>Added shortcut for Reload button<li>Force calculation<li>Retain username password<li>Make Configuration sheet configurable<li>Added version in code<li>Added last updated cell<li>Made user specific configuration separate<li>     For autofit<li>     For default username / password<li>     Last updated date-time target<li>Redirect query output to specific range<br>
<tr><td>	v4.2	</td><td><ul>	<li>Added state check for Connection and recordset before closing out<br>
<tr><td>	v4.1	</td><td><ul>	<li>Remove autofilter before deleting data<br>
<tr><td>	v4.0	</td><td><ul>	<li>Continue on previous sheet if consecutive output sheets are the same<br>
<tr><td>	v3.4	</td><td><ul>	<li>Password and username entry<li>Clear contents only<br>
<tr><td>	v3.2	</td><td><ul>	<li>User lookup. <li>Formula columns.<li>Optimization.<br>
<tr><td>	v3.1	</td><td><ul>	<li>Intermediate - till query is finalized and user lookup is done<br>
<tr><td>	v3.0	</td><td><ul>	<li>Multiple databases. <li>Hide configuration and User Lookup.<li>Protected workbook.<br>
<tr><td>	v2.0	</td><td><ul>	<li>Updated for multiple queries.<li>Protected code.<br>
<tr><td>	v1.0	</td><td><ul>	<li>Initial version