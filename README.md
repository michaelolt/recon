# Preamble

## What's a revision?

It's July 1st. We have to tell the market and the networks how many kWh we believe were supplied to our customers in the month that has just been, June, so that they can invoice FOGY.

In the middle of July, we send new versions of the files with the June data to ensure changes to our data for that month are passed along to the market and networks. 

We'll do this again in September (R3), January (R7) and a final time the _next_ August (R14), always sending the June data to ensure any changes are passed along.

We do this for every month we are operating.  
That is to say, for each month, we will send five sets of data detailing the kWh supplied to our customers - R0, R1, R3, R7 and R14.

## Shorthand glossary
I'm gonna use some shorthand for programs instead of saying "In Microsoft Access, ..." constantly
   - `MA` - Microsoft Access
   - `EX` - Excel
   - `FZ` - FileZilla
   - `VS` - Visual Studio
   - `SQ` - An SQL stored procedure run through SSMS on the `FOGY` database
      - From the Object Explorer view in SSMS, expand `FOGY`, then `Programmability`, then `Stored Procedures`.
      - Find the right one, right click, and hit `Execute Stored Procedure`
   - `RM` - [The Reconciliation Manager Portal](https://www.electricityreconciliation.co.nz/), log in details in bitwarden
      - Use the `uat` site if you want to explore & play around


## TODOs and future 
1. Tony maintains his copy of `Calendar.xlsx` with due dates for each revision & highlights the next revisions due. If we start sharing the reconciliation process between us we should maintain this file in sharepoint rather than downloading it individually.

# Set up

## Tony's resources
To start, you need to download [all the bits Tony uses.](https://mpowerwork.sharepoint.com/sites/ForOurGood/Shared%20Documents/Forms/AllItems.aspx?isAscending=false&id=%2Fsites%2FForOurGood%2FShared%20Documents%2FReconciliation%2FHandover&sortField=Modified&viewid=f2dbcc41-1cdc-4275-a4a6-68dc25137b61)

Note that these resources are largely evergreen except for the `.bak` file inside `FOGY_SQL_DB.zip`.  

This needs refreshing with new data for new months. Tony has a tool (included in the Sharepoint folder linked above) to sync consumption & invoice data, but there are tables containing network prices, tou periods and other data that Tony updates himself, so while we are unclear on how to do that we should get a new copy of `FOGY.bak` when Tony makes those isolated changes.

## Applications
You need 
- Microsoft Access 
  - MS Access was a huge pain for me to get, I ended up 'acquiring' a copy. <-- Suss
- an instance of SQL Server
  - If you have ever run part of the OLT solution and used a local database, you will have this.
  - You'll need the name of the instance too. If you know it, great, otherwise I'll help you find it later.
- SSMS
- Visual Studio
- FileZilla

## Add FOGY.bak to your local SQL Server

[Instructions here!](https://dev.azure.com/imagimation/FOG/_wiki/wikis/Development/1954/How-to-open-up-Tony's-SQL-DB)

## New Data Source Name (DSN)
This is used by MS Access to connect to the database you just added.
- Start menu search for `ODBC Data Sources (64-bit)` and open it
- Open the `System DSN` tab
- Click `[Add...]`
- Select 'SQL Server' in the Driver popup dialogue
- Set the `Name` to 'FOGY' for tony's setup to work
  - `Description` is just for you
- If you are running the default SQL Server instance, just put a dot `.` in the Server box
- Try clicking Finish and then Test Data Source -
  - if `TESTS COMPLETED SUCCESSFULLY!`, you're done here!
  - If `TESTS FAILED!`, find the name of your SQL Server instance:
    - Start menu search for `SQL Server Configuration Manager` and open it
    - On the left, select `SQL Server Services`
    - In the name of the item called `SQL Server` there will be some text in brackets - that's the name of your instance. Put that in the Server box during DSN configuration.
      - If it says `(MSSQLSERVER)`, you are indeed running the default instance.

### Query timeouts
If you get this error trying to run anything in MS Access, you can increase the timeout by:
1. Right clicking the query under Queries in the All Access Objects panel
   - I believe all the queries behind the buttons have a Earth next to them and begin `qry_`
1. Selecting Design View
1. Editing the `ODBC Timeout` value (in seconds) in the Property Sheet
   - Using `0` as the value allows an infinite timeout
   - For the Read vs HHR query (`qry_ReadCheck`) for example, I needed 180 seconds.

## Opening the Access db
### Unblocking macros
Before you open `FOGY_Rec.accdb` for the first time, you gotta tell your computer it's ok.

Right click it and open its Properties dialogue.

Under the General tab, check Unblock, then hit Apply and OK.

### Hitting ctrl+s (don't)
For some reason every time I did this the program would freeze and I would have to download it again.

Don't be like me.

This may mean you have to make the following file path edit repeatedly - sorry.

_Having said that_, it appears changes you make to MS Access query timeouts won't apply unless you save in the design view.  
I don't know. I'll figure it out and update this wiki (lol).

## File paths
A few of tony's buttons create or read local files. This is what the folder structure in `Tony folders.zip` is for.  
Unzip these folders and put them somewhere you can find them.

In MS Access, you can edit the `CreateFiles` and `ImportFiles` 'Modules' at the bottom of the Objects panel by double clicking them.  

You can then Find & Replace All of:

`C:\Users\tonym\OneDrive\Desktop\FOGY`  

with your own path, e.g. mine:

`D:\Reconciliation\Tony folders`

## FOGY_Data_Import setup
Steps to edit parameters & connections to work for your specific machine.
1. `VS` Open `FOGY_Data_Import.sln`
1. `VS` If it tells you that a type of file is not supported (a `dtproj` file):
   1. `VS` Go to Extensions -> Manage Extensions
   1. `VS` Search for "SQL Server Integration Services Projects 2022" and install it
   1. This will download a setup exe
      - If running the exe tells you to close some mystery processes, close everything or restart your PC
   1. Completing the install with the exe will tell you to restart again, so do
   1. `VS` Open the `FOGY_Data_Import.sln` file again, right click the project in the solution explorer and hit `Reload Project`
1. `VS` If a popup tells you that a 'Provider' is not available or registered, you may need to [download this](https://www.microsoft.com/en-us/download/confirmation.aspx?id=50402)
1. `VS` Double click `Dataload_from_Prod_To_Loacal.dtsx` (sic) to open the design view
1. `VS` At the bottom, under Connection Managers, find `DESKTOP-I47MSDF\TAUMATA.FOGY`
1. `VS` Right click -> Edit
1. `VS` Enter `localhost` in the `Server name` box, Test Connection, and hit OK
1. `VS` Again under Connection Managers, find `fog.database.windows.net.fog-prod.tony`
1. `VS` Right click -> Edit
1. `VS` Enter the `User name` and `Password` of your own `fog-prod` database user in the boxes, Test Connection, and hit OK
   - Unsure whether this database user requires write access for the VS program to work

## Import_Sent_Files setup
Steps to edit parameters & connections to work for your specific machine.
1. `VS` Open `FOGY_Imports.sln`
1. `VS` Open `Import_Sent_Files.dtsx`
1. `VS` At the bottom, under Connection Managers:
1. `VS` Right click `FOGY`, -> Edit
1. `VS` Enter `localhost` in the `Server name` box, Test Connection, and hit OK
1. `VS` Right click `AV_File, -> Edit
1. `VS` Replace the tony path in `File name:` with yours, hit OK
   - Don't forget to end it with a `\`!
1. `VS` Open the Paramaters tab at the top 
1. `VS` Replace the tony paths in `ArchFolder` and `InputFolder` with yours
   - Don't forget to end it with a `\`!
1. `VS` Ensure you have 7Zip downloaded and installed at the location defined in the `ZipExe` param
   - If you don't have a C: drive or just don't want to move 7Zip, changing this value is ok too

# Reconciliation

Here we go!

## Data validation

Remember before starting to [sync the database.](https://dev.azure.com/imagimation/FOG/_wiki/wikis/Operations%20Manual/2047/Tony's-MS-Access-process?anchor=syncing-database)

### Load the right LIS file
1. `MA` Set the Reconciliation Month Start and End properties to the first and last of the month you are doing recon for
1. `MA` Set the Revision you are doing
1. `MA` Hit `Create REQ File`, it will appear in `To_RGST`
1. `FZ` Log in to the Registry SFTP, in bitwarden as `Electricity Registry - SFTP`
1. `FZ` Upload the REG file to the `toreg` folder
1. `FZ` Give it about 60 seconds and refresh the `fromreg` folder looking for a `LIS` file
1. `FZ` Move the `LIS` file to your `FROM_RGST` folder and **delete it** from the registry
1. `MA` Hit `Load LIS File`, wait for it to load
1. `MA` Hit the checkmark next to `Load LIS File` - this will check the loaded ICPs for any that aren't compliant
   - TODO what do we do about non compliant ICPs?

### Check missing data
1. `SQ` run `ret_Check_interval_Counts_Daily`, wait a few minutes for it to complete, copy results with headers
1. `EX` paste into the `Missing Data Report Template` 
1. `EX` Find & replace 'NULL' with '0'
1. `EX` Replace any zeroes with `ALT` where the zero is before the row's `StartDAte`
1. `EX` Identify and remove meter change rows:
   - Where there are two or more rows for an ICP and the good data clearly switches from one row to another:
   - ![image.png](/.attachments/image-fa2a10be-9f56-49ba-9c6b-8015dcdeb1e7.png)
1. `EX` Send the sheet to the estimation minions
1. ***Also:*** The query doesn't account for DST changeover days. 
   - For September, you should also check that the Interval47 and Intveral48 are `NULL` on the changeover day
   - For April, you should also check that the Interval49 and Intveral50 are **not** `NULL` on the changeover day

### Check Read vs HHR
1. `MA` Hit `Read vs HHR Check`
   - Anything over 10kWh, have a look
   - Over 5000, get worried
   - Negatives in the RegisterKWH column likely mean an RR is needed. -500 or higher need to be resolved at R3.

### Missing network variable prices
For any that are missing:

### Update the table
1. In the FOGY database `NetworkPrices` table, find the UN24 price for the same `Network` and `PriceCat`
   - If there is TOU pricing (e.g. PK,SH,OP) for the UN24 `RegisterMapping`, you'll need all those rows
   - **Note:** if using SSMS's 'Edit top X rows' to copy prices, be careful to edit the RegisterMapping right after you paste rather than committing the row and editing it after. Copying and pasting multiple rows at once will put you in this predicament, so copy and edit one at a time if you have to copy TOU pricing.
1. Duplicate the UN24 pricing rows within the table and change the `RegisterMapping` on the duplicated rows to the `RegisterMap` that was reported missing

### Update the stored procedures
Do this step for any piped (`|`) register maps that you add network prices for.

1. Find the `net_EIEP1` procedure and `Modify` instead of `Executing` to open the SQL it runs
1. Ctrl + F for `DIN16` to find several long `CASE` statements converting weirdo register maps to UN24
1. Add the `RegisterMap` that was reported missing if it isn't in there. Follow the convention set by each of the `CASE` statements.
1. Execute the SQL to update the stored procedure
1. Repeat for the `ret_Cost_Calculation_by_ICP` procedure

***Important:*** These changes are made to the `FOGY.bak` database. This means they will not be inherited by Tony or anyone else downloading the database to do recon. If we start sharing execution of recon, we need a process to disseminate these changes or replace the database in Sharepoint when it changes.

_Ideally_, we will edit the stored procedures & Access queries to run on the Dashboard database, and use that to do recon instead of Tony's database.

## Market file validation

It's now time again to [sync the database](https://dev.azure.com/imagimation/FOG/_wiki/wikis/Operations%20Manual/2047/Tony's-MS-Access-process?anchor=syncing-database), once data issues has been corrected.

1. `MA` Hit `Create AV090` and the other THREE buttons that create `AV` files
   - It might be a good idea to create, validate and fix the AV090 first. That one tends to be rejected the most.
1. `MA` Access will tell you where it puts them - in `/To_RM`
1. `RM` [Log in](https://www.electricityreconciliation.co.nz/), and from the left menu select `💼 File Manager` -> `File Checker`
1. `RM` In the top right, click `CHECK FILE`
   1. `UPLOAD FILE`, find the AV files you just created
   1. For each one, select the correct `File Type`:
      - `BILLED` = Electricity Supplied
      - `HHRAGGR` = Purchase HHR Aggregates
      - `HHRVOLS` = HHR Submissions
      - `ICPDAYS` = Purchaser ICP Days
   1. Submit. 
1. `RM` The file will validate, give it a second and press `REFRESH` to see the status
1. `RM` Click any that are `Rejected` to see the reasons why
1. `RM` _If_ any files are rejected, [fix them](https://dev.azure.com/imagimation/FOG/_wiki/wikis/Operations%20Manual/2047/Tony's-MS-Access-process?anchor=fixing-market-files)
   1. If you make changes _to the database_ to fix any AV files, recreate all four files
   1. Upload to the checker again (same file name is ok), repeat until all four are `Successful`
1. Once the file checker has confirmed your files are ok, move them into the `To_RM\To_Load_To_DB` folder
1. `VS` Open `FOGY_Imports.sln` (different from database sync sln)
1. `VS` Open `Import_Sent_Files.dtsx` & run it
   - This assumes you have done [the one-time setup](https://dev.azure.com/imagimation/FOG/_wiki/wikis/Operations%20Manual/2047/Tony's-MS-Access-process?anchor=import_sent_files-setup)
1. `MA` Open the design view of `qry_Revision_Variation_by_ICP` and edit the date in the ReconMonth column to be the first of the revision month. 
1. `MA` Close design view.
1. `MA` Double click `qry_Revision_Variation_by_ICP` to run it. 
   - This query uses data from the HHRAGGR file you just loaded to the database.
1. `MA` Copy the query results
1. `EX` Paste the results to the top left of the `Agg` sheet in the `Calendar.xlsx` file, replacing what may be there
1. `EX` Right click inside the Pivot table in columns M - O and hit 'Refresh'
   - ***Important note:*** If you are doing R0/Ri, you naturally won't have a previous revision to compare to, so you can ignore steps that mention a 'previous revision' or appear to be comparing two columns from the data you just pasted out of the query.
1. `EX` In the PivotTable Fields menu that slides out, select the number of the revision you are completing and then the number of the previous revision
   - e.g. if you are completing R3, Select `3` and `7`
1. `EX` Edit the formula in cell V4: Where the formula says `"Sum of [X]"`, change the first instance to your current revision number and the second instance to the previous revision number
   - e.g. for R3, it will read `=GETPIVOTDATA("Sum of 3",$M$1,"Direction","I","Network",M4)-GETPIVOTDATA("Sum of 7",$M$1,"Direction","I","Network",M4)`
1. `EX` Drag this change down to V21
1. `EX` Repeat for the formula in W4 - W21.
1. `EX` Copy & Paste the `Sum of [current R#]` columns for the I and X directions into the `I total` and `X total` columns to get the diff %
   - The diff % columns have a highlight rule to tell you when a diff is above the acceptable margin of error
1. **Explainer:** 
   - This will give you the overall difference in kWh between your current HHRAGGR file and what FOGY submitted to the market in the HHRAGGR file for the previous revision, for each Network & Direction.
   - Differences are expected, and are usually explainable. Tony has a threshold above which he will investigate (20k diff per 5m kWh, or 0.4%)
      - When an ICP withdraws, it may have data in the previous revision but not in the current one. This is **ok.**
      - When we get a backdated move in, the ICP may have data in the current revision but not in the previous one. This is **ok.**
      - When we get catch up data or otherwise modify Consumptions after including them in submitted market files, the data for the ICP will be different in the next revision. This is **ok.** 
      - These differences are all **ok** because they happened as a result of our data becoming more accurate.
   - We should be concerned with differences that point to issues with our processes.
      - Late withdrawals - seeing differences due to withdrawals at R3/7/14 is not ideal unless the circumstances made that withdrawal unavoidable.
   - Or differences that will raise flags for the RM
      - The RM can't see the causes of our big diffs not matter how reasonable, so if we concern them too often they might start to doubt our credibility
1. `EX` Any diffs that concern you can be investigated in the data you pasted into the left side of the `Agg` sheet:
   - You can deselect Networks to look at just one by clicking the arrow in the column header
   - You can filter out zeros in the Diff column:
      1. Click the arrow in the column header
      1. Select Number Filters -> Does Not Equal
      1. Enter '0' and hit OK
1. `SQ` Run `Submit_Return_Compare`
1. `SQ` Scroll down to find the recon month in the RecordPeriod column
1. `SQ` Copy the two numbers (I and X) for the current revision number
1. `EX` Paste these values into the `Agg` sheet so you can see if they are the same as in the Grand Total in the pivot table
   - This is a double check - this stored procedure runs on the HHRVOLS file data, so by checking the numbers there are the same as in the `Agg` sheet (which uses the HHRAGGR file data), we confirm that the rows in the two files sum to the same total

🎉 !You did it! 🎉

## Network file validation

1. `MA` Hit `Create All EIEP1 Files`
1. `MA` If you are doing R0 (i.e. it is currently the start of the month), hit `Create All EIEP4 Files`
1. `MA` Open the design view of `qry_EIEP1_KWH_Summary` and edit the date in the ReconMonth column to be the first of the revision month. 
1. `MA` Close design view, save changes.
1. `MA` Open the design view of `qry_HHR_Agg_Summary` and edit the date in the ReconMonth column to be the first of the revision month. 
1. `MA` Close design view, save changes.
1. `MA` Find the appropriate `qry_EIEP1_Agg_Compare_XX` query for the revision you are doing
1. `MA` Double click to run
1. `MA` Filter the results to get X rows only:
   1. `MA` Select one of the cells with an X or an I
   1. `MA` In the ribbon, click the big `Filter` button
   1. `MA` Deselect `I`
   1. `MA` Copy the results
   1. `EX` Paste the results into the `Agg` sheet at M27
1. `MA` Repeat for the I rows, pasting them into the `Agg` sheet at M50
1. `EX` Edit the formula in V28 to give `current revision EIEP - previous revision EIEP`, and drag down to V45
1. `EX` Edit the formula in W28 to give `current revision EIEP - HHRAGG`, and drag down to W45
1. `EX` Do the same for the set of data for the other Direction, below
1. **Interpreting:** 
   - `EIEP to HHRAGG` is the difference between the data in the network file and in the market file.
      - These should be small. Very small (single digit) differences are usually due to floating point rounding which is acceptable.
      - Some networks (e.g. LINE) will see bigger differences due to timesliced registers being estimated for all 48 intervals when they shouldn't be.
         - If those differences reach 300kWh or more in a given network, we need to go and fix the data. The network will notice.
   - `EIEP to EIEP` is the difference between the values submitted to the network in this revision and the previous one.
      - Differences here are expected because data changes between revisions, same as when validating Market files.
      - Differences here should also theoretically be the same as the differences in the Market revisions.
         - Market file diffs should still be in this same sheet so you can compare.
         - If they are, you can be confident you've explained the differences already while validating the Market files.
         - If they aren't, oh no. Tony didn't come across this while I was watching, so I don't know.

🎉 ! You did it again ! 🎉

Save a copy of the `Agg` sheet with the Market and Network file validation checks, in case we need to refer to this validation again.

## File submission

### Market
1. The market files will be in a .zip in the `To_RM\To_Load_to_DB\Archive` folder on your computer, after you loaded them to your database
      - is this how tony gets them..?
1. `FZ` Connect to the RM SFTP
   - Details in Bitwarden under `NZX Energy Market SFTP - WITS & Reconciliation`
1. `FZ` Upload the created files to the `reconciliation/to_rm` folder
1. `RM` Log in and check the Uploads page for your files. Ensure they are marked 'Successful'.

### Network
1. `MA` The Network files will be placed in the `To_Network` tony folder
1. `FZ` Connect to the Registry SFTP
1. `FZ` Drop all EIEP files in the `EIEP_OUT` directory
1. `FZ` The Registry will send back 2 sets of files in time, putting these in the `fromreg` directory
   1. Acknowledgement of receipt files
   1. Confirmation of send files
1. `FZ` Download these files, delete them from the SFTP, and store them in your `FROM_RGST` tony folder
1. There is one further step where Tony loads these return files into his db but he didn't show me that.

***R E C O N C I L I A T I O N&nbsp;&nbsp;&nbsp;S L A I N***

# Appendix

## Syncing database
Using Tony's program that copies our prod data into his recon database (`FOGY`).

You'll need to do this twice during this process:
1. At the very beginning
1. After you've corrected data issues in the `fog-prod` database highlighted during data validation

ok

1. `VS` Open `FOGY_Data_Import.sln`
1. `VS` Press Start to begin, and watch the progress!
   - This assumes you have done [the one-time setup](https://dev.azure.com/imagimation/FOG/_wiki/wikis/Operations%20Manual/2047/Tony's-MS-Access-process?anchor=fogy_data_import-setup)

## Fixing Market files
1. `"Checksum is not correct"` most likely means there are still data issues to fix (either missing intervals or too many intervals). 
   - Try running `ret_Check_interval_Counts_Daily` again and check over the results, or get Michael/Mathew to check them.
1. `"Submission at GXP0000-NETW-GN-D has incorrect number of records (X)"` means you need to either add or remove rows from the file, depending on what the `"Expected Y records"` error says.
   - If Y is greater than X, you need to add empty rows (zeroes) for missing dates until you have Y rows. This is what usually happens.
      - Here's a zeroes row you can copy in and edit as needed:
      - `GXP0000,NETW,GN,FOGY,HHR,LOSSCAT,DIRECTION(X/I),DEDICATED(Y/N),XX/XX/202X,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0`
   - If Y is less than X, you need to remove rows that are outside our contract dates for the GXP-LossCat-Direction in question.
   - If you add or remove rows in a market file, you need to correct the row count in the header of the file.  
Notepad++ gives you row numbers to make this easy. The number you use should exclude the header row.
1. `"Number of trading periods is not correct"` flows on from the previous fix! It means you need to add or remove zeroes from the rows you inserted that fall on DST changeover days 
   - Remove two zeroes for Septermber, add two zeroes for April.
1. `"Submission at GXP0000-NETW-GN-D has no contract and trading notification"` means what it says, Tony suggests - that there is no contract in place for the GXP. The RM manages these contracts so we shouldn't see this usually, but if you're seeing it ask for some assistance resolving.

Misc notes:
```
When we gain an ICP for a GXP & loss factor partway through the month and the ICP has a different value 
in the 'dedicated' field there will not be a row for each day of the month. 
Fill the missing days with zeroes. Count the rows (excluding the header) and update the row count. 

Can also happen when we get an ICP with generation, which adds another set of data with 'I' for the direction. 

Can ALSO happen when we lose the last ICP for a GXP/LF combo and have no data for the latter half of the month. 

There ARE cases where we should not have the whole month of days and should not fill with zeroes. 
This is when our contract only starts, or ends, partway through the month.
```
