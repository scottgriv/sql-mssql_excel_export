-- Author: Scott Grivner
-- Website: linktr.ee/scottgriv
-- Abstract: Export Microsoft SQL Server to Excel

-- xp_cmdshell is a system stored procedure in SQL Server. It allows executing Windows shell commands from the SQL Server environment. While commands are passed as an input string, the shell's output is returned as rows of text.
-- Change -S to your DB server and -d to your Database, then change the query to what you want.
-- If you want to hide the column headers, add the flag: -h-1 after the .csv" syntax. The only issue is that the dashes are tied to the column header.
-- To remove the dashes but keep the header, do a UNION in your query with your column headers (see below).
-- 700 is the max width I set so you can adjust it if your columns are smaller/larger. This is for ALL columns.
-- I’ve also remove the “Rows Affected” output from the bottom of the file by the “no_output” switch at the end of the command
-- Adjust the -o flag to your desired output path and file name.

-- Hide Column Headers Example Query:
SELECT 'ID', 'Column_Name' FROM [Database_Name].[dbo].[Table_Name]
UNION
SELECT [ID]
      ,[Column_Name]
  FROM [Database_Name].[dbo].[Table_Name]

-- Examples:
EXEC  master..xp_cmdshell 'SQLCMD -S Server_Name -d Your_Directory -E -Q "set nocount on;SELECT * FROM [Database_Name].[dbo].[Table_Name]" -o "C:\temp\export\test.csv" -h-1 -s"," -w 700', no_output
EXEC  master..xp_cmdshell 'SQLCMD -S Server_Name -d Your_Directory -E -Q "set nocount on;SELECT * FROM [Database_Name].[dbo].[Table_Name]" -o "C:\temp\export\test.csv" -s"," -w 700', no_output
EXEC  master..xp_cmdshell 'SQLCMD -S Server_Name -d Your_Directory -E -Q "SELECT * FROM [Database_Name].[dbo].[Table_Name]" -b -o C:\temp\export\test.csv'
EXEC  master..xp_cmdshell 'SQLCMD -S Server_Name -d Your_Directory -E -Q "SELECT * FROM [Database_Name].[dbo].[Table_Name]" -o C:\temp\export\test.csv -s"," -w 700'
EXEC  master..xp_cmdshell 'SQLCMD -S Server_Name -E -Q "SELECT * FROM [Database_Name].[dbo].[Table_Name]" -b -o C:\temp\export\test.csv', no_output