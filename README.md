
About ods2sql
=============

I often have the problem that I have a table in OpenOffice/LibreOffice, but need to do deeper analysis than the software allows.

So, I wrote this script which generates a series of SQL statements that you can then import into your own database.

Note that this is highly experimental and I don't even know if I want to invest more time. Patches are welcome however :D

Usage
-----

Usage is pretty simple: Pipe the ODS file into STDIN, and SQL comes out of STDOUT:


    $ python ods2sql.py < data.ods | sqlite data.db

Every sheet in your ODS file becomes a table in your DB, the fields are named after the column names (A, B, C, ...)


Future plans
------------

Simply said, none. However, I have plans to implement the same functionality and more, in a generic way over at https://github.com/winged/wingeddata.

