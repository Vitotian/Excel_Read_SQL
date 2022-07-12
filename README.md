# Excel Read SQL
Read SQL Data with only ONE Function!

Check your Excel fits:
`2019+` and **X64** platform
NOTE: If EXCEL version is below `2019` you may use `Ctrl+Shift+Enter` to apply array function.
(NOT TEST YET)


## Installation
- copy folder to local (like `Mysql` to `D:\lib\vba_sql`)
- In EXCEL: Developer -> Excel Add-ins -> Browse (D:\lib\vba_sql\Mysql) -> `mysql_UTF8.xlam` and check on `mysql` -> Ok
- Enjoy

Or if you have own libary:
- import `xxx\xxx.bas` and `vba_list\List.cls` to your own project
- **change real lib path at `xxxx_Initialize()`**


## MySQL
usage:
```
=MSQLR("user:password@host:port/database","SELECT date,data from table",TRUE)
# Do not Start with `mysql://`
```


## PostgreSQL
usage:
```
=PSQLR("postgresql://user:password@host:port/database","SELECT date,data from table",TRUE)
# surport:
# postgres://localhost
# postgresql://localhost:5433
# postgresql://localhost/mydb
# postgresql://user@localhost
# postgresql://user:secret@localhost
# postgresql://other@localhost/otherdb?connect_timeout=10&application_name=myapp
# postgresql://host1:123,host2:456/somedb?target_session_attrs=any&application_name=myapp
```


## Sqlite
Codes are basically from [SQLiteForExcel](https://github.com/govert/SQLiteForExcel), but keep `Read` related functions.
usage:

```
=SQLR("C:\data\data.db","SELECT date,data from table",TRUE)
```



## Encoding Issue
Other encodings please check [msdoc](https://docs.microsoft.com/en-us/windows/win32/intl/code-page-identifiers) to get encoding value
(i.e.	`iso-8859-1`)
### mysql_GB2312.bas
Set `SYSENCODING`=28591 and `mysql_set_character_set` mysqlHandler, "iso-8859-1"

### postgres_UTF8.bas
Set `SYSENCODING`=28591



## Dependencies

[VBA_List](https://github.com/Vitosh/VBA_List)
