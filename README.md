# Excel Read SQL
Read SQL Data with only ONE Function!

Check your Excel fits:


`2019+` and **X64** platform


NOTE: If EXCEL version is below `2019` you may use `Ctrl+Shift+Enter` to apply array function.

## Installation

- copy folder to local (e.g `src\Mysql` to `D:\lib\vba_sql`)
- In EXCEL: Developer -> Excel Add-ins -> Browse (D:\lib\vba_sql\Mysql) -> `mysql_UTF8.xlam` and check on `mysql_UTF8` -> Ok
- Enjoy

Or if you have own libary:

- import `src\xxx\xxx.bas` and `src\vba_list\List.cls` to your own project
- copy dll files (e.g `D:\lib\sqllib`)
- **put real lib dir path into `xxxx_Initialize()`** --> `xxxx_Initialize("D:\lib\sqllib")`

## Usage

### MySQL

```
=MSQLR("user:password@host:port/database","SELECT date,data from table",TRUE)
# Do not Start with `mysql://`
```

### PostgreSQL
```
=PSQLR("postgresql://user:password@host:port/database","SELECT date,data from table",TRUE)
# surport:
# postgres://localhost
# postgresql://localhost:5433
# postgresql://localhost/mydb
# postgresql://user@localhost
# postgresql://user:secret@localhost
# postgresql://other@localhost/otherdb?connect_timeout=10&application_name=myapp
# postgres://host1:123,host2:456/somedb?target_session_attrs=any&application_name=myapp
```

### Sqlite
Codes are basically from [SQLiteForExcel](https://github.com/govert/SQLiteForExcel), but keep `Read` related functions.


```
=SQLR("C:\data\data.db","SELECT date,data from table",TRUE)
```

## Encoding Issue
Other encodings please check [msdoc](https://docs.microsoft.com/en-us/windows/win32/intl/code-page-identifiers) to get encoding value


(e.g.	`iso-8859-1`)

### mysql_GB2312.bas

Set `SYSENCODING`=28591 and `mysql_set_character_set` mysqlHandler, "iso-8859-1"

### postgres_UTF8.bas

Set `SYSENCODING`=28591


## Dependencies

[VBA_List](https://github.com/Vitosh/VBA_List) in \src\vba_list
