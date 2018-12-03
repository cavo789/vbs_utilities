# Database - Test connection

> Try to establish a connection to a SQL database

This script will allow you to quickly check if access to your SQL Server database is possible.

The objective is to establish a connection and check if it works before starting, e. g., to investigate your program code or the permissions required for the user to use your tables, views, stored procedures,...

This script will only do this, i.e. try to connect to the database, which will eliminate the possibility of a login problem.

## Table of Contents

- [Usage](#usage)
- [License](#license)

## Usage

Run the script with four parameters like:

```
cscript test_connection.vbs serverName dbName login password
```

The output will looks like:

```
Try to connect [dbName] on [serverName]...

Connection string: Provider=SQLOLEDB;Data Source=[serverName];Trusted_Connection=False;Initial Catalog=[dbName];User ID=[login];Password=[password];

SELECT db_name()
Active database name= [dbName]

Test successful, database connection has been successfully established

Time taken: 0,01
```

## License

[MIT](LICENSE)
