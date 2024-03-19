-- On MySQL Server, return a listing of all parent level objects

-- Tables
SELECT 
	`TABLE_SCHEMA` AS `schema`,
    `TABLE_NAME` AS `name`,
    coalesce(`UPDATE_TIME`, `CREATE_TIME`) AS `last_modified`,
    'tables' as `folder`,
    null as `definition`
FROM information_schema.tables
where TABLE_TYPE = 'BASE TABLE'
and TABLE_SCHEMA NOT IN ('information_schema', 'mysql', 'performance_schema')

-- Views
UNION SELECT
	`TABLE_SCHEMA` AS `schema`,
    `TABLE_NAME` AS `name`,
    null AS `last_modified`,
    'views' as `folder`,
    `VIEW_DEFINITION` as `definition`
FROM information_schema.views
where 1=1
and TABLE_SCHEMA NOT IN ('mysql')

-- Functions
UNION SELECT
	`ROUTINE_SCHEMA` AS `schema`,
    `ROUTINE_NAME` AS `name`,
    coalesce(`LAST_ALTERED`, `CREATED`) AS `last_modified`,
    'functions' as `folder`,
    null as `definition`
FROM information_schema.routines
where ROUTINE_TYPE = 'FUNCTION'
and ROUTINE_SCHEMA NOT IN ('mysql')

-- Stored Procedures
UNION SELECT
	`ROUTINE_SCHEMA` AS `schema`,
    `ROUTINE_NAME` AS `name`,
    coalesce(`LAST_ALTERED`, `CREATED`) AS `last_modified`,
    'procedures' as `folder`,
    null as `definition`
FROM information_schema.routines
where ROUTINE_TYPE = 'PROCEDURE'
and ROUTINE_SCHEMA NOT IN ('mysql')

-- Triggers
UNION SELECT
	`TRIGGER_SCHEMA` AS `schema`,
    `TRIGGER_NAME` AS `name`,
    `CREATED` AS `last_modified`,
    'triggers' as `folder`,
    `ACTION_STATEMENT` as `definition`
FROM information_schema.triggers
where 1=1
