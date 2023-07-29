SELECT o.[name],
    SCHEMA_NAME(o.[schema_id]) AS [schema],
    o.modify_date AS last_modified, c.modify_date AS test,
	o.type_desc,
	CASE o.[type]
		WHEN 'V' THEN 'views'
		WHEN 'U' THEN 'tables'
		WHEN 'IT' THEN 'tables'		-- Internal tables
		WHEN 'TR' THEN 'tables'
		WHEN 'P' THEN 'procedures'
		WHEN 'FN' THEN 'functions'	-- Scalar function
		WHEN 'IF' THEN 'functions'	-- Inline table valued function
		WHEN 'TF' THEN 'functions'	-- Table valued function
		WHEN 'TT' THEN 'types'		-- Type table
		WHEN 'SO' THEN 'sequences'	-- Sequence object
		WHEN 'SN' THEN 'synonymns'	-- Synonyms
		ELSE 'unknown'
	END as folder,
	o.[type] AS object_type
    -- ,*
FROM sys.objects o
-- Join child objects to find greatest last_modified date
LEFT JOIN 
sys.objects c ON c.object_id = o.parent_object_id
WHERE 1 = 1
--AND o.type = 'TT'
AND o.parent_object_id = 0
AND o.[type] NOT IN (
	 'S'  -- System Tables
	,'SQ' -- Service queues
	,'TR'  -- Triggers saved from tables
	,'IT'  -- Internal tables
	,'TT'  -- Type tables
	,'SO'  -- Sequence objects
	)
