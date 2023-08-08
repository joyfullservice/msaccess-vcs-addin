-- On Microsoft SQL Server, return a listing of all parent level objects
SELECT o.[name],
    SCHEMA_NAME(o.[schema_id]) AS [schema],
	CASE
		-- Return the most recent modfied date of the object or any dependent object
		WHEN isnull(c.max_modified, 0) > o.modify_date THEN c.max_modified
		ELSE o.modify_date
	END AS last_modified,
	o.type_desc,
	CASE o.[type]
		WHEN 'V'  THEN 'View'
		WHEN 'U'  THEN 'Table'
		WHEN 'IT' THEN 'Table'			-- Internal tables
		WHEN 'TR' THEN 'Trigger'
		WHEN 'P'  THEN 'StoredProcedure'
		WHEN 'FN' THEN 'Function'		-- Scalar function
		WHEN 'IF' THEN 'Function'		-- Inline table valued function
		WHEN 'TF' THEN 'Function'		-- Table valued function
		WHEN 'TT' THEN 'UserDefinedTableType'	-- Type table
		WHEN 'SO' THEN 'Sequence'		-- Sequence object
		WHEN 'SN' THEN 'Synonymn'		-- Synonyms
		ELSE 'unknown'
	END as folder,
	o.[type] AS object_type
    -- ,*
FROM sys.objects o
LEFT JOIN 
	-- Get most recent modified date of any child object
	(select 
		parent_object_id,
		max(modify_date) AS max_modified
		from sys.objects
		WHERE parent_object_id > 0
		GROUP BY parent_object_id
	)AS c 
	ON c.parent_object_id = o.object_id
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
