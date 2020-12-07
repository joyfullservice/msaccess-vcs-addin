SELECT MSysNavPaneGroups.Name AS GroupName, MSysNavPaneGroups.Flags AS GroupFlags, MSysNavPaneGroups.Position AS GroupPosition, MSysObjects.Type AS ObjectType, MSysObjects.Name AS ObjectName, MSysNavPaneGroupToObjects.Flags AS ObjectFlags, MSysNavPaneGroupToObjects.Icon AS ObjectIcon, MSysNavPaneGroupToObjects.Position AS ObjectPosition
FROM MSysNavPaneGroups LEFT JOIN (MSysNavPaneGroupToObjects LEFT JOIN MSysObjects ON MSysNavPaneGroupToObjects.ObjectID = MSysObjects.Id) ON MSysNavPaneGroups.Id = MSysNavPaneGroupToObjects.GroupID
WHERE (((MSysNavPaneGroups.Name) Is Not Null) AND ((MSysNavPaneGroups.GroupCategoryID)=3))
ORDER BY MSysNavPaneGroups.Name, MSysObjects.Type, MSysObjects.Name;
