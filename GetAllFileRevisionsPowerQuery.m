// This PowerQuery retrieves version history information for files in a sharepoint library
//
// Usage: 
// This PowerQuery can be used in PowerBI or Excel and will prompt you for your sharepoint login details.
// Note that your will have to have access rights to the library you're trying to access.
// To use this function, modify the SiteURL and LibraryPath variables to match your sharepoint library location
//
// The PowerQuery will produce a table with the following columns:
//
// RelativeUrl	FileName	Title	CreatedOn	LastModifiedOn	FileSize	CurrentVersion	PreviousVersion	VersionCreatedOn	RelativeVersionURL	VersionFileSize
//


let
    // Replace with your SharePoint site and library details. Use URLEncoded names.
    SiteURL = "https://[your_company_name].sharepoint.com/sites/[sitename]",
    LibraryPath = "/sites/[sitename]/[library_name]/",
    
    // function GetFiles retrieves all files in a sharepoint library, including those in folders.
	GetFiles = (libPath as text, siteURL as text) as table => 
        let
            src = Xml.Tables(Web.Contents(Text.Combine({siteURL,"/_api/web/GetFolderByServerRelativeUrl('",libPath ,"')/Files"}))),
                entry = src{0}[entry],
                #"Removed Other Columns2" = Table.SelectColumns(entry,{"content"}),
                #"Expanded content" = Table.ExpandTableColumn(#"Removed Other Columns2", "content", {"http://schemas.microsoft.com/ado/2007/08/dataservices/metadata"}, {"content"}),
                #"Expanded content1" = Table.ExpandTableColumn(#"Expanded content", "content", {"properties"}, {"properties"}),
                #"Expanded properties" = Table.ExpandTableColumn(#"Expanded content1", "properties", {"http://schemas.microsoft.com/ado/2007/08/dataservices"}, {"properties"})
        in
            #"Expanded properties",
        
	Source = GetFiles(LibraryPath, SiteURL),
    resultsTable = Table.ExpandTableColumn(Source, "properties", {"CheckInComment", "CheckOutType", "ContentTag", "CustomizedPageStatus", "ETag", "Exists", "ExistsAllowThrowForPolicyFailures", "ExistsWithException", "IrmEnabled", "Length", "Level", "LinkingUri", "LinkingUrl", "MajorVersion", "MinorVersion", "Name", "ServerRelativeUrl", "TimeCreated", "TimeLastModified", "Title", "UIVersion", "UIVersionLabel", "UniqueId"}, {"properties.CheckInComment", "properties.CheckOutType", "properties.ContentTag", "properties.CustomizedPageStatus", "properties.ETag", "properties.Exists", "properties.ExistsAllowThrowForPolicyFailures", "properties.ExistsWithException", "properties.IrmEnabled", "properties.Length", "properties.Level", "properties.LinkingUri", "properties.LinkingUrl", "properties.MajorVersion", "properties.MinorVersion", "properties.Name", "properties.ServerRelativeUrl", "properties.TimeCreated", "properties.TimeLastModified", "properties.Title", "properties.UIVersion", "properties.UIVersionLabel", "properties.UniqueId"}),

    // Function to get version history for a file
    GetFileVersions = (FilePath as text) as table =>
        let
            // create an empty table to use when no results are found
            EmptyTable = Table.FromRecords({}, {"http://www.w3.org/XML/1998/namespace"}),

            // URLencode the file location
            EncodedFilePath = Uri.EscapeDataString(FilePath),

            // build the complete URL to get the file versions
            VersionsURL = SiteURL & "/_api/web/GetFileByServerRelativeUrl('" & EncodedFilePath & "')/versions",

            // try to retrieve results. if it doesn't work, use an empty table instead
            FileVersions = try Xml.Tables(Web.Contents(VersionsURL)) otherwise EmptyTable,

            // if results were found clean up the data
            Result = if FileVersions <> null and Table.HasColumns(FileVersions, "entry") then 
            let
            #"Expanded entry" = Table.ExpandTableColumn(FileVersions, "entry", {"id", "category", "link", "title", "updated", "author", "content"}, {"entry.id", "entry.category", "entry.link", "entry.title", "entry.updated", "entry.author", "entry.content"}),
            #"Removed Columns1" = Table.RemoveColumns(#"Expanded entry",{"entry.link", "entry.title", "http://www.w3.org/XML/1998/namespace"}),
            #"Expanded entry.content" = Table.ExpandTableColumn(#"Removed Columns1", "entry.content", {"http://schemas.microsoft.com/ado/2007/08/dataservices/metadata"}, {"entry.content.http://schemas.microsoft.com/ado/2007/08/dataservices/metadata"}),
            #"Expanded entry.content.http://schemas.microsoft.com/ado/2007/08/dataservices/metadata" = Table.ExpandTableColumn(#"Expanded entry.content", "entry.content.http://schemas.microsoft.com/ado/2007/08/dataservices/metadata", {"properties"}, {"entry.content.http://schemas.microsoft.com/ado/2007/08/dataservices/metadata.pro"}),
            #"Expanded entry.content.http://schemas.microsoft.com/ado/2007/08/dataservices/metadata.pro" = Table.ExpandTableColumn(#"Expanded entry.content.http://schemas.microsoft.com/ado/2007/08/dataservices/metadata", "entry.content.http://schemas.microsoft.com/ado/2007/08/dataservices/metadata.pro", {"http://schemas.microsoft.com/ado/2007/08/dataservices"}, {"entry.content.http://schemas.microsoft.com/ado/2007/08/dataservices/metadata.p.1"}),
            #"Expanded entry.content.http://schemas.microsoft.com/ado/2007/08/dataservices/metadata.p.1" = Table.ExpandTableColumn(#"Expanded entry.content.http://schemas.microsoft.com/ado/2007/08/dataservices/metadata.pro", "entry.content.http://schemas.microsoft.com/ado/2007/08/dataservices/metadata.p.1", {"CheckInComment", "Created", "ID", "IsCurrentVersion", "Length", "Size", "Url", "VersionLabel"}, {"entry.content.http://schemas.microsoft.com/ado/2007/08/dataservices/metadata.p.2", "entry.content.http://schemas.microsoft.com/ado/2007/08/dataservices/metadata.p.3", "entry.content.http://schemas.microsoft.com/ado/2007/08/dataservices/metadata.p.4", "entry.content.http://schemas.microsoft.com/ado/2007/08/dataservices/metadata.p.5", "entry.content.http://schemas.microsoft.com/ado/2007/08/dataservices/metadata.p.6", "entry.content.http://schemas.microsoft.com/ado/2007/08/dataservices/metadata.p.7", "entry.content.http://schemas.microsoft.com/ado/2007/08/dataservices/metadata.p.8", "entry.content.http://schemas.microsoft.com/ado/2007/08/dataservices/metadata.p.9"}),
            #"Expanded entry.content.http://schemas.microsoft.com/ado/2007/08/dataservices/metadata.p.3" = Table.ExpandTableColumn(#"Expanded entry.content.http://schemas.microsoft.com/ado/2007/08/dataservices/metadata.p.1", "entry.content.http://schemas.microsoft.com/ado/2007/08/dataservices/metadata.p.3", {"Element:Text"}, {"entry.content.http://schemas.microsoft.com/ado/2007/08/dataservices/metadata.p.1"}),
            #"Removed Columns2" = Table.RemoveColumns(#"Expanded entry.content.http://schemas.microsoft.com/ado/2007/08/dataservices/metadata.p.3",{"entry.content.http://schemas.microsoft.com/ado/2007/08/dataservices/metadata.p.4"}),
            #"Expanded entry.content.http://schemas.microsoft.com/ado/2007/08/dataservices/metadata.p.5" = Table.ExpandTableColumn(#"Removed Columns2", "entry.content.http://schemas.microsoft.com/ado/2007/08/dataservices/metadata.p.5", {"Element:Text"}, {"entry.content.http://schemas.microsoft.com/ado/2007/08/dataservices/metadata.p.3"}),
            #"Expanded entry.content.http://schemas.microsoft.com/ado/2007/08/dataservices/metadata.p.6" = Table.ExpandTableColumn(#"Expanded entry.content.http://schemas.microsoft.com/ado/2007/08/dataservices/metadata.p.5", "entry.content.http://schemas.microsoft.com/ado/2007/08/dataservices/metadata.p.6", {"Element:Text"}, {"entry.content.http://schemas.microsoft.com/ado/2007/08/dataservices/metadata.p.4"}),
            #"Expanded entry.content.http://schemas.microsoft.com/ado/2007/08/dataservices/metadata.p.7" = Table.ExpandTableColumn(#"Expanded entry.content.http://schemas.microsoft.com/ado/2007/08/dataservices/metadata.p.6", "entry.content.http://schemas.microsoft.com/ado/2007/08/dataservices/metadata.p.7", {"Element:Text"}, {"entry.content.http://schemas.microsoft.com/ado/2007/08/dataservices/metadata.p.5"}),
            #"Expanded entry.author" = Table.ExpandTableColumn(#"Expanded entry.content.http://schemas.microsoft.com/ado/2007/08/dataservices/metadata.p.7", "entry.author", {"name"}, {"entry.author.name"}),
            #"Removed Columns3" = Table.RemoveColumns(#"Expanded entry.author",{"entry.category"})
            
            in
            #"Removed Columns3"

            else FileVersions
        in Result,
    
    // Add version history to each file
    path = [properties.ServerRelativeUrl],
    enCodedPath = Uri.EscapeDataString(path),
    AddVersions = Table.AddColumn(resultsTable, "Versions", each GetFileVersions([properties.ServerRelativeUrl])),
    #"Expanded Versions" = Table.ExpandTableColumn(AddVersions, "Versions", {"id", "title", "updated", "author", "http://www.w3.org/XML/1998/namespace", "entry.id", "entry.updated", "entry.author.name", "entry.content.http://schemas.microsoft.com/ado/2007/08/dataservices/metadata.p.2", "entry.content.http://schemas.microsoft.com/ado/2007/08/dataservices/metadata.p.1", "entry.content.http://schemas.microsoft.com/ado/2007/08/dataservices/metadata.p.3", "entry.content.http://schemas.microsoft.com/ado/2007/08/dataservices/metadata.p.4", "entry.content.http://schemas.microsoft.com/ado/2007/08/dataservices/metadata.p.5", "entry.content.http://schemas.microsoft.com/ado/2007/08/dataservices/metadata.p.8", "entry.content.http://schemas.microsoft.com/ado/2007/08/dataservices/metadata.p.9"}, {"Versions.id", "Versions.title", "Versions.updated", "Versions.author", "Versions.http://www.w3.org/XML/1998/namespace", "Versions.entry.id", "Versions.entry.updated", "Versions.entry.author.name", "Versions.entry.content.http://schemas.microsoft.com/ado/2007/08/dataservices/met", "Versions.entry.content.http://schemas.microsoft.com/ado/2007/08/dataservices/m.1", "Versions.entry.content.http://schemas.microsoft.com/ado/2007/08/dataservices/m.2", "Versions.entry.content.http://schemas.microsoft.com/ado/2007/08/dataservices/m.3", "Versions.entry.content.http://schemas.microsoft.com/ado/2007/08/dataservices/m.4", "Versions.entry.content.http://schemas.microsoft.com/ado/2007/08/dataservices/m.5", "Versions.entry.content.http://schemas.microsoft.com/ado/2007/08/dataservices/m.6"}),
    #"Expanded properties.MajorVersion" = Table.ExpandTableColumn(#"Expanded Versions", "properties.MajorVersion", {"Element:Text"}, {"properties.MajorVersion.Element:Text"}),
    #"Expanded properties.MinorVersion" = Table.ExpandTableColumn(#"Expanded properties.MajorVersion", "properties.MinorVersion", {"Element:Text"}, {"properties.MinorVersion.Element:Text"}),
    #"Expanded properties.TimeCreated" = Table.ExpandTableColumn(#"Expanded properties.MinorVersion", "properties.TimeCreated", {"Element:Text"}, {"properties.TimeCreated.Element:Text"}),
    #"Expanded properties.TimeLastModified" = Table.ExpandTableColumn(#"Expanded properties.TimeCreated", "properties.TimeLastModified", {"Element:Text"}, {"properties.TimeLastModified.Element:Text"}),
    #"Expanded properties.UIVersion" = Table.ExpandTableColumn(#"Expanded properties.TimeLastModified", "properties.UIVersion", {"Element:Text"}, {"properties.UIVersion.Element:Text"}),
    #"Expanded properties.UniqueId" = Table.ExpandTableColumn(#"Expanded properties.UIVersion", "properties.UniqueId", {"Element:Text"}, {"properties.UniqueId.Element:Text"}),
    #"Expanded properties.Length" = Table.ExpandTableColumn(#"Expanded properties.UniqueId", "properties.Length", {"Element:Text"}, {"properties.Length.Element:Text"}),
    #"Expanded properties.Level" = Table.ExpandTableColumn(#"Expanded properties.Length", "properties.Level", {"Element:Text"}, {"properties.Level.Element:Text"}),

    // Clean up dataset by reordering columns and removing the ones we don't need
    #"Reordered Columns" = Table.ReorderColumns(#"Expanded properties.Level",{"properties.ServerRelativeUrl", "properties.Name", "properties.Length.Element:Text", "properties.LinkingUri", "properties.MajorVersion.Element:Text", "properties.MinorVersion.Element:Text", "properties.TimeCreated.Element:Text", "properties.TimeLastModified.Element:Text", "properties.Title", "properties.UIVersionLabel", "properties.UniqueId.Element:Text", "Versions.id", "Versions.updated", "Versions.author", "Versions.entry.id", "Versions.entry.updated", "Versions.entry.author.name", "Versions.entry.content.http://schemas.microsoft.com/ado/2007/08/dataservices/met", "Versions.entry.content.http://schemas.microsoft.com/ado/2007/08/dataservices/m.1", "Versions.entry.content.http://schemas.microsoft.com/ado/2007/08/dataservices/m.2", "Versions.entry.content.http://schemas.microsoft.com/ado/2007/08/dataservices/m.3", "Versions.entry.content.http://schemas.microsoft.com/ado/2007/08/dataservices/m.4", "Versions.entry.content.http://schemas.microsoft.com/ado/2007/08/dataservices/m.5", "Versions.entry.content.http://schemas.microsoft.com/ado/2007/08/dataservices/m.6"}),
    #"Reordered Columns1" = Table.ReorderColumns(#"Reordered Columns",{"properties.ServerRelativeUrl", "properties.Name", "properties.Title", "properties.Length.Element:Text", "properties.LinkingUri", "properties.TimeCreated.Element:Text", "properties.TimeLastModified.Element:Text", "properties.UIVersionLabel", "properties.UniqueId.Element:Text", "Versions.id", "Versions.updated", "Versions.author", "Versions.entry.id", "Versions.entry.updated", "Versions.entry.author.name", "Versions.entry.content.http://schemas.microsoft.com/ado/2007/08/dataservices/met", "Versions.entry.content.http://schemas.microsoft.com/ado/2007/08/dataservices/m.1", "Versions.entry.content.http://schemas.microsoft.com/ado/2007/08/dataservices/m.2", "Versions.entry.content.http://schemas.microsoft.com/ado/2007/08/dataservices/m.3", "Versions.entry.content.http://schemas.microsoft.com/ado/2007/08/dataservices/m.4", "Versions.entry.content.http://schemas.microsoft.com/ado/2007/08/dataservices/m.5", "Versions.entry.content.http://schemas.microsoft.com/ado/2007/08/dataservices/m.6"}),
    #"Renamed Columns" = Table.RenameColumns(#"Reordered Columns1",{{"properties.Length.Element:Text", "CurrentFileSize"}}),
    #"Renamed Columns1" = Table.RenameColumns(#"Renamed Columns",{{"properties.TimeCreated.Element:Text", "DocumentCreatedOn"}, {"properties.UIVersionLabel", "CurrentVersion"}}),
    #"Renamed Columns2" = Table.RenameColumns(#"Renamed Columns1",{{"properties.TimeLastModified.Element:Text", "LastModifiedOn"}, {"DocumentCreatedOn", "CreatedOn"}, {"CurrentFileSize", "FileSize"}, {"properties.Title", "Title"}, {"properties.Name", "FileName"}, {"properties.ServerRelativeUrl", "RelativeUrl"}}),
    #"Renamed Columns3" = Table.RenameColumns(#"Renamed Columns2",{{"Versions.entry.content.http://schemas.microsoft.com/ado/2007/08/dataservices/m.3", "VersionFileSize"}, {"Versions.entry.content.http://schemas.microsoft.com/ado/2007/08/dataservices/m.2", "IsCurrentVersion"}, {"Versions.entry.content.http://schemas.microsoft.com/ado/2007/08/dataservices/m.5", "RelativeVersionURL"}, {"Versions.entry.content.http://schemas.microsoft.com/ado/2007/08/dataservices/m.6", "Version"}, {"Versions.entry.content.http://schemas.microsoft.com/ado/2007/08/dataservices/m.1", "VersionCreatedOn"}}),
    #"Renamed Columns4" = Table.RenameColumns(#"Renamed Columns3",{{"Version", "PreviousVersion"}}),
    #"Reordered Columns2" = Table.ReorderColumns(#"Renamed Columns4",{"RelativeUrl", "FileName", "Title", "CreatedOn", "LastModifiedOn", "FileSize", "CurrentVersion", "PreviousVersion", "VersionCreatedOn", "RelativeVersionURL", "VersionFileSize"}),
    #"Changed Type" = Table.TransformColumnTypes(#"Reordered Columns2",{{"VersionCreatedOn", type datetimezone}, {"LastModifiedOn", type datetimezone}, {"CreatedOn", type datetimezone}}),
    #"Removed Columns" = Table.RemoveColumns(#"Changed Type",{"properties.CheckOutType", "properties.CustomizedPageStatus", "properties.Exists", "properties.ExistsAllowThrowForPolicyFailures", "properties.ExistsWithException", "properties.IrmEnabled"}),
    #"Removed Columns1" = Table.RemoveColumns(#"Removed Columns",{"properties.UIVersion.Element:Text", "Versions.title", "Versions.http://www.w3.org/XML/1998/namespace", "properties.Level.Element:Text", "properties.CheckInComment", "properties.ContentTag", "properties.ETag", "properties.LinkingUrl"}),
    #"Removed Columns2" = Table.RemoveColumns(#"Removed Columns1",{"properties.MajorVersion.Element:Text", "properties.MinorVersion.Element:Text"}),
    #"Removed Columns3" = Table.RemoveColumns(#"Removed Columns2",{"Versions.author", "Versions.entry.author.name", "Versions.entry.updated", "Versions.entry.content.http://schemas.microsoft.com/ado/2007/08/dataservices/met", "Versions.entry.content.http://schemas.microsoft.com/ado/2007/08/dataservices/m.4"}),
    #"Removed Columns4" = Table.RemoveColumns(#"Removed Columns3",{"properties.LinkingUri"}),
    #"Removed Columns5" = Table.RemoveColumns(#"Removed Columns4",{"properties.UniqueId.Element:Text", "Versions.id"}),
    #"Removed Columns6" = Table.RemoveColumns(#"Removed Columns5",{"Versions.entry.id", "Versions.updated"}),
    #"Removed Columns7" = Table.RemoveColumns(#"Removed Columns6",{"IsCurrentVersion"})
in
    #"Removed Columns7"