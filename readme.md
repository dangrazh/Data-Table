# Data-Table
A VB class providing advanced array based data handling. Built with several features optimized for MS Excel but working as well with VB6 or any other application supporting VBA.

ATTENTION: While the basic functionality is fully tested and should work without issues, some of the functionaliy is still under construction and not fully tested!
 
## Classes
You need to include the following three files into your project to use the Data-Table class.

|Class|Description|
| --- | --- |
|cDataTable.cls|This is the main class exposing all the functionality.|
|cIndex.cls|This class provides a database index like encapsulation of either an idex with unique keys (based on Hash Table class)|
|cHashTable.cls|This class implements a hashtable, a structure that offers many of the features of a collection or dictionary.|

## Methods and Attributes
The Data-Table class exposes the following Attributes and Methods.

##### Attributes
|Direction  | Attribute            | Description|
|-----------|----------------------|--------------------------------------|
|Get        | Version              | Returns a String containing the version info|
|Get        | About                | Returns a String containing an info about the class|
|Get        | ClassName            | Returns a String containing the class name|
|Get / Let  | Name                 | Sets/Returns a String containing the name of the data table|
|Get        | Headers              | Returns a Collection with the column names|
|Get        | HeaderList           | Returns a string with a comma delimited list of the column names|
|Get        | HasHeaders           | Returns a Boolean indicating if the data tables has defined column names|
|Get        | RecordCount          | Returns a Long containing the number of records in the data table|
|Get        | ColumnCount          | Returns a Long containing the number of columns in the data table|
|Get        | NumItems             | (Legacy - do not use anymore) Returns a Long containing the number of records in the data table|
|Get        | NumCols              | (Legacy - do not use anymore) Returns a Long containing the number of columns in the data table|
|Get / Let  | GarbageCollection    | Sets/Returns a Boolean indicating if garbage collection is active|
|Get / Let  | ObjectStorageEnabled | Sets/Returns a Boolean indicating if object can be stored in the data table|
|Get        | TableSummary         | Returns a String containing a summary of the data table including as name, no of records, no of columns|
|Get        | SADescrPtr           | Returns a LongPtr to the Safe Array Description of the data table|
|Get        | SAStructPtr          | Returns a LongPtr to the Safe Array Structure of the data table|
|Get        | RsEOF                | Returns a Boolean indicating if the end of the table has been reached|
|Get        | RsBOF                | Returns a Boolean indicating if the beginning of the table has been reached|
|Get / Let  | RsBookmark           | Sets/Returns a Long containing the currently active record in the data table|
|Get / Let  | ItemName             | Requires the Column Number as parameter, returns a String containing the Column Name or sets the column name to the value provided|
|Get        | ItemIndex            | Requires the Column Name as parameter, returns the Column Number|
|Get / Let  | Item                 | Requires the Column Number or Column Name and (optional) the Row Number as parameter, Returns a Variant contianing the data stored in the item|
|Get        | ItemRaw              | Requires the Row Number and Column Number as parameter, Returns a Variant contianing the data stored in the item. Attention: This is accessing the item directly, with no checks and safety nets! This is not thread save either!|
|Get / Let  | Record               | Requires the Row Number as parameter, Returns a Variant Array containing the record specified|

##### Methods
|Method            | Parameters | Description|
|------------------|------------|------------|
|DefineTable       | NoOfColumns As Long, Optional ColumnHeaders As String = "n/a", Optional NoOfRows As Long = 1| Define the structure of the table. This is the frist thing to do after creating the object, unless you load a range or a delimited file.|
|CreateEmptyCopy       | --| Creates an empty copy of the data table including structure but without any data or indices being copied|
|ColumnsAdd        | ParamArray NewColumns() As Variant | Add 1..n new columns to an existing Data-Table|
|IndexAdd          | ColumnName As String, idxType As IndexType|Add an index to the specified column|
|IndexRemove       | ColumnName As String|Add the index from the specified column|
|TruncateTable     | AskForConfirmation As Boolean| Truncate the table and delete all content but not the structure|
|RecordAddOld      | ByVal aRecord As Variant| Adding a record to the data table with an array containing the data as input
|RecordAdd         | ParamArray Record() As Variant| Adding a record to the data table passing each column as argument to the method|
|RecordRemove      | ByVal Position As Long| Delete a single record from the data table|
|LoadRange         | InputTable As Range, TableHasHeaders As Boolean| Load an MS Excel worksheet range into an empty data table. Note: if you load a range, you do not have to define the data table before.|
|LoadDelimTextFile | Filename As String, Optional FieldDelimiter As String = ",", Optional RecordDelimiter As String = vbNewLine, Optional TableHasHeaders As Boolean = True, Optional TextQualifier As String = """" | Load a delimited text file into an empty data table. Note: if you load a delimited text file, you do not have to define the data table before.|
|AppendToTable     | tSource As cDataTable, Optional tyAppend As AppendType|Append another Data-Table to this Data-Table|
|RsMoveFirst       | --|Move the RSBookmark to the first record in the data table|
|RsMoveLast        | --|Move the RSBookmark to the last record in the data table|
|RsMoveNext        | --|Move the RSBookmark to the next record in the data table| 
|RsMovePrevious    | --|Move the RSBookmark to the previous record in the data table|
|RsFindFirst       | Index As Variant, match As MatchType, Criteria As Variant|Move the RSBookmark to frist record matching the criteria - currently only single columns can be used in the search criteria.|
|RsFindNext        | Index As Variant, match As MatchType, Criteria As Variant|Move the RSBookmark to next record matching the criteria - currently only single columns can be used in the search criteria.|
|SelectDistinctData | ParamArray Fields() As Variant|Select a distinct set of data based on the fields passed as input. The result is returned as new data table|
|DumpToRange       | TargetWorksheet As Worksheet, TargetCell As Range, Optional IncludeHeader As Boolean = True, Optional CompressOnRowOverflow As Boolean = False|Write the full content of the data table to a MS Excel Worksheet range|
|DumpToFile        | TargetFile As String, Delimiter As String, Optional IncludeHeader As Boolean = True, Optional OutputMode As OutputType = OverwriteIfExists|Write the full content of the data table to a delimited text file|
|Sort              | ParamArray SortOrder As Variant|Sort the data table with a stable (merge sort) sort mechanism - multiple sort columns as well as sort dirctions can be specified.|
|SortUnstable      | Index As Variant, Optional Direction As String = "asc"|Sort the data table with a fast but unstable sort mechanism - only one sort column can be specified.

##### Methods - Pending Implementation
|Method            | Parameters | Description|
|------------------|------------|------------|
|AnalyzeDataTypes  | --|§§§ Not yet implemented §§§|
|IndicesRefresh    | --|§§§ Not yet implemented §§§|
|RunStats          | --|§§§ Not yet implemented §§§|
|SelectData        | --|§§§ Not yet implemented §§§|