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
|Get        | HasHeaders           | Returns a Boolean indicating if the data tables has defined column names|
|Get        | NumItems             | Returns a Long containing the number of records in the data table|
|Get        | NumCols              | Returns a Long containing the number of columns in the data table|
|Get / Let  | GarbageCollection    | Sets/Returns a Boolean indicating if garbage collection is active|
|Get / Let  | ObjectStorageEnabled | Sets/Returns a Boolean indicating if object can be stored in the data table|
|Get        | TableSummary         | Returns a String containing a summary of the data table including as name, no of records, no of columns|
|Get        | SADescrPtr           | Returns a LongPtr to the Safe Array Description of the data table|
|Get        | SAStructPtr          | Returns a LongPtr to the Safe Array Structure of the data table|
|Get        | RsEOF                | Returns a Boolean indicating if the end of the table has been reached|
|Get        | RsBOF                | Returns a Boolean indicating if the beginning of the table has been reached|
|Get / Let  | RsBookmark           | Sets/Returns a Long containing the currently active record in the data table|
|Get        | ItemName             | Requires the Column Number as parameter, returns a String containing the Column Name|
|Get / Let  | Item                 | Requires the Column Number or Column Name and (optional) the Row Number as parameter, Returns a Variant contianing the data stored in the item|
|Get / Let  | Record               | Requires the Row Number as parameter, Returns a Variant Array containing the record specified|

##### Methods
|Method|Description|
| --- | --- |
|To be completed||
