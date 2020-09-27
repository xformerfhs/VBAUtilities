# VBAUtilities

This is a collection of VBA utilities that I wrote in the last 20 years. I did this because I really love simplicity and elegance and far too many code examples on the internet have neither of these. 

So, I hope you will find these utilities useful. Use them as you wish. Just include a copyright notice that references me. Enjoy.

The utilities are categorized in the following way:

## Access

### AccessLockFileReader

A class to read an access lock file in order to find out which computer is holding a lock on the access database. This can either be the database that this class is part of or another database file. It supports mdb and accdb files.

### ADOFieldWrapper

Wrapper around ADO fields for easier property access.

### ADOTableWrapper

Wrapper around ADO tables for easier property access.

### DAOPropertyManager

Easily set and get DAO properties.

### DBBackupManager

Make a copy of a database and keep a specified number of copied files.

### DBCompressor

Manage compression of an Access database.

### DBTableLinkHelper

Get and change the path to a linked table.

### StatusLine

Show and clear an Access status line. The status line is automatically cleared when the class is destroyed.

## Crypto

### HashCng

A universal hashing class. It calculates [SHA-1](https://en.wikipedia.org/wiki/SHA-1) and [SHA-2](https://en.wikipedia.org/wiki/SHA-2) hashes (with 256, 384 and 512 bits length) and also [HMAC](https://en.wikipedia.org/wiki/HMAC) values with these hashes. It uses  the Windows CNG (Crypto Next Generation) API, so all calculations are done by Windows.

### SecureRandomNumberCng

Secure CNG random number generator which is a wrapper around the Windows CNG (Crypto Next Generation) RNG API and uses the [BCryptGenRandom](https://docs.microsoft.com/en-us/windows/win32/api/bcrypt/nf-bcrypt-bcryptgenrandom) function.

## ErrorHandling

### MessageManager

When calling Windows API functions one gets a return code. In order to find out what it means and present that return code as a text to the user this class translates the return code to a string. It just asks Windows what the meaning of the return code ist. BTW, it also has a method to handle message templates with positional parameter substitution.

## ExcelUtilities

### WorkbookCustomPropertyHandler

This class handles custom properties for an Excel workbook.

### WorksheetCustomPropertyHandler

This class handles custom properties for an Excel worksheet.

## FileHandling

### DriveHelper

Get information about a drive, i.e. what type it is and whether it is a network drive.

### FileCompressionManager

Managing file compression is horribly complicated under Windows as this is not an attribute but something that is achieved by issuing I/O control command.
This class puts a simple wrapper around the complexity of handling compressed files. One can create a compressed file and read, set and clear the compression state of a file.

### RandomFileName

Helper class to generate a unique random file name.

## Internet

### FTPClient

A simple FTP client for VBA programs. It has the following public methods

* Connect
* Disconnect
* GetFile
* PutFile
* CreateDirectory
* RemoveDirectory
* DeleteFile
* GetCurrentDirectory
* SetCurrentDirectory
* DirFiles

## Math

### BearingHelper

A little helper class to calculate interpolations between different navigational [bearings](https://en.wikipedia.org/wiki/Bearing_(navigation)).

### SphereDistanceCalculator

Helper class to calculate distances and bearings on a sphere when the positions (latitude, longitude) are known.

### Trigonometrics

Adds missing trigonometric functions to VBA:

* ArcCos
* ArcSin
* ArcTan2
* RadiantToDegree
* DegreeToRadiant

## NumberConversion

### Base64Converter

Pure VB implementation of a Base64 converter. Converts byte arrays to and from Base64 representation.

### Base64ConverterCryptAPI

Implementation of a Base64 converter as a wrapper around crypt32.dll API calls. Converts byte arrays to and from Base64 representation.

### HexConverter

Converts byte arrays to and from hexadecimal string representation.

### RomanNumberConverter

Converts integer to and from roman number representations.

## OSUtilities

### SetPriorityClass

Class to set the currently running processes priority class to give it a higher or lower scheduling priority.

### SpecialFolder

Get Windows special folder names.

### SystemInformation

Get some system informations.

## Sorting

### InsertionSort

An implementation of the insertion sort algorithm.

### InsertionSortWithIndex

Implementation of the insertion sort algorithm where not the data array is sorted but an index into the data array. This is especially useful when moving data is an expensive operation like e.g. for strings. 

### PureQuickSort

A pure quicksort implementation.

### PureQuickSortWithIndex

Pure quicksort implementation where not the data array is sorted but an index into the data array. This is especially useful when moving data is an expensive operation like e.g. for strings. 

### QuickSort

An optimized quicksort implementation. Here quicksort is combined with insertion sort to make the implementation faster.

### QuickSortWithIndex

Optimized quicksort implementation where not the data array is sorted but an index into the data array. This is especially useful when moving data is an expensive operation like e.g. for strings. 

### Stack

An implementation of a stack. Used by the quicksort type sorter classes.

## StringHandling

### StringBuilder

An implementation of one of the most important classes that is missing in VBA: A string builder. It allows method chaining like in e.g. `sb.SetTo("Content").Append(aVar).Append(anotherVar)`.

### UTF8Converter

Converts VBA strings from and to [UTF-8](https://en.wikipedia.org/wiki/UTF-8) encoding. Note that the UTF-8 values are byte arrays, not strings. Storing UTF-8 encodings in VBA strings is seriously wrong.

## Time utilitites

### TimeConverter

Converts VBA timestamps from and to Unix timestamps or local time from and to UTC time.

## Timing

### HighPrecisionTimer

A high precision timer that uses the Windows Performance Counter which has a resolution better than 0.000001 seconds (1µs). 

### Waiter

Suspend program execution for a specified amount of time or for a random amount of time while keeping Access responsive.
