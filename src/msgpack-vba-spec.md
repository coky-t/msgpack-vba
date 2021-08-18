# MessagePack for VBA specification

## Reference

### MessagePack specification

Last modified at 2017-08-09 22:42:07 -0700
Sadayuki Furuhashi c 2013-04-21 21:52:33 -0700
https://github.com/msgpack/msgpack/blob/9aa092d6ca81f12005bd7dcbeb6488ad319e5133/spec.md

### VBA data types

https://docs.microsoft.com/en-us/office/vba/language/reference/user-interface-help/data-type-summary

## Serialization: type to format conversion for VBA

MessagePack for VBA serializers convert VBA types into MessagePack formats as following:

VBA types                  | source types | output format
-------------------------- | ------------ | ------------------------------------------------------------------------------------
Empty                      | Nil          | nil
Null                       | Nil          | nil
Integer                    | Integer      | int format family (positive fixint, negative fixint, int 8/16 or uint 8)
Long                       | Integer      | int format family (positive fixint, negative fixint, int 8/16/32 or uint 8/16)
Single                     | Float        | float format family (float 32)
Double                     | Float        | float format family (float 64)
Currency                   | Extension    | ext format family (fixext 1/2/4/8, type: 6)
Date                       | Extension    | ext format family (fixext 4/8 or ext 8, type: -1(timestamp)) (or fixext 8, type: 7)
String                     | String       | str format family (fixstr or str 8/16/32)
Object                     | -            | -
Error                      | -            | -
Boolean                    | Boolean      | bool format family (false or true)
Variant                    | -            | -
DataObject                 | -            | -
Decimal                    | Extension    | ext format family (fixext 1/2/4/8 or ext 8, type: 14)
Byte                       | Integer      | int format family (positive fixint or uint 8)
LongLong                   | Integer      | int format family (positive fixint, negative fixint, int 8/16/32/64 or uint 8/16/32)
User-defined Type          | -            | -
Array (Single Dimension)   | Array        | array format family (fixarray or array 16/32)
Array (Multiple Dimension) | -            | -
Byte() (Single Dimension)  | Binary       | bin format family (bin 8/16/32)

VBA types                  | source types | output format
-------------------------- | ------------ | ------------------------------------------------------------------------------------
Nothing                    | Nil          | nil
User-defined Class         | -            | -
Collection                 | Array        | array format family (fixarray or array 16/32)
Dictionary                 | Map          | map format family (fixmap or map 16/32)

VBA types                  | source types | output format
-------------------------- | ------------ | ------------------------------------------------------------------------------------
(none)                     | Integer      | uint 64

## Deserialization: format to type conversion

MessagePack for VBA deserializers convert MessagePack formats into VBA types as following:

source formats                                 | output type | VBA types
---------------------------------------------- | ----------- | ---------------------
positive fixint, and uint 8                    | Integer     | Byte
negative fixint, and int 8/16                  | Integer     | Integer
int 32, and uint 16                            | Integer     | Long
int 64, and uint 32                            | Integer     | LongLong (or Decimal)
uint 64                                        | Integer     | Decimal
nil                                            | Nil         | Null
false and true                                 | Boolean     | Boolean
float 32                                       | Float       | Single
float 64                                       | Float       | Double
fixstr and str 8/16/32                         | String      | String
bin 8/16/32                                    | Binary      | Byte()
fixarray and array 16/32                       | Array       | Collection (or Array)
fixmap map 16/32                               | Map         | Dictionary
fixext     4/8 and ext 8, type: -1 (timestamp) | Extension   | Date
fixext 1/2/4/8          , type:  6             | Extension   | Currency
fixext       8          , type:  7             | Extension   | Date
fixext 1/2/4/8 and ext 8, type: 14             | Extension   | Decimal

## Limitation

target                      | maximum                                  | Message Pack spec | Message Pack for VBA
--------------------------- | ---------------------------------------- | ----------------- | --------------------
Whole Serialized Byte Array | maximum length                           | unlimited         | `(2^31)-1`
Binary object               | maximum length                           | `(2^32)-1`        | less than `(2^31)-1`
String object               | maximum byte size                        | `(2^32)-1`        | less than `(2^31)-1`
Array object                | maximum number of elements               | `(2^32)-1`        | less than `(2^31)-1`
Map object                  | maximum number of key-value associations | `(2^32)-1`        | less than `(2^31)-1`

target                      | minimum/maximum/etc. | Message Pack spec                         | Message Pack for VBA
--------------------------- | -------------------- | ----------------------------------------- | --------------------
Timestamp 64                | nanosec              | 0 to 999999999                            | ignored
Timestamp 96                | nanosec              | 0 to 999999999                            | ignored
Timestamp 96                | minimum date/time    | -292277022657-01-27 08:29:52 UTC          | 100-01-01 00:00:00
Timestamp 96                | maximum date/time    | 292277026596-12-04 15:30:07.999999999 UTC | 9999-12-31 23:59:59
