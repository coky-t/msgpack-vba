# MessagePack for VBA implementation

## Specification

[MessagePack for VBA specification](src/msgpack-vba-spec.md)

## Usage

### Prepare

Import [MsgPack.bas](src/MsgPack.bas)

### Serialization

```
    Dim MPBytes() As Byte
    MPBytes = MsgPack.GetMPBytes(Value)
    ' do anything
```

### Deserialization

```
    Dim Value
    If MsgPack.IsMPObject(MPBytes) Then
        Set Value = MsgPack.GetValue(MPBytes)
        If TypeName(Value) = "Collection" Then
            ' do anything
        ElseIf TypeName(Value) = "Dictionary" Then
            ' do anything
        End If
    Else
        Value = MsgPack.GetValue(MPBytes)
        ' do anything
    End If
```

## License

[MIT License](LICENSE)
