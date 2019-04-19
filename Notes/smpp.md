# Note for SMPP

## 1. tip for Long SMS (no test yet, just for remark only)
```csharp
// 参考《3GPP TS 23.040 V6.8.1 (2006-10).pdf》

// Set UDHI Flag Data.SM_UDH_GSM=0x40
request.setEsmClass((byte) Data.SM_UDH_GSM);

// 设置UDH内容
ByteBuffer ed = new ByteBuffer();
ed.appendByte((byte) 5); // UDH Length
ed.appendByte((byte) 0x00); // IE Identifier
ed.appendByte((byte) 3); // IE Data Length
ed.appendByte((byte) refNum); // Reference Number
ed.appendByte((byte) totalSegments); // Number of pieces
ed.appendByte((byte) i); // Sequence number
StringBuilder builder = new StringBuilder();

// 将短信内容编码
try {
	ed.appendString(message, encode);
	request.setShortMessageData(ed);
	return request;
} catch (Exception e) {
	
}

```

ref [source](http://www.voidcn.com/article/p-qdnnuwvj-bck.html)

