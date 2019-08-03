# **Note for SMPP**

## 1. **Tip for Long SMS (no test yet, just for remark only)**

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
try
{
    ed.appendString(message, encode);
    request.setShortMessageData(ed);
    return request;
} catch (Exception e)
{
}
```

ref [source](http://www.voidcn.com/article/p-qdnnuwvj-bck.html)

## 2. **SMPP simulator**

### a. Start Simulator

```Dos
open SMPPSim/startsmppsim.bat
```

After launch, <http://127.0.0.1:88> , check if the SMPPSim page display normal.

ref [source](https://blog.csdn.net/shulai123/article/details/68922174)

### b. Maybe need update the configuration file (conf/smppsim.props) as follow:

1. By default the SMPP simulator tries to open port 88 for it’s HTTP console, since need administrative rights to open ports lower than 1024, then change that to 8080 instead. 
   Line 60 “HTTP_PORT=88” to "HTTP_PORT=8080";
2. Another change is on line 83 "ESME_TO_ESME=TRUE" to "ESME_TO_ESME=FALSE".
3. line 69, change to OUTBIND_ENABLED=true
   

### 3. **Implement Sample of SMPP by .net**

ref [source](https://blog.csdn.net/gllzqfe/article/details/86149990)

### 4. **Chinese SMS**

Change Chinese context to byte[] messageByte = Encoding.UTF8.GetBytes("中文字符").

And check the length of messageByte, maybe need remove the first two bytes.

### 5. **Other Notes**

Normally, the no of message send via SMPP is 5 per second. And this service can upgrade to 150 per second.

