# LotusScript Developer Layer (ls-dl)

Error handling, logging, tracing and unit testing in Lotus Script.
Easy, flexible, extensible.

## Error handling
```lss
Sub test()
	On Error GoTo catch
	
	'...
	
	GoTo finally
catch:
	throwException
finally:
End Sub
```

## Tracing
```lss
Sub test()
	traceIn
	On Error GoTo catch
	
	'...
	
	GoTo finally
catch:
	traceOut
	throwException
finally:
	traceOut
End Sub
```

## Logging
```lss
	logAssert ""
	logError ""
	logWarn ""
	logInfo ""
	logDebug ""
	logVerbose ""
```

## Demo
[lssdemo.nsf](../master/demo/lsdldemo20160918.zip)

![alt text](../master/demo/lsdl_demo1_1.gif "Error stack report")
