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
Download demo: [lssdemo.nsf](../master/demo/lsdldemo20160918.zip)

### Screenshots
Stack trace in a messagebox:

![Demo 1 - error details](../master/demo/lsdl_demo1_1.gif "Error stack report")

Trace report with timings in a messagebox:

![Demo 1 - trace report](../master/demo/lsdl_demo1_2.gif "Trace report")

Logging redirected to a file with error stack and tracing info:

![Demo 1 - logging example](../master/demo/lsdl_demo1_3.gif "Logging")

Detailed messages in stack:

![Demo 2 - error details](../master/demo/lsdl_demo2_1.gif "Error stack report")

Logging redirected to status bar:

![Demo 2 - logging to status bar](../master/demo/lsdl_demo2_2.gif "Logging")
