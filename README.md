# LotusScript Developer Layer (ls-dl)

Error handling, logging, tracing and unit testing in Lotus Script.
Easy, flexible, extensible.

Snippet error handling:
```Visual basic
Sub test()
	On Error GoTo catch
	
	
	
	GoTo finally
catch:
	throwException
finally:
End Sub
```
