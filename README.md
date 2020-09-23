<div align="center">

## IsFormLoaded?


</div>

### Description

A simple method to check if a form is loaded without referencing a property (such as .visible) that would in turn load the form.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Mike Douglas](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/mike-douglas.md)
**Level**          |Beginner
**User Rating**    |4.0 (28 globes from 7 users)
**Compatibility**  |VB 6\.0
**Category**       |[Miscellaneous](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/miscellaneous__1-1.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/mike-douglas-isformloaded__1-62582/archive/master.zip)





### Source Code

```
Public Function IsFormLoaded(ByVal FormName As String) As Boolean
  Dim frm As Form
  For Each frm In Forms
    If LCase(frm.name) = LCase(FormName) Then IsFormLoaded = True
  Next frm
End Function
```

