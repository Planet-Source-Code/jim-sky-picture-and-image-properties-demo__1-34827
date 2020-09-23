<div align="center">

## Picture and Image Properties Demo

<img src="GBT.jpg">
</div>

### Description

This is a demo to help sort out some the confusion about Image and Picture properties of a Picture control. I have been programming in VB for years and sometimes still get a bit confused about these so I wrote this to help myself as a reference, and thought perhaps some others might benefit. Also shows simplest usage of the BitBlt API function.
 
### More Info
 
Check the code for comments.


<span>             |<span>
---                |---
**Submitted On**   |2002-05-15 11:41:32
**By**             |[Jim Sky](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/jim-sky.md)
**Level**          |Beginner
**User Rating**    |5.0 (10 globes from 2 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Graphics](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/graphics__1-46.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[Picture\_an835845152002\.zip](https://github.com/Planet-Source-Code/jim-sky-picture-and-image-properties-demo__1-34827/archive/master.zip)

### API Declarations

```
Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
```





