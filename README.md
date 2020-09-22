<div align="center">

## SetColorBar


</div>

### Description



'  Creates a color bar background for a ListView when in

'  report mode. Passing the listview and picturebox allows

'  you to use this with more than one control. You can also

'  change the colors used for each by passing new RGB color

'  values in the optional color parameters.
 
### More Info
 


'  Required - cListView As ListView

'  Required - cColorBar As PictureBox

'  Optional - lColor1 As Long

'  Optional - lColor2 As Long



'  Add the following line of code to your program,

'  replacing "lvListView" and "picColorBar" with the

'  names of your own control values. The color values

'  are optional; while the default is Green/White,

'  these create gray bars.

'  SetColorBar lvListView, picColorBar, &HC0C0C0, &H808080



'  Sets ListView Picture to none if not in report

'  mode or on error condition.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[N/A](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/empty.md)
**Level**          |Unknown
**User Rating**    |4.2 (67 globes from 16 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Miscellaneous](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/miscellaneous__1-1.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/setcolorbar__1-2746/archive/master.zip)





### Source Code

```
Public Sub SetColorBar(cListView As ListView, cColorBar As PictureBox, Optional lColor1 As Long = &HE2F1E3, Optional lColor2 As Long = vbWhite)
' Creates a color bar background for a ListView when in
' report mode. Passing the listview and picturebox allows
' you to use this with more than one control. You can also
' change the colors used for each by passing new RGB color
' values in the optional color parameters.
 Dim iLineHeight As Long
 Dim iBarHeight As Long
 Dim lBarWidth As Long
 On Error GoTo SetColorBarError
 '  set picture to none and exit sub if not in report mode
 If Not cListView.View = lvwReport Then GoTo SetColorBarError
 '  these can be commented out if the cColorBar control
 '  is set correctly.
 cColorBar.AutoRedraw = True
 cColorBar.BorderStyle = vbBSNone
 cColorBar.ScaleMode = vbTwips
 cColorBar.Visible = False
 '  set the alignment to "Tile" and you only need
 '  two bars of color.
 cListView.PictureAlignment = lvwTile
 '  needed because ListView does not have "TextHeight"
 cColorBar.Font = cListView.Font
 '  set height to a single line of text plus a
 '  one pixel spacer.
 iLineHeight = cColorBar.TextHeight("|") + Screen.TwipsPerPixelY
 '  set color bars to 3-line wide.
 iBarHeight = iLineHeight * 3
 lBarWidth = cListView.Width
 '  resize the cColorBar picturebox
 cColorBar.Height = iBarHeight * 2
 cColorBar.Width = lBarWidth
 '  paint the two bars of color
 cColorBar.Line (0, 0)-(lBarWidth, iBarHeight), lColor1, BF
 cColorBar.Line (0, iBarHeight)-(lBarWidth, iBarHeight * 2), lColor2, BF
 '  set the cListView picture to the
 '  cColorBar image
 cListView.Picture = cColorBar.Image
 Exit Sub
SetColorBarError:
 '  clear cListView's picture and then exit
 cListView.Picture = LoadPicture("")
End Sub
```

