<div align="center">

## EZ RegRead


</div>

### Description

My First Submission to PSCode.Com-

After searching all over PSC for a SIMPLE registry reading article, i couldnt find one. Although i did find a Registry write and thats where i got this code from and altered it to do reading instead of writting. Plz vote, any comments welcome!!

NOTE: To use this place this code in the Form Load and place a TextBox called Text1 on your form and you should be set...have any questions, post them and i will check them out later!!
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Brandon Pinnick](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/brandon-pinnick.md)
**Level**          |Beginner
**User Rating**    |4.0 (40 globes from 10 users)
**Compatibility**  |VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[Registry](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/registry__1-36.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/brandon-pinnick-ez-regread__1-39863/archive/master.zip)





### Source Code

```
Set wshshell = CreateObject("WScript.Shell")<br>
Text1.Text = wshshell.RegRead("HKEY_CURRENT_USER\Software\Microsoft\Internet Explorer\Start Page")
```

