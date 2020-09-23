<div align="center">

## Javascipt and ASP


</div>

### Description

Javascript with ASP, this function is a example, how you can mix javascript and asp. The clue is it is clientside! In example a clientside calculator and need not to post anthing to the display.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Hohl David](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/hohl-david.md)
**Level**          |Intermediate
**User Rating**    |5.0 (10 globes from 2 users)
**Compatibility**  |ASP \(Active Server Pages\), HTML, VbScript \(browser/client side\)

**Category**       |[Coding Standards](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/coding-standards__4-33.md)
**World**          |[ASP / VbScript](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/asp-vbscript.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/hohl-david-javascipt-and-asp__4-7938/archive/master.zip)





### Source Code

```
function recalculate() {
				//Berechne mir die Ausgabe bei der Order
  totalsum = new Number();
 <%
 artSTMT = "SELECT * FROM article"
 Set artRS = nameConn.Execute(artStmt)
 while not artRS.eof
  artID = trim(artRS("artID"))
 %>
  price<%=artID%> = new Number(document.orderform.price<%=artID%>.value);
  quantity<%=artID%> = new Number(document.orderform.quantity<%=artID%>.value);
  linesum<%=artID%> = new Number(quantity<%=artID%> * price<%=artID%>);
  document.orderform.sumquantity<%=artID%>.value = linesum<%=artID%>;
  //document.orderform.quantity<%=artID%>.value = quantity<%=artID%>;
  totalsum = totalsum + linesum<%=artID%>;
 <%
  artRS.movenext()
 wend
 %>
 document.orderform.sumtotal.value = totalsum
 }
```

