<%
function writeBox(x0,y0,x1,y1,header,fill,text1,text2,text3,text4)
'wx = larghezza finestra
'wy = altezza finestra 
dim wx,wy,criga 
   pdf.SetFont "Arial","",10
   wx = cdbl(x1)-cdbl(x0)
   wy = cdbl(y1)-cdbl(y0)
   
   pdf.Rect x0,y0 ,wx ,wy, fill
   if header<>"" then 
      pdf.SetFont "Arial","B",6
      pdf.Text x0,y0 - 1 , header 
   end if 
   pdf.SetFont "Arial","B",10
   cY = y0
   cX = x0 + 0.5
   if text1<>"" then 
      cY = CY + 5
  pdf.Text cX, cY , text1
   end if 
   if text2<>"" then 
      cY = CY + 3
  pdf.Text cX, cY , text2
   end if 

end function 

function writeBoxCell(x0,y0,x1,y1,header,fill,text1)
dim wx,wy,criga,fi,ws   
Dim wxMax 
    
   pdf.SetFont "Arial","",10
   wx = cdbl(x1)-cdbl(x0)
   wxMax = wx - 4
   wy = cdbl(y1)-cdbl(y0)
   pdf.Rect x0,y0 ,wx ,wy, fill
   if header<>"" then 
      pdf.SetFont "Arial","B",6
      pdf.Text x0,y0 - 1 , header
   end if 
   pdf.SetFont "Arial","B",10
   cY = y0
   cX = x0 + 0.5
   if text1<>"" then 
      'ciclo 
      ttw = ""
      lastSpace = 0
      do while text1<>"" 
         if len(text1)=1 then 
            ttw = ttw & Text1
            Text1 = ""
         else 
            ttw = ttw & Mid(Text1,1,1)
            Text1 = Mid(Text1,2)
         end if 
         if mid(ttw,len(ttw),1)=" " then 
            lastspace = len(ttw)
         end if 
         ws = pdf.GetStringWidth(ttw)
         if ws > wxMax or Text1="" then 
            cY = CY + 5
            if lastSpace<>0 then 
               pdf.Text cX, cY , mid(ttw,1,lastspace)
               ttw = mid(ttw,lastSpace+1)
            else 
               pdf.Text cX, cY , ttw
               ttw = ""
            end if
            lastspace = 0
         end if 
      loop 
   end if 

end function 
%>