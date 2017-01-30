<%
' Comersus Sophisticated Cart 
' Comersus Open Technologies
' USA - 2006
' http://www.comersus.com 
' Details: RC4 encryption functions

%>
<%    

function encrypt(text, password)
 if pEncryptionMethod="DES" then
  encrypt=DESEncrypt(text, password)
 else
  encrypt=RC4EnCryptASC(text, password)
 end if
end function

function decrypt(text, password)
 if pEncryptionMethod="DES" then
  decrypt=DESDecrypt(text, password)
 else
  decrypt=RC4DeCryptASC(text, password)
 end if
end function


 ' This script performs 'RC4' Stream Encryption (Based on what is widely thought to be RSA's RC4 algorithm. It produces output streams that are identical to the commercial products). This script is Copyright © 1999 by Mike Shaffer ALL RIGHTS RESERVED WORLDWIDE                              

   Dim sbox(255)
   Dim rc4Key(255)

   Sub RC4Initialize(strPwd)
   
   ' this routine called by EnDeCrypt function. Initializes the sbox and the key array

      dim tempSwap, a, b      
      
      ' get length of the key
      intLength = len(strPwd)
      
      ' iterate through all characters contained in key repeating number of characters is 255
      for a = 0 To 255
         ' load ANSI for each char contained in the key
         rc4Key(a)  = asc(mid(strpwd, (a mod intLength)+1, 1))
         ' load numbers from 0 to 255
         sbox(a) = a                       
      next

      b = 0
      ' iterate through arrays 
      for a = 0 To 255
         b = (b + sbox(a) + rc4Key(a)) Mod 256         
         tempSwap = sbox(a)
         sbox(a) = sbox(b)
         sbox(b) = tempSwap
      Next
   
   End Sub
   
   function EnDeCrypt(plaintxt, psw)      

      dim temp, a, i, j, k, cipherby, cipher      

      i = 0
      j = 0

      RC4Initialize psw

      for a = 1 To Len(plaintxt)
         i 	= (i + 1) Mod 256
         j 	= (j + sbox(i)) Mod 256
         temp 	= sbox(i)
         sbox(i)= sbox(j)
         sbox(j)= temp
   
         k 	= sbox((sbox(i) + sbox(j)) Mod 256)

         cipherby 	= Asc(Mid(plaintxt, a, 1)) Xor k
         cipher 	= cipher & Chr(cipherby)         
      next

      enDeCrypt = cipher            
      
   end function
   
   function RC4EnCryptASC(plaintxt, psw)      

      dim temp, a, i, j, k, cipherby, cipher      

      i = 0
      j = 0

      RC4Initialize psw

      for a = 1 To Len(plaintxt)
         i 	= (i + 1) Mod 256
         j 	= (j + sbox(i)) Mod 256
         temp 	= sbox(i)
         sbox(i)= sbox(j)
         sbox(j)= temp
   
         k 		= sbox((sbox(i) + sbox(j)) Mod 256)

         cipherby 	= Asc(Mid(plaintxt, a, 1)) Xor k         
         cipher 	= cipher & "|"& cipherby         
      next            
      
      RC4EnCryptASC = cipher            
      
   end function
   
   function RC4DeCryptASC(plaintxt, psw)      

      plaintxt = transformToCHR(plaintxt)
      
      dim temp, a, i, j, k, cipherby, cipher      

      i = 0
      j = 0      
      
      dim arrayEncrypted            
      
      RC4Initialize psw         
      
       for a = 1 To Len(plaintxt)
         i 	= (i + 1) Mod 256
         j 	= (j + sbox(i)) Mod 256
         temp 	= sbox(i)
         sbox(i)= sbox(j)
         sbox(j)= temp
   
         k 	= sbox((sbox(i) + sbox(j)) Mod 256)

         cipherby 	= Asc(Mid(plaintxt, a, 1)) Xor k
         cipher 	= cipher & Chr(cipherby)         
      next

      RC4DeCryptASC = cipher                       
      
   end function
   
function transformToCHR(plaintxt)

      ' transform to CHR, insert into array
      
      dim returnText, arrayEncrypted
      arrayEncrypted = split(plaintxt, "|")            
      
      returnText=""            
      
      ' transform to CHR
      for a = 1 to ubound(arrayEncrypted)         
         returnText=returnText&CHR(arrayEncrypted(a))         
      next            
      
      transformToCHR = returnText
end function


%>