<%
' Comersus Sophisticated Cart 
' Comersus Open Technologies
' USA - 2006
' http://www.comersus.com 
' Details: Referrer parsing 
%>
<%
sub referrerParsing(referrer, searchEngine, searchKeywords)

  referrer=lcase(referrer)
  
  if instr(referrer,"google")<>0 then
  	searchEngine		="Google"   
  	pStartingPosition	=instr(referrer,"q=")
  	
  	if pStartingPosition>0 then  	
  	 ' sum &q=
  	 pStartingPosition=pStartingPosition+2
  	 
  	 pEndingPosition		=instr(pStartingPosition+1,referrer,"&")
  	 
  	 if pEndingPosition=0 then
  	  searchKeywords		=mid(referrer,pStartingPosition)  	  	  
  	 else
  	  searchKeywords		=mid(referrer,pStartingPosition,pEndingPosition-pStartingPosition)  	  	  
  	 end if
  	 
  	end if
  	  	
  end if
  
  if instr(referrer,"msn.com")<>0 then
  	searchEngine		="MSN"   
  	pStartingPosition	=instr(referrer,"&q=&q=")
  	
  	if pStartingPosition>0 then  	
  	 ' sum &q=
  	 pStartingPosition=pStartingPosition+6
  	 
  	 pEndingPosition		=instr(pStartingPosition+1,referrer,"&")
  	 
  	 if pEndingPosition=0 then
  	  searchKeywords		=mid(referrer,pStartingPosition)  	  	  
  	 else
  	  searchKeywords		=mid(referrer,pStartingPosition,pEndingPosition-pStartingPosition)  	  	  
  	 end if
  	 
  	end if
  	  	
  end if
  
  if instr(referrer,"altavista.com")<>0 then
  	searchEngine		="Altavista"   
  	pStartingPosition	=instr(referrer,"q=")
  	
  	if pStartingPosition>0 then  	
  	 ' sum &q=
  	 pStartingPosition=pStartingPosition+2
  	 
  	 pEndingPosition		=instr(pStartingPosition+1,referrer,"&")
  	 
  	 if pEndingPosition=0 then
  	  searchKeywords		=mid(referrer,pStartingPosition)  	  	  
  	 else
  	  searchKeywords		=mid(referrer,pStartingPosition,pEndingPosition-pStartingPosition)  	  	  
  	 end if
  	 
  	end if
  	  	
  end if
  
  if instr(referrer,"aol.co")<>0 then
  	searchEngine		="Aol"   
  	pStartingPosition	=instr(referrer,"query=")
  	
  	if pStartingPosition>0 then  	
  	 ' sum &q=
  	 pStartingPosition=pStartingPosition+6
  	 
  	 pEndingPosition		=instr(pStartingPosition+1,referrer,"&")
  	 
  	 if pEndingPosition=0 then
  	  searchKeywords		=mid(referrer,pStartingPosition)  	  	  
  	 else
  	  searchKeywords		=mid(referrer,pStartingPosition,pEndingPosition-pStartingPosition)  	  	  
  	 end if
  	 
  	end if
  	  	
  end if
  
  if instr(referrer,"mama.com")<>0 or instr(referrer,"mamma.com")<>0 then
  	searchEngine		="Mama"   
  	pStartingPosition	=instr(referrer,"query=")
  	
  	if pStartingPosition>0 then  	
  	 ' sum &q=
  	 pStartingPosition=pStartingPosition+6
  	 
  	 pEndingPosition		=instr(pStartingPosition+1,referrer,"&")
  	 
  	 if pEndingPosition=0 then
  	  searchKeywords		=mid(referrer,pStartingPosition)  	  	  
  	 else
  	  searchKeywords		=mid(referrer,pStartingPosition,pEndingPosition-pStartingPosition)  	  	  
  	 end if
  	 
  	end if
  	  
  end if
  
  if instr(referrer,"yahoo.co")<>0 then
  	searchEngine		="Yahoo"   
  	pStartingPosition	=instr(referrer,"p=")
  	
  	if pStartingPosition>0 then  	
  	 ' sum &q=
  	 pStartingPosition=pStartingPosition+2
  	 
  	 pEndingPosition		=instr(pStartingPosition+1,referrer,"&")
  	 
  	 if pEndingPosition=0 then
  	  searchKeywords		=mid(referrer,pStartingPosition)  	  	  
  	 else
  	  searchKeywords		=mid(referrer,pStartingPosition,pEndingPosition-pStartingPosition)  	  	  
  	 end if
  	 
  	end if
  	  	
  end if
  
  if instr(referrer,"metacrawler.com")<>0 then
  	searchEngine		="MetaCrawler"   
  	pStartingPosition	=instr(referrer,"web/")
  	
  	if pStartingPosition>0 then  	
  	 ' sum &q=
  	 pStartingPosition=pStartingPosition+4
  	 
  	 pEndingPosition		=instr(pStartingPosition+1,referrer,"&")
  	 
  	 if pEndingPosition=0 then
  	  searchKeywords		=mid(referrer,pStartingPosition)  	  	  
  	 else
  	  searchKeywords		=mid(referrer,pStartingPosition,pEndingPosition-pStartingPosition)  	  	  
  	 end if
  	 
  	end if
  	  	
  end if
        
  
  if instr(referrer,"ask.com")<>0 then
  	searchEngine		="Ask.com"   
  	pStartingPosition	=instr(referrer,"q=")
  	
  	if pStartingPosition>0 then  	
  	 ' sum &q=
  	 pStartingPosition=pStartingPosition+2
  	 
  	 pEndingPosition		=instr(pStartingPosition+1,referrer,"&")
  	 
  	 if pEndingPosition=0 then
  	  searchKeywords		=mid(referrer,pStartingPosition)  	  	  
  	 else
  	  searchKeywords		=mid(referrer,pStartingPosition,pEndingPosition-pStartingPosition)  	  	  
  	 end if
  	 
  	end if
  	  	
  end if
  
  if instr(referrer,"vivisimo.com")<>0 then
  	searchEngine		="Vivisimo.com"   
  	pStartingPosition	=instr(referrer,"query=")
  	
  	if pStartingPosition>0 then  	
  	 ' sum &q=
  	 pStartingPosition=pStartingPosition+6
  	 
  	 pEndingPosition		=instr(pStartingPosition+1,referrer,"&")
  	 
  	 if pEndingPosition=0 then
  	  searchKeywords		=mid(referrer,pStartingPosition)  	  	  
  	 else
  	  searchKeywords		=mid(referrer,pStartingPosition,pEndingPosition-pStartingPosition)  	  	  
  	 end if
  	 
  	end if
  	  	
  end if
   	
end sub
%>