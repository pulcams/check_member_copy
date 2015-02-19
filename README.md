check-member-copy
==============
###Under active development

A simple [AutoIT](https://www.autoitscript.com/site/) macro to help student workers quickly sort member copy to expedite processing. It does this by running cursory checks on bib records in the Voyager Cataloging module. Outputs MsgBox prompts and a csv log file (in the same directory as the script by default). The logic is as follows:
```
 Decision tree: 
			    complete[1]
				/     \
			  yes	  no - 'to the hold'
			   |        
		  full ELvL[2]     
	    	/    \
		  yes     no - 'ENCODING LEVEL'
           |         
       rel des.[3]   
         /  \
        yes  no - 'Relationship designators'    
        |
     'OK to process'
```
1. *Complete* means that the record has a single, full call no. and either has an LoC subject heading or is literature.
 - `050` must contain numbers
 - just one `050$a`
 - `050$b` is present
 - `090` if present has `$b`
 - `050$a` or `090$a` beginning with "P" are children's lit.
 - 008/33 if not blank = lit.
 - presence of 6xx_0 
2. Encoding level (ELvL) is full
	 - LDR/17 is not 3,5,8 or M
3. Relationship designators are present
	- 1xx if present has $e
	- 7xx if present has $e

Also, if `300` contains "online" the record has been mistakenly overlain; alert the human. 