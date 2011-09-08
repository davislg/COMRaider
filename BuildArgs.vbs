' "parent" object is the calling instance of the frmTlbViewer form
' lngs, ints, dbls, and strs are all collections used to hold the 
' argument variations we wish to inject into the COM objects methods
' which we are fuzzing for a given variable type. The strs collection
' is used for both BSTR and VARIANT types atm.. Mabey I should make
' variants cycle through all collections well see..

function GetLongArgs()
	parent.lngs.add 2147483647
	parent.lngs.add -1
	parent.lngs.add 0
	parent.lngs.add -2147483647
end function

function GetIntArgs()
	parent.ints.add 32767
	parent.ints.add -1
	parent.ints.add 0
	parent.ints.add -32768
end function

function GetDblArgs()
	parent.dbls.add 1.79769313486231E+308
	parent.dbls.add -1
	parent.dbls.add 0
	parent.dbls.add 3.39519326559384E-313
end function

function GetStrArgs()
    
    'you can use script to build arg too
    for i=1024 to 15000 step 1024
	parent.strs.add "String(" & (i+20) & ", ""A"")"
    next

    parent.strs.add """%n%n%n%n%n%n%n%n%n%n%n%n%n%n%n%n%n%n%n%n%n%n%n%n%n%n%n%n%n%n%n%n%n%n%n%n"""
    parent.strs.add """%s%s%s%s%s%s%s%s%s%s%s%s%s%s%s%s%s%s%s%s%s%s%s%s%s%s%s%s%s%s%s%s%s%s%s%s"""

    parent.strs.add """http://test\test\test\te?s\test\test\tes\ttest\test\te@st\tes\test\test\tes.\ttest\test\test\tes\test\test\te.s\ttest\test\test\tes\test\test\tes\t\\\\\\\\\:#$%test\test\test\te?s\test\test\tes\\:#$%\ttest\test\te@st\tes\test\test\tes.\ttest\test\test\tes\test\test\te.s\ttest\test\test\tes\test\test\tes\t\\\\\\\\\:#$%test\test\test\te?s\test\test\tes\\:#$%\ttest\test\te@st\tes\test\test\tes.\ttest\test\test\tes\test\test\te.s\ttest\test\test\tes\test\test\tes\t\\\\\\\\\:#$%test\test\test\te?s\test\test\tes\\:#$%\ttest\test\te@st\tes\test\test\tes.\ttest\test\test\tes\test\test\te.s\ttest\test\test\tes\test\test\tes\t\\\\\\\"""

end function


'lets you add custom tests based on variable names (strings and variants only)
'varName will be a lcase copy of the name used in teh typelib for variable x
function CustomStrArg(varName)

	if instr(varName,"url") > 0 or instr(varName,"uri") > 0 then 
	    parent.cust.add """http://www."" + string(2040,""A"") + "".com/"""
    	    parent.cust.add """http://"" + string(2040,""A"") + "":test@www.jfdkljfksljfl.com/"""	
    	    parent.cust.add """http://test:"" + string(2040,""A"") + ""@www.jfdkljfksljfl.com/"""
    	    parent.cust.add """http://www.jfdkljfksljfl.com/"" + string(2040,""A"") + "".txt"""	
	end if 

end function

