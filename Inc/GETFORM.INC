<%
function FloatFormData(value){
	value=String(value)
	switch(value){
	case "undefined":value=""
	case "":  result=null; break
	default: result=parseFloat(value,10)
	}
	return result
}
function IntFormData(value){
	value=String(value)
	switch(value){
	case "undefined":value=""
	case "":  result=null; break
	default: result=parseInt(value,10)
	}
	return result
}
function StrFormData(value){
	value=String(value)
	switch(value){
	case "undefined":value=""
	case "":  result=null; break
	default: result=value
	}
	return result
}

function TextFormData(value,def_value){
	result=StrFormData(value)
	result=result==null?def_value:result
	return result
}

%>