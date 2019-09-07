<%
function ListAdoErrors(){
	result=""
	for(i=0;i<Connect.Errors.Count;i++){
		result+=Connect.Errors(i).Description+"<br>"
	}
	return result
}
%>