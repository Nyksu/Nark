<%
function NextID(GeneratorName){
	if (String(GeneratorName)=="undefined"){GeneratorName="universal"}
	var recs=CreateRecordSet()
	recs.Source="select gen_id("+GeneratorName+",1) as id from dual"
	recs.Open()
	retval=recs("ID").Value
	recs.Close()
	delete recs
	return retval
}
%>