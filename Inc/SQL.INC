<%
var goodslstsql="Select t1.*, t2.rate, t3.rate as unrate, t3.VALNAME from GOODS t1, valuta t2, valuta t3 where t1.service_id=%SRV and t2.id=t1.valuta_id and t3.id=t1.valuta_id_show Order by t1.name"

var insgoods="Insert into GOODS (ID,NAME,PRICE,PRICE_1,PRICE_2,PRICE_S,PRICE_DATE,URL,EXISTENCE,VALUTA_ID,VALUTA_ID_SHOW,VALNAME_ID,SERVICE_ID,ARTICUL,IS_SERVICE) "+
             "Values (%ID,'%NAM',%PR,%PR1,%PR2,%PRS,'TODAY','%URL',%EXT,%VAL,%VALS,%VALN,%SERV,'%ART',%ISSRV)"

var getgoods_hid_divis="Select t1.service_id as hid, t2.division_id as divis from goods t1, service t2 where t1.service_id=t2.id and t1.id="

var edgoods="UPDATE GOODS "+
            "SET NAME='%NAM', URL='%URL', EXISTENCE=%EXT, VALUTA_ID=%VAL, VALUTA_ID_SHOW=%VALS, VALNAME_ID=%VALN, ARTICUL='%ART', IS_SERVICE=%ISSRV "+
            "WHERE ID=%ID"

var edprice="UPDATE GOODS "+
            "SET PRICE=%PR, PRICE_1=%PR1, EXISTENCE=%EXT, PRICE_2=%PR2, PRICE_S=%PRS, PRICE_DATE='TODAY' "+
            "WHERE ID=%ID"

var insservice="Insert into SERVICE (ID,NAME,HI_ID,URL,DIVISION_ID)"+
               "VALUES (%ID, '%NAM', %HID, '%URL', %DIVIS)"

var insvaluta="Insert into VALUTA (ID, VALNAME, VALUTA_TYPE_ID, ENTERPRISE_ID, RATE, RATE_DATE) "+
              "VALUES (%ID, '%NAM', %VALID, %COMP, %RATE, 'TODAY')"

var newratevaluta="Update VALUTA set RATE=%RATE, RATE_DATE='TODAY' where ID=%ID"

var insvalname="Insert into VALNAME(ID,NAME) Values(%ID,'%NAME')"

var edvalname="Update VALNAME Set NAME='%NAME' Where ID=%ID"

var updatecompany="Update ENTERPRISE set city_id=%CITY, name='%NAME', address='%ADR', address_p='%ADRP', address_f='%ADRF', postindex='%IND', email='%EML', phone='%PHON', domen='%DOM', inn='%INN', okpo='%OKPO', okonh='%OKONH', rs='%RS', bank='%BANK', ks='%KS', bik='%BIK', bossfam='%BF', bossio='%BIO' where id=%ID"

var divpsw="Update DIVISION Set PSW_CLIENT='%PS' Where ID=%ID"

var divedit="Update DIVISION Set NAME='%NAME', MANAGER='%MAN', PHONE='%PHON', FAX='%FAX', ADDRESS='%ADR', EMAIL='%EML' Where ID=%ID"

var usrpsw="Update USERS Set PSW='%PS' Where ID=%ID"

var delgoods="Delete from GOODS Where ID=%ID"

var edserv="Update SERVICE set NAME='%NAME', URL='%URL' where ID=%ID"

var inspubusr="Insert into smi_usr(ID,NAME,NIK_NAME,PSW,SMI_ID,USR_TYPE_ID,STATE,EMAIL,PHONE) values (%ID, '%NAM', '%NIK', '%PS', %SMI, %TP, 0, '%EML', '%PHN')"

var insheading="Insert into heading(ID,SMI_ID,NAME,HI_ID,PICTURE,URL,PERIOD,PAGE_LENGTH,SMI_USR_ID,ISNEWS) "+
		"values(%ID,%SMI,'%NAM',%HID, '', '%URL', %PER, %PL, %USR, %ISNEWS)"

var edheading="Update heading set NAME='%NAM', URL='%URL', PERIOD=%PER, PAGE_LENGTH=%PL, SMI_USR_ID=%USR, ISNEWS=%ISNEWS where ID=%ID"

var inspub="Insert into publication(ID,NAME,DIGEST,DATE_CREATE,PUBLIC_DATE,STATE,HEADING_ID,COMENT,AUTOR,URL,CREATER_ID,ISHTML) "+
		"values(%ID,'%NAM','%DIG','NOW','%PDAT',%ST, %HID, '%COM', '%AUT', '%URL', %USR, %ISH)"

var insblock="Insert into block_news(ID,NAME,SUBJ,SMI_ID) values (%ID,'%NAM','%SUBJ',%SMI)"

var edpub="Update publication set NAME='%NAM',DIGEST='%DIG',PUBLIC_DATE='%PDAT',COMENT='%COM',AUTOR='%AUT',URL='%URL',ISHTML=%ISH where id=%ID"

var inscousr="Insert into COMPUSER(ID,FIO,NIKNAME,PSW,EMAIL,STATE) values (%ID, '%NAM', '%NIK', '%PS', '%EML', 0)"

var allpubheading="select distinct t1.id from publication t1, getbranchheading(%HID,%SMI) p1 where p1.hid=t1.heading_id Union Select distinct t2.id from publication t2 where t2.heading_id=%HID1"

var edblock="Update block_news set NAME='%NAM', SUBJ='%SUBJ' where ID=%ID and SMI_ID=%SMI"

%>