* WSAA Ticket Acceso
LOCAL _dcert, d_file

**
* Ver si el TA está vigente
_dcert = Directorio donde se encuentran el Certificado y la Llave
_dfile = _dcert + "ta_fe.xml"
IF FILE(_dfile)
	s_ta = FILETOSTR(_dfile)
	s_ta = STRTRAN(s_ta, '&lt;', '<')			&& Transformo la respuesta en XML
	s_ta = STRTRAN(s_ta, '&gt;', '>')
	t_exp = CTOT(EXTRAE_XML(s_ta, 'expirationTime'))
	ta_rt = EXTRAE_XML(s_ta, 'loginCmsReturn')
ELSE 
	t_exp = CTOT('')
ENDIF 

IF t_exp > DATETIME() + 60
	RETURN ta_rt
ENDIF 

LOCAL _dopen
_dopen = Directorio donde se encuentra OpenSSL.exe  'ej: c:\OpenSSL-Win32\bin'	

LOCAL lcNow, lcExp, lcUniqueId, lcXml, lcResponse, loHttp

* Generación del UniqueId 
lcUniqueId = ALLTRIM(STR(INT(SECONDS() * 1000), 10))

* Generación y expiración del ticket YYYY-MM-DDThh:mm:ss
lcNow = ALLTRIM(TTOC(DATETIME() - 120, 3))  
lcExp = ALLTRIM(TTOC(DATETIME() + 180, 3)) 	&& Expira en minutos nn
* Creación del XML de solicitud
lcXml = '<loginTicketRequest>' + ;
			'<header>' +;
				'<uniqueId>' + lcUniqueId + '</uniqueId>' + ;
				'<generationTime>' + lcNow + '</generationTime>' + ;
				'<expirationTime>' + lcExp + '</expirationTime>' + ;
			'</header>' + ;
			'<service>wsfe</service>' + ;
		'</loginTicketRequest>'

* Guardar el archivo XML
_dfile = _dcert + "ticket.xml"
IF FILE(_dfile)
    ERASE &_dfile
ENDIF 
STRTOFILE(lcXml, _dfile)

* Verificar y eliminar archivos previos
_dfile = _dcert + "ticket_firmado.cms"
IF FILE(_dfile)
    ERASE &_dfile
ENDIF 

* busco los archivos .key y .crt C:\GESTION\nn
DIMENSION _sgnr(1)
DIMENSION _crtf(1)
ADIR(_sgnr, _dcert + '*.key')     && no debe haber más de un certificado en la carpeta
ADIR(_crtf, _dcert + '*.crt')

IF EMPTY(ALLTRIM(_sgnr(1))) .or. EMPTY(ALLTRIM(_crtf(1)))
	MESSAGEBOX('Falta certificado y/o llave',16,'No se puede facturar',5000)
	RETURN ''
ENDIF 

_cpem = STRTRAN(UPPER(_crtf(1)),'.CRT','.PEM')

IF !FILE(_dcert + _cpem)    && En caso que no se encuentre el certificao ".PEM"
	_run = '"' + _dopen + 'openssl" x509 -in "' + _dcert + _crtf(1) + '" -out "' + _dcert + _cpem + '" -outform PEM'
	RUN /N7 &_run 
	_sleep(200)
ENDIF 

* Generar el CMS firmado con OpenSSL
_run = '"' + _dopen + 'openssl" cms -sign -in "' + _dcert + 'ticket.xml" -out "' + _dcert + 'ticket_firmado.cms" -signer "' + _dcert + _cpem + '" -inkey "' + _dcert + _sgnr(1) + '" -nodetach -outform PEM'

RUN /N7 &_run 

_ctr = 0
_dfile = _dcert + 'ticket_firmado.cms'

DO WHILE .t.
	_sleep(100)
	IF FILE(_dfile)
		EXIT
	ELSE 
		_ctr = _ctr + 1
		IF _ctr > 10
			MESSAGEBOX('No se pudo generar el ticket firmado',48,'',5000)
			RETURN ''
		ENDIF 
		LOOP
	ENDIF 
ENDDO 

* Leer el CMS firmado
tcTRA = ALLTRIM(FILETOSTR(_dfile))

* Ajuste de formato del contenido del CMS firmado
tcTRA = SUBSTR(tcTRA, 22, LEN(tcTRA) - 41)

* URL del servicio WSAA
tcWsdlUrl = ALLTRIM(RESTPARAM('URLWSAA',"https://wsaahomo.afip.gov.ar/ws/services/LoginCms"))

* Crear el objeto HTTP
loHttp = CREATEOBJECT("MSXML2.ServerXMLHTTP.6.0")
loHttp.Open("POST", tcWsdlUrl, .F.)
loHttp.SetRequestHeader("Content-Type", "application/xml")
loHttp.SetRequestHeader("SOAPAction", "http://ar.gov.afip.dif.facturaelectronica/")
loHttp.SetRequestHeader("User-Agent", "Mozilla/5.0")  && Mejora compatibilidad

ta_rt = ''

TRY 

    lcXmlSoap = ;
	'<soapenv:Envelope xmlns:soapenv="http://schemas.xmlsoap.org/soap/envelope/" xmlns:wsaa="http://wsaa.view.sua.dvadac.desein.afip.gov">' + ;
		"<soapenv:Header/>" + ;
	    "<soapenv:Body>" + ;
	    	"<wsaa:loginCms>" + ;
	        	"<wsaa:in0>" + tcTRA + "</wsaa:in0>" + ;
	    	"</wsaa:loginCms>" + ;
	   "</soapenv:Body>" + ;
	"</soapenv:Envelope>"


    * Enviar la solicitud
    loHttp.Send(lcXmlSoap)

    * Capturar la respuesta del servidor
    lcResponse = loHttp.ResponseText
	
    * Eliminación de archivos temporales
	_dfile = _dcert + "ticket.xml"
	IF FILE(_dfile)
	    ERASE &_dfile
	ENDIF 

	_dfile = _dcert + "ticket_firmado.cms"
	IF FILE(_dfile)
	    ERASE &_dfile
	ENDIF 

    * Guardar la respuesta en archivo
	IF '<loginCmsReturn>' $ lcResponse			&& Respuesta positiva
		_dfile = _dcert + "ta_fe.xml"

		STRTOFILE(lcResponse, _dfile)
		
		s_ta = STRTRAN(lcResponse, '&lt;', '<')			&& Transformo la respuesta en XML
		s_ta = STRTRAN(s_ta, '&gt;', '>')
		ta_rt = EXTRAE_XML(s_ta, 'loginCmsReturn')

	ENDIF 
	
	IF '<faultstring>' $ lcResponse
		_msg = EXTRAE_XML(lcResponse, 'faultstring')
		MESSAGEBOX(_msg, 48, 'Error al obtener Ticket',10000)
	ENDIF 
	
CATCH TO loException
    * Manejo de errores
    MESSAGEBOX(loException.Message, 48, 'Error', 10000)
ENDTRY

? ta_rt

* Extraigo del XML
FUNCTION extrae_xml
PARAMETERS _xml, _str
*LOCAL _lla, _ini, _fin
_lla = '<' + UPPER(ALLTRIM(_str)) + '>'
_ini = AT(_lla, UPPER(_xml))
IF _ini > 0
	_ini = _ini + LEN(_lla)
	_lla = '</' + ALLTRIM(_str) + '>'
	_fin = AT(_lla,_xml) - _ini
	RETURN SUBSTR(_xml, _ini, _fin)
ELSE
	RETURN ''
ENDIF 
ENDFUNC 

* Funciion de pausa en milisegundos
FUNCTION _Sleep(lnMiliSeg)
lnMiliSeg = IIF(TYPE("lnMiliSeg") = "N", lnMiliSeg, 1000)
DECLARE Sleep ;
  IN WIN32API ;
  INTEGER nMillisecs
RETURN Sleep(lnMiliSeg)
ENDFUNC

