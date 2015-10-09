import com.eviware.soapui.support.XmlHolder

def groovyUtils = new com.eviware.soapui.support.GroovyUtils( context )
def holder = groovyUtils.getXmlHolder( "Request-xml#ResponseAsXml" )

def simpleTypeClass3 = holder.getNodeValue("//urr:simpleTypeClass3")
assert simpleTypeClass3.equals("xxxxxxxxxxxx")

def resultCode = holder.getNodeValue("//urr:resultCode")
assert resultCode.equals("xxxxxxxxxxxxxx")

def resultDescription1=holder.getNodeValue("//urr:resultListi18nDescription//urr:descriptioni18nResult//urr:resultDescription")
assert resultDescription1.equals("xxxxxxxxxxxxxxxxxxxxxxxx")
