import groovy.json.JsonSlurper 

def response = messageExchange.response.responseContent
def slurper = new JsonSlurper()
def json = slurper.parseText response

assert json.result.severityResultResponse.simpleTypeClass3=="xxxxxxxx"
assert json.result.resultParameterList.resultParameter[0].resultParameterName=="xxxxxxx"
    
