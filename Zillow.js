// ZILLOW API has been deprecated / is no longer accessible. Module retained for reference.

/**
 * Spreadsheet function for Zillow API
 */
function ZILLOWSUMMARY(street, cityStateZip) {
  var zillow = CacheService.getScriptCache().get('zillow');
  var params = CacheService.getScriptCache().get('params');
  if (zillow != null & params != null) {
    if ([street, cityStateZip] == params['address']) {
      return [['Zestimate', zillow.price],
          ['Comparables', zillow.comp]];
    }
  }
  var zillow = getZillow_(street, cityStateZip, 25);
  CacheService.getScriptCache().put('zillow', zillow, 21600);
  CacheService.getScriptCache().put('params', {'address': [street, cityStateZip]}, 21600);
    
  return [['Zestimate', zillow.price],
          ['Comparables', zillow.comp]];
}

/**
 * Helper function for Zillow API
 */
function getZillow_(stAddress,cityStateZip,num) { 
  
  // API Example: http://www.zillow.com/webservice/GetSearchResults.htm?zws-id=<ZWSID>&address=2114+Bigelow+Ave&citystatezip=Seattle%2C+WA
  // GAS Zillow App Example: https://raw.githubusercontent.com/TillerHQ/tiller-zillow-simple/master/zillow.js
  var heads = {};
  heads['zws-id'] = PROPS.getProperty('zws-id');
  heads.address = encodeURIComponent(stAddress);
  heads.citystatezip = encodeURIComponent(cityStateZip);
  
  // Run GetSearchResults with address to get ZPID
  var baseURL = 'http://www.zillow.com/webservice/GetSearchResults.htm';
  var fullURL = baseURL + '?';
  Object.keys(heads).forEach(k => fullURL += k + '=' + heads[k] + '&');
  fullURL = fullURL.slice(0,-1);
  var resp = UrlFetchApp.fetch(fullURL);
  var searchXML = XmlService.parse(resp);
  var zpid = searchXML.getRootElement().getChild('response').getChild('results').getChild('result').getChild('zpid');
  
  // Run GetDeepComps with ZPID to get comparable homes
  baseURL = 'http://www.zillow.com/webservice/GetDeepComps.htm';
  resp = UrlFetchApp.fetch(baseURL+'?zws-id='+heads['zws-id']+'&zpid='+zpid.getValue()+'&count='+num);
  var respXml = XmlService.parse(resp);
  var fmtRsp = XmlService.getPrettyFormat().format(respXml);
  var properties = respXml.getRootElement().getChild('response').getChild('properties');
  
  // Get property zestimate
  var principal = properties.getChild('principal');
  var zest = parseInt(principal.getChild('zestimate').getChild('amount').getValue());
  
  // Calculate average of comps with estimates
  var comps = properties.getChild('comparables').getChildren().filter(c => c.getChild('zestimate').getChild('amount').getValue() != '');
  var avgComp = 0;
  comps.forEach(c => avgComp += parseInt(c.getChild('zestimate').getChild('amount').getValue()));
  avgComp = Math.round(avgComp / comps.length);
  
  var zData = {
    'price': zest.toLocaleString(),
    'comp': avgComp.toLocaleString(),
    'appreciation': (zest - 390000).toLocaleString()
  };
  
  return zData;
}