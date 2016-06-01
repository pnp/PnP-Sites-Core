var headID = document.getElementsByTagName('head')[0];               
var newScript = document.createElement('script');               
newScript.type = 'text/javascript';
newScript.src = '{InfrastructureSiteUrl}/Style%20Library/SP.Responsive.UI/{ScriptName}?rev=bf19e4f64b204e1ebc2f762e33afcc97';
newScript.id = 'PnPResponsiveUI';
headID.appendChild(newScript);