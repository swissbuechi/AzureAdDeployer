function Get-AADTableOfContents {
    return @"
   <h3 class='TOC'><a href="#AAD">Azure Active Directory</a></h3>
   <ul>
       <li><a href="#AAD_USER_SETTINGS">User settings</a></li>
       <li><a href="#AAD_DEVICE_JOIN_SETTINGS">Device join settings</a></li>
       <li><a href="#AAD_SKU">Licenses</a></li>
       <li><a href="#AAD_ADMINS">Admin role assignments</a></li>
       <li><a href="#AAD_BG">BreakGlass account</a></li>
       <li><a href="#AAD_MFA">User MFA status</a></li>
       <li><a href="#AAD_GUEST">Guest accounts</a></li>
       <li><a href="#AAD_SEC_DEFAULTS">Security defaults</a></li>
       <li><a href="#AAD_CA">Conditional access policies</a></li>
       <li><a href="#AAD_CA_LOCATIONS">Named locations</a></li>
       <li><a href="#AAD_APP_POLICY">App protection policies</a></li>
   </ul>
"@
}
function Get-SPOTableOfContents {
    return @"
   <h3 class='TOC'><a href="#SPO">SharePoint Online</a></h3>
   <ul>
       <li><a href="#SPO_SETTINGS">Tenant settings</a></li>
   </ul>
"@
}
function Get-EXOTableOfContents {
    return @"
   <h3 class='TOC'><a href="#EXO">Exchange Online</a></h3>
   <ul>
       <li><a href="#EXO_DOMAIN">Domains</a></li>
       <li><a href="#EXO_CONNECTOR_IN">Inbound mail connector</a></li>
       <li><a href="#EXO_CONNECTOR_OUT">Outbound mail connector</a></li>
       <li><a href="#EXO_USER">User mailbox</a></li>
       <li><a href="#EXO_SHARED">Shared mailbox</a></li>
       <li><a href="#EXO_UNIFIED">Unified mailbox</a></li>
   </ul>
"@
}