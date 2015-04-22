<xsl:stylesheet xmlns:xsl="http://www.w3.org/1999/XSL/Transform" version="1.0">
  <xsl:output method="xml"/>
  <xsl:template match="/root">
    <!--If your service requires some kind of soap-header, format this here-->
      <wsse:Security soap:mustUnderstand="0"
                     xmlns:wsse="http://docs.oasis-open.org/wss/2004/01/oasis-200401-wss-wssecurity-secext-1.0.xsd"
                     xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/">
           <wsse:UsernameToken wsu:Id="UsernameToken-2"
                               xmlns:wsu="http://docs.oasis-open.org/wss/2004/01/oasis-200401-wss-wssecurity-utility-1.0.xsd">
                <wsse:Username><xsl:value-of select="properties/entry[@key='username']"/></wsse:Username>
                <wsse:Password Type="http://docs.oasis-open.org/wss/2004/01/oasis-200401-wss-username-token-profile-1.0#PasswordText"><xsl:value-of select="properties/entry[@key='password']"/></wsse:Password>
           </wsse:UsernameToken>
      </wsse:Security>
  </xsl:template>
</xsl:stylesheet>



