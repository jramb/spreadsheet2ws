<xsl:stylesheet xmlns:xsl="http://www.w3.org/1999/XSL/Transform" version="1.0"
  xmlns:env="http://schemas.xmlsoap.org/soap/envelope/"
  xmlns:inv="http://invoice.ws.fmt/v8"
  exclude-result-prefixes="env,inv,inv">
  <xsl:template match="/">
    <xsl:apply-templates select="//env:Fault"/>
    <xsl:apply-templates select="//inv:Invoice"/>
  </xsl:template>
  <xsl:template match="//env:Fault">
    <row>
      <c col="0">Error</c>
      <c col="1">
        <xsl:choose>
          <xsl:when test="detail/inv:Text">
            <xsl:value-of select="detail/inv:Text"/>
          </xsl:when>
          <xsl:when test="faultstring!=''">
            <xsl:value-of select="faultstring"/>
          </xsl:when>
          <xsl:otherwise>
            (see log)
          </xsl:otherwise>
        </xsl:choose>
      </c>
    </row>
  </xsl:template>
  <xsl:template match="//inv:Invoice[1]">
    <row>
      <c col="0">OK</c>
      <c col="1"><xsl:value-of select='inv:Supplier'/></c>
      <c col="2"><xsl:value-of select='inv:ID[@schemeName="InvoiceNumber"]'/></c>
      <c col="3"><xsl:value-of select='inv:ID[@schemeName="Amount"]'/></c>
    </row>
  </xsl:template>
</xsl:stylesheet>

