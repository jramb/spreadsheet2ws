<xsl:stylesheet xmlns:xsl="http://www.w3.org/1999/XSL/Transform" version="1.0">
  <xsl:output method="text" indent="no" omit-xml-declaration="yes" encoding="UTF-8"/>
  <xsl:strip-space elements="*"/>
  <xsl:template match="/">
    <xsl:for-each select="/root/row">
      <xsl:if test="c[@col='0']='NEW' or c[@col='0']='' or not(c[@col='0'])">
        <xsl:value-of select="c[@col='4']"/>
        <!--<xsl:value-of select="/root/row/c[@col='0']"/>-->
      </xsl:if>
    </xsl:for-each>
  </xsl:template>
</xsl:stylesheet>

