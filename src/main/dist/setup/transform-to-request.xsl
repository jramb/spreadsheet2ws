<xsl:stylesheet xmlns:xsl="http://www.w3.org/1999/XSL/Transform" version="1.0"
xmlns:java="http://xml.apache.org/xslt/java"
exclude-result-prefixes="java">
<xsl:template match="/">
  <LoadInvoice xmlns="http://invoice.ws.fmt/v8">
    <xsl:attribute name="environment">
      <xsl:value-of select="/root/properties/entry[@key='environment']"/> 
    </xsl:attribute>
    <Operator><xsl:value-of select="/root/properties/entry[@key='username']"/></Operator>
    <!--<CreationDate><xsl:value-of select="java:format(java:java.text.SimpleDateFormat.new('yyyy-MM-dd'), java:java.util.Date.new())"/></CreationDate>-->
    <CreationDate><xsl:value-of select="/root/properties/entry[@key='now']"/></CreationDate>
    <Batch>excel-load-<xsl:value-of select="/root/properties/entry[@key='uuid']"/></Batch>
    <SupplierID><xsl:value-of select="/root/row/c[@col='5']"/></SupplierID>
    <SupplierName><xsl:value-of select="/root/row/c[@col='4']"/></SupplierName>
    <InvoiceNumber><xsl:value-of select="/root/row/c[@col='7']"/></InvoiceNumber>
  </LoadInvoice>
</xsl:template>
</xsl:stylesheet>


