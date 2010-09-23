<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform" xmlns:fo="http://www.w3.org/1999/XSL/Format">
<xsl:output method="html" indent="yes" encoding="UTF-8" />

<xsl:variable name="winquisitor_xsl_version">0.1</xsl:variable>

<xsl:template match="/winquisitor_audit">

  <html>
    <head>
    	<xsl:comment>
        generated with winquisitor.xsl - version 
        <xsl:value-of select="$winquisitor_xsl_version"/>
      </xsl:comment>
  	  <title>Winquisitor Report</title>
  	  <style type="text/css">
  	  body {
  	    font-family: Verdana, Helvetica, sans-serif;
        font-size: 9pt;
        color:#000000;
  	  }
  	  h1 {
  	    font-size: 14pt;
  	  }
  	  .scan_header {
  	    width: 100%;
  	    border-collapse: collapse; 
  	    border: 1px solid #000;
  	    font-family: Verdana, Helvetica, sans-serif;
        font-size: 9pt;
  	    text-align: left;
  	  }	
  	  .scan_header td {
  	    border: 1px solid #000;
  	    padding: 2px 4px 2px 4px;
  	  }
  	  .scan_header th {
  	    width: 10%;
  	    padding: 2px;
  	    border: 1px solid #000;
  	    background: #9BCDFF;
  	    font-weight: bold;
  	  }
  	  .scan_results {
  	    width: 100%;
  	    border-collapse: collapse; 
  	    border: 1px solid #000;
  	    font-family: Verdana, Helvetica, sans-serif;
        font-size: 9pt;
  	    text-align: left;
  	    margin-top: 20px;
  	    margin-bottom: 30px;
  	  }	
  	  .scan_results td {
  	    border: 1px solid #000;
  	    padding: 2px 4px 2px 4px;
  	  }
  	  .scan_results th {
  	    border: 1px solid #000;
  	    padding: 2px 4px 2px 4px;
  	    background: #C0C0C0;
  	  }
  	  .scan_results hr {
  	    border: none 0;
        border-top: 1px dashed #000;
        width: 99%;
        height: 1px;
  	  }
  	  .computer_header {
  	    padding: 2px;
  	    border: 1px solid #000;
  	    background: #E1E1E1;
  	  }

  	  </style>
    </head>
  	
    <body>
    	<h1>Winquisitor Results</h1>
      <xsl:for-each select="scan">
        <div class="scan">
        	<table class="scan_header">
        		<tr>
        			<th>Scan</th>
        			<td><xsl:value-of select="scan_info"/></td>
        		</tr>
        		<tr>
        			<th>Started</th>
        			<td><xsl:value-of select="start_time"/> on <xsl:value-of select="start_date"/></td>
        		</tr>
        		<tr>
        			<th>Ended</th>
        			<td><xsl:value-of select="end_time"/> on <xsl:value-of select="end_date"/></td>
        		</tr>
        	</table>
        	
          <table class="scan_results">
          	<tr>
          		<th>Computer</th>
          		<th>Test type</th>
          		<th>Parameter</th>
          		<th>Result</th>
          	</tr>
          	<xsl:for-each select="target">
          		<xsl:sort select="connection" order="descending" data-type="text"/>
          		<xsl:sort select="computer" order="ascending" data-type="text"/>
          			<tr>
          				<td class="computer_header"><xsl:value-of select="computer"/></td>
          				<td colspan="4" class="computer_header">
          					<xsl:if test="connection = 'Failed'">
          					  <xsl:value-of select="error"/>
          					</xsl:if>
          				</td>
          			</tr>
          			<xsl:if test="count(test) > 0">
            			<xsl:for-each select="test">
            				<tr>
            					<td></td>
            					<td><xsl:value-of select="type"/></td>
            					<td><xsl:value-of select="value"/></td>
            					<td>
            						<xsl:for-each select="result">
            							  <xsl:call-template name="break">
           					  	      <xsl:value-of select="."/>
           					  	    </xsl:call-template>
           					  	    <xsl:if test="position() != last()">
           					  	    	<br/><hr/>
           					  	    </xsl:if>
           					  	</xsl:for-each>
            					</td>
            				</tr>
            			</xsl:for-each>
          			</xsl:if>
          	</xsl:for-each>
          </table>
        </div>
      </xsl:for-each>
    </body>
  </html>

</xsl:template>


<xsl:template name="break">
   <xsl:param name="text" select="."/>
   <xsl:choose>
   <xsl:when test="contains($text, '||')">
      <xsl:value-of select="substring-before($text, '||')"/>
      <br/>
      <xsl:call-template name="break">
          <xsl:with-param name="text" select="substring-after($text,'||')"/>
      </xsl:call-template>
   </xsl:when>
   <xsl:otherwise>
	<xsl:value-of select="$text"/>
   </xsl:otherwise>
   </xsl:choose>
</xsl:template>


</xsl:stylesheet>