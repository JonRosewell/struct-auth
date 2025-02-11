<?xml version="1.0" encoding="UTF-8" ?>
<xsl:transform xmlns:xsl="http://www.w3.org/1999/XSL/Transform" version="2.0">
    <xsl:output method="xml" encoding="UTF-8" />
    
    <!-- strip any existing numbers -->
    <xsl:template match="Number" />
    
    <xsl:template match="Session/Title" >
        <xsl:copy>
            <Number><xsl:number count="Session" level="multiple" format="1"/></Number>
            <xsl:apply-templates />         
        </xsl:copy>
    </xsl:template>
    
    <xsl:template match="Section/Title" >
        <xsl:copy>
            <Number><xsl:number count="Session|Section" level="multiple" format="1.1"/></Number>
            <xsl:apply-templates />         
        </xsl:copy>
    </xsl:template>
    
    <xsl:template match="SubSection/Title" >
        <xsl:copy>
            <Number><xsl:number count="Session|Section|SubSection" level="multiple" format="1.1.1"/></Number>
            <xsl:apply-templates />         
        </xsl:copy>
    </xsl:template>
    
    <xsl:template match="Figure/Caption" >
        <xsl:copy>
            <Number>Figure&#x00a0;<xsl:number count="Figure" level="any" format="1"/></Number>
            <xsl:apply-templates />         
        </xsl:copy>
    </xsl:template>
    
    <xsl:template match="Table/TableHead" >
        <xsl:copy>
            <Number>Table&#x00a0;<xsl:number count="Table" level="any" format="1"/></Number>
            <xsl:apply-templates />         
        </xsl:copy>
    </xsl:template>
    
    <xsl:template match="@*|node()">
        <xsl:copy>
            <xsl:apply-templates select="@*|node()"/>
        </xsl:copy>
    </xsl:template>
    
</xsl:transform>

