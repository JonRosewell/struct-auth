<?xml version="1.0" encoding="UTF-8" ?>
<xsl:transform xmlns:xsl="http://www.w3.org/1999/XSL/Transform" version="2.0">
    <xsl:output method="xml" encoding="UTF-8" />
    
    <!-- Structured Authoring: MS Word to OU Structured Content XML
        MS Word to OU structured content conversion, designed to replace OU IT/LDS 
        customisation for oXygen.
        Jon Rosewell, Jan 2025
        https://github.com/JonRosewell/struct-auth 
        number-sc.xsl: autonumbers sections, figures etc in OU structured content XML
    -->
    
    <!-- strip any existing Numbers -->
    <xsl:template match="Number" />

    <!-- don't strip any existing Labels (eg typically only some equations are numbered -->
    
    <!-- following templates all replace existing numbers with value from appropriate sequence -->
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
    
    <xsl:template match="Activity/Heading" >
        <xsl:copy>
            <Number>Activity&#x00a0;<xsl:number count="Activity" level="any" format="1"/></Number>
            <xsl:apply-templates />         
        </xsl:copy>
    </xsl:template>
    
    <xsl:template match="Exercise/Heading" >
        <xsl:copy>
            <Number>Exercise&#x00a0;<xsl:number count="Exercise" level="any" format="1"/></Number>
            <xsl:apply-templates />         
        </xsl:copy>
    </xsl:template>
    
    <xsl:template match="SAQ/Heading" >
        <xsl:copy>
            <Number>SAQ&#x00a0;<xsl:number count="SAQ" level="any" format="1"/></Number>
            <xsl:apply-templates />         
        </xsl:copy>
    </xsl:template>
    
    <xsl:template match="ITQ/Heading" >
        <xsl:copy>
            <Number>ITQ&#x00a0;<xsl:number count="ITQ" level="any" format="1"/></Number>
            <xsl:apply-templates />         
        </xsl:copy>
    </xsl:template>
    
    <xsl:template match="Label" >
        <xsl:copy>(<xsl:number count="Label" level="any" format="1"/>)</xsl:copy>
    </xsl:template>
    
    <!-- default copy template -->    
    <xsl:template match="@*|node()">
        <xsl:copy>
            <xsl:apply-templates select="@*|node()"/>
        </xsl:copy>
    </xsl:template>
    
</xsl:transform>

