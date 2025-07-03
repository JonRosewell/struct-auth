<?xml version="1.0" encoding="UTF-8" ?>
<xsl:transform xmlns:xsl="http://www.w3.org/1999/XSL/Transform" version="3.0">
    <xsl:output method="xml" encoding="UTF-8" />
    
    <!-- Structured Authoring: MS Word to OU Structured Content XML
        MS Word to OU structured content conversion, designed to replace OU IT/LDS 
        customisation for oXygen.
        Jon Rosewell, Jan 2025
        https://github.com/JonRosewell/struct-auth 
        crossref-sc.xsl: updates references to number/label elements in OU structured content XML
    -->
    
    <xsl:template match="CrossRef">
        <xsl:copy>
            <xsl:apply-templates select="@idref"/>      <!-- preserve idref -->
            <xsl:apply-templates select="//*[@id=current()/@idref]" mode="xref"/>
        </xsl:copy>
    </xsl:template>
    
    <xsl:template mode="xref" match="Session | Section | SubSection">
        <xsl:value-of select="name()"/><xsl:text>&#x00a0;</xsl:text>
        <xsl:apply-templates mode="xref" select="Title/Number"/>  <!-- or "Title" for full text -->
    </xsl:template>

    <xsl:template mode="xref" match="Figure">
        <xsl:apply-templates mode="xref" select="Caption/Number" />
    </xsl:template>
    
    <xsl:template mode="xref" match="Activity | Box">
        <xsl:apply-templates mode="xref" select="Heading/Number" />
    </xsl:template>
    
    <xsl:template mode="xref" match="Equation">
        <xsl:apply-templates mode="xref" select="Label" />
    </xsl:template>
    
    <xsl:template mode="xref" match="Number | Label">
        <xsl:apply-templates />
    </xsl:template>            
    

    <!-- default copy template -->    
    <xsl:template match="@*|node()">
        <xsl:copy>
            <xsl:apply-templates select="@*|node()"/>
        </xsl:copy>
    </xsl:template>
    
</xsl:transform>

