<?xml version="1.0" encoding="UTF-8"?>
<xsl:stylesheet version="3.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
    xmlns:h="http://www.w3.org/1999/xhtml" 
    xmlns:o="urn:schemas-microsoft-com:office:office"
    exclude-result-prefixes="h o">
    
    <!-- Structured Authoring: MS Word to OU Structured Content XML
        MS Word to OU structured content conversion, designed to replace OU IT/LDS 
        customisation for oXygen.
        Jon Rosewell, Jan 2025
        https://github.com/JonRosewell/struct-auth 
        html-to-sc.xsl: driver for file conversion of xhtml from Word into OU structured content XML
    -->
    

    <!-- Namespaces: output is OU SC xml file which has no namespace declaration
        Input will be XHTML converted from Word HTML and may also contain MS Office tags
    -->

    <!-- This is driver for Word to OU SC conversion intended for use with entire *files* 
        rather than copy/paste of html fragments from Word into oXygen. All work is done
        by the imported xhtml2sc.xsl code except for the starting template which here 
        overrides to create a skeleton Item/Unit structure. 
        Input must be an XHTML file. This can be created in oXygen using 
        File > Import/Convert > HTML File to XHTML... > XHTML 1.0 Transitional.
    -->
    
    <!-- A few items of metadata must have values for successful render and so default 
        values are provided here. Others can be omitted until publishing. If values are 
        provided in Word they will take precedence.
        See SC Tag Guide for advice https://learn3.open.ac.uk/mod/oucontent/view.php?id=185747 -->
    <xsl:variable name="__CourseCode">X123</xsl:variable>
    <xsl:variable name="__ItemID">X_xxx</xsl:variable>
    <xsl:variable name="__Title">Example</xsl:variable>     <!-- too many titles! assume one only -->
    
    <xsl:import href="xhtml2sc.xsl"/>
    
    
    <!-- kick off by converting html body to Item -->
    <xsl:template mode="styling" match="h:body">
        <xsl:variable name="title" select="//h:p[@class='UnitTitle']"/>
        <xsl:variable name="code" select="//h:p[@class='CourseCode']"/>
        <xsl:variable name="itemID" select="//h:p[@class='ItemID']"/>
        <Item id="{$__ItemID}" TextType="CompleteItem" SchemaVersion="2.0" PageStartNumber="1"
            Template="Generic_A4_Unnumbered" DiscussionAlias="Discussion" SecondColour="None"
            ThirdColour="None" FourthColour="None" Logo="colour" Rendering="VLE2 modules (learn2)">
            <meta content="True" name="VideoPosterFrame"/>
            <CourseCode>
                <xsl:value-of select="if ($code!='') then $code else $__CourseCode"/>
            </CourseCode>
            <CourseTitle>
                <xsl:value-of select="//h:p[@class='CourseTitle']"/>
            </CourseTitle>
            <ItemID>
                <xsl:value-of select="if ($title!='') then $title else $__Title"/>
            </ItemID>
            <ItemTitle>
                <xsl:value-of select="if ($title!='') then $title else $__Title"/>
            </ItemTitle>
            <Unit>
                <UnitID>
                    <xsl:value-of select="//h:p[@class='UnitID']"/>
                </UnitID>
                <UnitTitle>
                    <xsl:value-of select="if ($title!='') then $title else $__Title"/>
                </UnitTitle>
                <ByLine>
                    <xsl:value-of select="//h:p[@class='ByLine']"/>
                </ByLine>
                <xsl:apply-templates mode="styling"/>
            </Unit>
        </Item>
    </xsl:template>
    <!-- metadata already dealt with by Item so suppress -->
    <xsl:template mode="styling" match="h:p[@class='CourseCode']"/>
    <xsl:template mode="styling" match="h:p[@class='CourseTitle']"/>
    <xsl:template mode="styling" match="h:p[@class='ItemID']"/>
    <xsl:template mode="styling" match="h:p[@class='ItemTitle']"/>
    <xsl:template mode="styling" match="h:p[@class='UnitID']"/>
    <xsl:template mode="styling" match="h:p[@class='UnitTitle']"/>
    <xsl:template mode="styling" match="h:p[@class='ByLine']"/>
    

    <!-- override starting nodes for structural passes: boxing, sectioning
        start process of building sections at Item/Unit, not root 
    -->
    <xsl:template mode="boxing" match="/Item/Unit">
        <xsl:call-template name="buildBoxes"/>
    </xsl:template>
    
    <xsl:template mode="sectioning" match="/Item/Unit">
        <Unit>
            <xsl:call-template name="buildUnit"/>
        </Unit>
    </xsl:template>
    

</xsl:stylesheet>