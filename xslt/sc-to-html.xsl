<?xml version="1.0" encoding="UTF-8"?>
<xsl:stylesheet xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
    xmlns="http://www.w3.org/1999/xhtml" exclude-result-prefixes="" version="3.0">
    
    <!-- Namespaces: OU SC xml files have no namespace declaration
        Output in default namespace which is Word html so can be input to html-to-sc for round trip
    -->
    <xsl:output method="html"/>
    <xsl:strip-space elements="*"/>
    
    <xsl:include href="sc-to-html-css.xsl"/>
    
    <!-- TODO:
        More box-like things: KeyPoints, MultiPart/Part,...
        ??? More section-like things: FrontMatter, BackMatter, References, 
        Check top level - body is equivalent to Unit, so Item not represented. html-to-sc currently converts body to Item/Unit, making up stuff at Item level, which fits with process of copy/paste into template doc
    -->
    
    <!-- ===== default cases should do the bulk of the work ===== -->
    
    <!-- Copy all attributes (not many in SC) 
        NB need to be careful to apply select="@* | node()" or attribs won't be selected anyway;
        needed even at Paragraph level eg for links.
        Some attribs are also copied into text so they can be perserved and edited in Word:
        - src for images
        - id for sections, boxes, activities as destination for CrossRef
    -->    
    <xsl:template match="@*" >
        <xsl:copy>
            <xsl:apply-templates select="node()"/>
        </xsl:copy>
    </xsl:template>
    
    <!-- BEWARE: this changes an attribute to child text and could provoke error: 
          An attribute node ([somename]) cannot be created after a child of the containing element.
        which occurs because handling @id has created a child before @somename 
        when processed by apply-templates select="@*".
        Solution adopted is to write pairs
            <xsl:apply-templates select="@*[name()!='id']"/>
            <xsl:apply-templates select="@id"/>
        to ensure other attribs are copied safely before creating content. 
        Looks a bit nicer to have ids at end of para anyway...
        Unfortunately ids can occur on many elements larger than para: sections, boxes, tables, figures...
        Generalise template so that can be used for other attributes that need translating to text in future
    -->
    <xsl:template match="@id">
        <span class="attribute">
            <xsl:text>{</xsl:text>
            <xsl:value-of select="name()"/>
            <xsl:text>="</xsl:text>
            <xsl:value-of select="."/>
            <xsl:text>"}</xsl:text>
        </span>
    </xsl:template>
    
    
    <!-- An element becomes <p class="name"> which then becomes Word para style. 
        Elements which become spans override this. -->
    <xsl:template match="*">
        <p class="{local-name()}">
            <xsl:apply-templates select="@* | node()"/>
        </p>
    </xsl:template>
    
    <!-- Some elements are preserved as html tags -->
    <xsl:template match="b | i | u | sub | sup | table | tbody | tr | th | td | br | a | font">
        <xsl:element name="{local-name()}">
            <xsl:apply-templates select="@* | node()"/>
        </xsl:element>
    </xsl:template>
    
    <!-- Specific elements become named spans which then become Word character styles 
        Full set from OU schema Paragraph initially included here, with additional inner 
        elements that are preferable as Word character styles: TeX, MathML,...:
        Specific cases later removed when overriden.
        -->
    <xsl:template match="AuthorComment | EditorComment | ComputerCode | GlossaryTerm | IndexTerm | InlineEquation | InlineFigure | InlineChemistry | Icon | ComputerUI | footnote | language | SecondVoice | SideNote | SideNoteParagraph | Number | Hours | Minutes | TeX">
        <span class="{local-name()}">
            <xsl:apply-templates select="@* | node()"/>
        </span>
    </xsl:template>
    
    
    <!-- ===== simple overrides ===== -->
    
    <!-- basic Paragraph becomes <p> -->
    <xsl:template match="Paragraph">
        <p><xsl:apply-templates select="@* | node()"/></p>
    </xsl:template>
    
    <!-- small caps -->
    <xsl:template match="smallCaps">
        <span style="font-variant:small-caps">
            <xsl:apply-templates select="@* | node()"/>
        </span>
    </xsl:template>
    
    <!-- SideNote/Heading treated as span since occurs within Paragraph (unlike other Headings) -->
    <xsl:template match="SideNote/Heading">
        <span class="SideNoteHeading">
            <xsl:apply-templates select="@* | node()"/>
        </span>
    </xsl:template>
    
    <!-- links: standard hyperlink will look after itself; olink and CrossRef converted 
        to something like markdown: [link text](destination#detail) -->
    <xsl:template match="CrossRef">
        <span class="CrossRef">[<xsl:apply-templates/>](<xsl:value-of select="@idref"/>)</span>
    </xsl:template>
    
    <xsl:template match="olink">
        <span class="olink">[<xsl:apply-templates/>](<xsl:value-of  select="@targetdoc"/>#<xsl:value-of select="@targetptr"/>)</span>
    </xsl:template>
    
    
    <!-- list elements -->
    <xsl:template match="ListItem | SubListItem">
        <li>
            <xsl:apply-templates select="@* | node()"/>
        </li>
    </xsl:template>
    
    <xsl:template match="NumberedList | NumberedSubsidiaryList">
        <ol>
            <xsl:attribute name="type">
                <xsl:choose>
                    <xsl:when test="@class='lower-roman'">i</xsl:when>
                    <xsl:when test="@class='upper-roman'">I</xsl:when>
                    <xsl:when test="@class='lower-alpha'">a</xsl:when>
                    <xsl:when test="@class='upper-alpha'">A</xsl:when>
                    <xsl:otherwise>1</xsl:otherwise>
                </xsl:choose>
            </xsl:attribute>
            <xsl:apply-templates select="@* | node()"/>
        </ol>
    </xsl:template>
    
    
    <xsl:template match="BulletedList | BulletedSubsidiaryList">
        <ul>
            <xsl:apply-templates select="@* | node()"/>
        </ul>
    </xsl:template>
    
    <xsl:template match="UnNumberedList | UnNumberedSubsidiaryList">
        <ul type="none">
            <xsl:apply-templates select="@* | node()"/>
        </ul>
    </xsl:template>
    
    <!-- table elements -->
    <xsl:template match="Table">
        <table>
            <xsl:apply-templates select="@*"/>
            <caption><xsl:apply-templates select="TableHead/node()"/></caption>
            <xsl:apply-templates select="tbody"/>
        </table>
        <xsl:apply-templates select="Description"/>
    </xsl:template>
    <xsl:template match="Table/TableHead"/>
    
    <!-- tr, td, tbody are copied -->
    
    <!-- image related -->
    <xsl:template match="Figure">
<!--        <xsl:apply-templates select="@* | node()"/>-->
        <xsl:apply-templates />
    </xsl:template>
    
    <xsl:template match="Figure/Image">
        <p class="Figure">
            <img src="{@src}"/>     <!-- image tag will show in Word if permission allows -->
        </p>
        <p class="FigureSrc">       <!-- but rely on source as text of separate para -->
            <xsl:value-of select="@src"/>
        </p>
    </xsl:template>
    
    <xsl:template match="InlineFigure/Image">
        <!-- image tag will show in Word if permission allows but rely on source in text -->
        <img src="{@src}"/>     
        [<xsl:value-of select="@alt"/>](<xsl:value-of select="@src"/>)
    </xsl:template>
    
    <xsl:template match="Description | Description/Paragraph">
        <p class="Description">
            <xsl:apply-templates />
        </p>
    </xsl:template>

    <!-- program listing -->
    <xsl:template match="ProgramListing | ProgramListing/Paragraph">
        <p class="ProgramListing">
            <xsl:apply-templates/>
        </p>
    </xsl:template>
    
    <!-- escape for anything too difficult: dump as RawXML -->
    <xsl:template match="MediaContent | InPageActivity | FreeResponse | FreeResponseDisplay | SingleChoice | MultipleChoice | Matching | FreeResponse | VoiceRecorder">
        <p class="RawXML">
            <xsl:value-of select="serialize(.)"/>
        </p>
    </xsl:template>

    <xsl:template match="MathML">
        <span class="MathML">
            <xsl:value-of select="serialize(node())"/>
        </span>
    </xsl:template>
    
    
    
    <!-- ===== overrides that deal with large structures ===== -->

    <!-- Sections are containers that only require an initial para since implicitly closed by later style.
    They have a following Title (except SubSubSection has Heading) which becomes text of para -->
    <xsl:template name="makeSection">
        <xsl:param name="level" select="'h3'"/>
        <xsl:element name="{$level}">
            <xsl:apply-templates select="@*[name()!='id']"/>
            <xsl:if test="$level!='h4'">
                <xsl:apply-templates select="Title/node()"/>
            </xsl:if>
            <xsl:if test="$level='h4'">
                <xsl:apply-templates select="Heading/node()"/>
            </xsl:if>
            <xsl:apply-templates select="@id"/>
        </xsl:element>
        <xsl:apply-templates/>
    </xsl:template>
    
    <xsl:template match="Session | Section | SubSection | SubSubSection">
        <xsl:variable name="type" select="local-name()"/>
        <xsl:call-template name="makeSection">
            <xsl:with-param name="level" select=
                "if ($type='SubSubSection') then 'h4'
                else if ($type='SubSection') then 'h3'
                else if ($type='Section') then 'h2'
                else 'h1'"/>
        </xsl:call-template>
    </xsl:template>
    
    <!-- suppress section titles since dealt with above -->
    <xsl:template match="(Session | Section | SubSection)/Title" /> 
    <xsl:template match="SubSubSection/Heading" /> 
    
    <!-- Box-like structures that require matching xxxHead and closing xxxEnd paras 
        Note: use of xxxHead is consistent with existing LDS conversion 
        https://learn3.open.ac.uk/mod/oucontent/view.php?id=185740&extra=tablelandscape_idm78
    -->
    <xsl:template name="makeBox">
        <xsl:param name="type" select="'Box'"/>
        <!-- head para has text of heading if provided, or nbsp to stop Word losing empty lines -->
        <p class="{concat($type, 'Head')}">
            <xsl:apply-templates select="@*[name()!='id']"/>
            <xsl:apply-templates select="Heading/node()"/>
            <xsl:if test="not(Heading/node())">
                <xsl:text>&#x00a0;</xsl:text>
            </xsl:if>
            <xsl:apply-templates select="@id"/>
        </p>
        <!-- content of box -->
        <xsl:apply-templates/>
        <!-- end marker is nbsp but no other text -->
        <p class="{concat($type, 'End')}">&#x00a0;</p>
    </xsl:template>
    
    <xsl:template match="Box | CaseStudy | Dialogue | Example | Extract | Quote | Reading | StudyNote | Verse | InternalSection | KeyPoints | Activity | Exercise | ITQ | SAQ ">
        <xsl:call-template name="makeBox">
            <xsl:with-param name="type" select="local-name()"/>
        </xsl:call-template>
    </xsl:template>
    
    <!-- suppress box heading since dealt with above -->
    <xsl:template match="(Box | CaseStudy | Dialogue | Example | Extract | Quote | Reading | StudyNote | Verse | InternalSection | KeyPoints | Activity | Exercise | ITQ | SAQ )/Heading" />
    
    <!-- Dividers that serve to separate parts of container; they have no text content of their own -->
    <xsl:template match="Question | Interaction | Answer | Discussion">
        <p class="{local-name()}"><xsl:value-of select="local-name()"/></p>
        <xsl:apply-templates/>
    </xsl:template>
    
    
    
    <!-- ===== document level ===== -->
    
    <!-- document level -->
    <xsl:template match="/">
        <html >
            <head>
                <meta name="generator" content="Jon Rosewell sc-to-html"/>
                <title><xsl:value-of select="Item/ItemTitle"/></title>
                <xsl:call-template name="buildCSS" />
            </head>
            <body>
                <xsl:apply-templates select="/Item/Unit"/>
            </body>
        </html>
    </xsl:template>
    
    <xsl:template match="Unit">
        <xsl:apply-templates/>
    </xsl:template>
    
    <xsl:template match="FrontMatter">
        <xsl:message>FrontMatter: not supported, omitted</xsl:message>
    </xsl:template>
    <xsl:template match="BackMatter">
        <xsl:message>BackMatter: not supported, omitted</xsl:message>
    </xsl:template>
    
    
</xsl:stylesheet>
