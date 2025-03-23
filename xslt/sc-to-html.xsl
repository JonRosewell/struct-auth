<?xml version="1.0" encoding="UTF-8"?>
<xsl:stylesheet xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
    xmlns:xs="http://www.w3.org/2001/XMLSchema"
    xmlns="http://www.w3.org/1999/xhtml" exclude-result-prefixes="" version="3.0">
    
    <!-- Structured Authoring: MS Word to OU Structured Content XML
        MS Word to OU structured content conversion, designed to replace OU IT/LDS 
        customisation for oXygen.
        Jon Rosewell, Jan 2025
        https://github.com/JonRosewell/struct-auth 
        sc-to-html.xsl: converts OU structured content XML into styled HTML for import into Word
    -->
    
    <!-- Namespaces: OU SC xml files have no namespace declaration
        Output in default namespace which is Word html so can be input to html-to-sc for round trip
    -->
    <xsl:output method="html"/>
    <xsl:strip-space elements="*"/>
    
    <xsl:include href="sc-to-html-css.xsl"/>
    
    <!-- TODO:
        ??? More section-like things: FrontMatter, BackMatter, References, 
        Some Word quirks catered for but maybe need checking for consistency, eg avoiding blank lines which could be lost in html stage or Word input (use nbsp or maybe <o:p/> as Word does?) 
        Attributes: piecemeal so far, but maybe just turn all into styled text? Generic may be less code, as well as better round-trip
    -->
    
    <!-- ===== default cases should do the bulk of the work ===== -->
    
    <!-- Copy attributes (not many in SC) where make sense in html. 
        Other attribs are instead copied into text so they can be perserved and edited in Word (see below)
        Slightly risky that same attribute name can be used for diff purposes, eg type for lists and boxes, and
        sometimes needs handling differently, eg for list type, act on it rather than output as attribute span.
        NB need to be careful to apply select="@* | node()" or attribs won't be selected anyway;
        needed even at Paragraph level eg for links.
    -->    
    <xsl:template match="@*" >
        <xsl:copy>
            <xsl:apply-templates select="node()"/>
        </xsl:copy>
    </xsl:template>
    
    <!-- For specific attributes, convert into visible / editable text:
        * src for images
        * id for sections, boxes, activities, figures as destination for CrossRef
        * resource = icons for activities and boxes
        BEWARE: this changes an attribute to child text and could provoke error: 
          An attribute node ([somename]) cannot be created after a child of the containing element.
        which occurs because handling @id has created a child before @somename when 
        processed by apply-templates select="@*".
        Solution adopted is to write surrounding pairs, eg:
            <xsl:apply-templates select="@*[name()!='id']"/>
            <xsl:apply-templates />
            <xsl:apply-templates select="@id"/>
        to ensure other attribs are copied safely before creating content. 
        Looks a bit nicer to have ids at end of para anyway...
    -->
    <xsl:template match="@id | @style | @type | @resource1 | @resource2 | @resource3">
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
        <p class="{name()}">
            <xsl:apply-templates select="@* | node()"/>
        </p>
    </xsl:template>
    
    <!-- Some elements are preserved as html tags -->
    <xsl:template match="b | i | u | sub | sup | table | tbody | tr | td | br | a | font">
        <xsl:element name="{name()}">
            <xsl:apply-templates select="@* | node()"/>
        </xsl:element>
    </xsl:template>
    
    <!-- Specific elements become named spans which then become Word character styles 
        Full set from OU schema Paragraph initially included here, with additional inner 
        elements that are preferable as Word character styles: TeX, MathML,...:
        Specific cases later removed when overriden.
        -->
    <xsl:template match="AuthorComment | EditorComment | ComputerCode | GlossaryTerm | IndexTerm | InlineEquation | InlineFigure | InlineChemistry | Icon | ComputerUI | footnote | language | SecondVoice | SideNote | SideNoteParagraph | Number | Hours | Minutes | TeX | Label">
        <span class="{name()}">
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

    <!-- Conversion of SC lists to html tags should be straightforward, but some tweaks help 
        to round-trip cleanly through Word:
        * if some list items have extended content, then whole list should be wrapped in 
        ListHead/ListEnd to avoid ambiguity over where items end 
        * ...not required for simpler lists where each item is a single paragraph
        * ...limitation: not detecting occasions when item has content *following* a 
        sub list; if so, a SubListHead/SubListEnd around sublist would also be needed.
        * avoid list items with no text by inserting nbsp
    -->
    
    <!-- look for items/subitems that extend to include more than one para, figures, tables etc -->
    <!-- nb not quite perfect since test for children doesn't distinguish tags within 
        para (b, i, sup etc) from tags that are para-like (Paragraph, Figure etc). But too messy to
        exclude long list of things -->
    <xsl:template name="checkForExtendedItems" as="xs:boolean">
        <!-- count lines (ie elements that are not sublists) within each list item of current list -->
        <xsl:variable name="lineCounts" select="ListItem/count(*[not(contains(name(), 'Subsidiary'))])"/>
        <!-- count lines within sublists of current list -->    
        <xsl:variable name="innerLineCounts" select="ListItem/*[contains(name(), 'Subsidiary')]/SubListItem/count(*)"/>
        <!-- has extended content if any count is greater than one -->
        <xsl:value-of select="max(($lineCounts, $innerLineCounts)) gt 1"/>
    </xsl:template>
    
    <xsl:template match="ListItem | SubListItem">
        <li>
            <!-- tweaks to encourage more consistent Word list structure: -->
            <xsl:choose>
                <xsl:when test="text()">    <!-- has raw text, wrap it in Paragraph -->
                    <p><xsl:apply-templates select="@* | node()"/></p>
                </xsl:when>                 <!-- has content children (ignoring any sublists) -->
                <xsl:when test="exists(*[not(contains(name(), 'Subsidiary'))])">
                    <xsl:apply-templates select="@* | node()"/>
                </xsl:when>
                <xsl:otherwise>             <!-- has no content (except maybe sublist), generate nbsp -->
                    <p><xsl:text>&#x00a0;</xsl:text></p>
                    <xsl:apply-templates select="@* | node()"/>
                </xsl:otherwise>
            </xsl:choose>
        </li>
    </xsl:template>
    
    <xsl:template match="NumberedList">
        <xsl:variable name="extended" as="xs:boolean">
            <xsl:call-template name="checkForExtendedItems"/>
        </xsl:variable>
        <xsl:if test="$extended">
            <p class="ListHead"><xsl:text>&#x00a0;[list head]</xsl:text></p>
        </xsl:if>
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
        <xsl:if test="$extended">
            <p class="ListEnd"><xsl:text>&#x00a0;[list end]</xsl:text></p>
        </xsl:if>
    </xsl:template>
    
    <xsl:template match="NumberedSubsidiaryList">
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
    
    
    <xsl:template match="BulletedList">
        <xsl:variable name="extended" as="xs:boolean">
            <xsl:call-template name="checkForExtendedItems"/>
        </xsl:variable>
        <xsl:if test="$extended">
            <p class="ListHead"><xsl:text>&#x00a0;</xsl:text></p>
        </xsl:if>
        <ul>
            <xsl:apply-templates select="@* | node()"/>
        </ul>
        <xsl:if test="$extended">
            <p class="ListEnd"><xsl:text>&#x00a0;</xsl:text></p>
        </xsl:if>
    </xsl:template>
    
    <xsl:template match="BulletedSubsidiaryList">
        <ul>
            <xsl:apply-templates select="@* | node()"/>
        </ul>
    </xsl:template>
    
    <xsl:template match="UnNumberedList">
        <xsl:variable name="extended" as="xs:boolean">
            <xsl:call-template name="checkForExtendedItems"/>
        </xsl:variable>
        <xsl:if test="$extended">
            <p class="ListHead"><xsl:text>&#x00a0;[list head]</xsl:text></p>
        </xsl:if>
        <ul type="none">
            <xsl:apply-templates select="@* | node()"/>
        </ul>
        <xsl:if test="$extended">
            <p class="ListEnd"><xsl:text>&#x00a0;[list end]</xsl:text></p>
        </xsl:if>
    </xsl:template>
    
    <xsl:template match="UnNumberedSubsidiaryList">
        <ul type="none">
            <xsl:apply-templates select="@* | node()"/>
        </ul>
    </xsl:template>
    
    
    <!-- table elements -->
    <!-- close map from Table to html <table> but use initial TableHead para instead of 
        html <caption> since Word seems to convert that into additional table row
        Table/@id has to be placed at end of TableHeader -->
    <xsl:template match="Table">
        <xsl:apply-templates select="TableHead"/> 
        <table>
            <xsl:apply-templates select="@*[name()!='id' and name()!='style']"/>
            <xsl:apply-templates select="tbody"/>
        </table>
        <xsl:apply-templates select="Description"/>
    </xsl:template>
    
    <xsl:template match="TableHead">
        <p class="TableHead">
            <xsl:apply-templates select="node()"/>
            <xsl:apply-templates select="parent::Table/@id"/>
            <xsl:apply-templates select="parent::Table/@style"/>
        </p>
    </xsl:template>

    <!-- tr, th, td, tbody are copied but some attribs should be converted to html equivs -->
    <!-- change td,th alignment to html attribs (with 'decimal' added as fudge) -->
    <xsl:template match="td/@class">
        <xsl:attribute name="align" 
            select="if (. ='TableCentered') then 'center' else
                    if (. ='TableRight') then 'right' else
                    if (. ='TableDecimal') then 'decimal' else 'left'" />
    </xsl:template>
    <xsl:template match="th/@class">
        <xsl:attribute name="align" 
            select="if (. ='ColumnHeadCentered') then 'center' else
                    if (. ='ColumnHeadRight') then 'right' else
                    if (. ='ColumnHeadDecimal') then 'decimal' else 'left'" />
    </xsl:template>
    
    <!-- for borders, need to construct a style string from separate attribs -->
    <xsl:template match="td/@*[contains(name(), 'border')] | th/@*[contains(name(), 'border')]">
        <xsl:attribute name="style">
            <xsl:for-each select="../@*">
                <xsl:if test="name() = 'borderleft' and . = 'true'">
                    <xsl:text>border-left:solid 1pt;</xsl:text>
                </xsl:if>
                <xsl:if test="name() = 'borderright' and . = 'true'">
                    <xsl:text>border-right:solid 1pt;</xsl:text>
                </xsl:if>
                <xsl:if test="name() = 'bordertop' and . = 'true'">
                    <xsl:text>border-top:solid 1pt;</xsl:text>
                </xsl:if>
                <xsl:if test="name() = 'borderbottom' and . = 'true'">
                    <xsl:text>border-bottom:solid 1pt;</xsl:text>
                </xsl:if>
            </xsl:for-each>
        </xsl:attribute>
    </xsl:template>
    
    
    <!-- th will likely be converted to td on pass through Word, so also preserve as a span -->    
    <xsl:template match="th">
        <th>
            <xsl:apply-templates select="@*"/>
            <span class="th">
                <xsl:apply-templates select="node()"/>
            </span>
        </th>
    </xsl:template>
    
    
    <!-- image related -->
    <xsl:template match="Figure">
        <!-- may have @id but attach that to para containing visible image -->
        <xsl:apply-templates />
    </xsl:template>
    
    <xsl:template match="Figure/Image">
        <p class="Figure">
            <img src="{@src}"/>     <!-- image will show in browser or Word (if permission allows) -->
            <xsl:apply-templates select="parent::Figure/@id"/> <!-- tag on id (maybe used by crossref) -->
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
            <xsl:apply-templates />
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
        <xsl:variable name="type" select="name()"/>
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
            <xsl:apply-templates select="@* except(@id | @type | @resource1 | @resource2 | @resource3)"/>
            <xsl:apply-templates select="Heading/node()"/>
            <xsl:if test="not(Heading/node())">
                <xsl:text>&#x00a0;</xsl:text>
            </xsl:if>
            <xsl:apply-templates select="@id | @type | @resource1 | @resource2 | @resource3"/>
        </p>
        <!-- content of box -->
        <xsl:apply-templates/>
        <!-- end marker is nbsp but no other text -->
        <p class="{concat($type, 'End')}">&#x00a0;</p>
    </xsl:template>
    
    <xsl:template match="Box | CaseStudy | Dialogue | Example | Extract | Quote | Reading | StudyNote | Verse | InternalSection | KeyPoints | Activity | Exercise | ITQ | SAQ ">
        <xsl:call-template name="makeBox">
            <xsl:with-param name="type" select="name()"/>
        </xsl:call-template>
    </xsl:template>
    
    <!-- suppress box heading since dealt with above -->
    <xsl:template match="(Box | CaseStudy | Dialogue | Example | Extract | Quote | Reading | StudyNote | Verse | InternalSection | KeyPoints | Activity | Exercise | ITQ | SAQ )/Heading" />
    
    <!-- Dividers that serve to separate parts of container; divider itself has fixed text -->
    <xsl:template match="Interaction | Answer | Discussion">
        <p class="{name()}"><xsl:value-of select="name()"/></p>
        <xsl:apply-templates/>
    </xsl:template>
    <!-- Question part doesn't need divider since question text always follows heading -->
    <xsl:template match="Question" >
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
