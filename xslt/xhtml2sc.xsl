<?xml version="1.0" encoding="UTF-8"?>
<xsl:stylesheet 
    xmlns:xsl="http://www.w3.org/1999/XSL/Transform" version="3.0" 
    xmlns:h="http://www.w3.org/1999/xhtml" 
    xmlns:o="urn:schemas-microsoft-com:office:office" 
    exclude-result-prefixes="h o"
    >
    
    <!-- Namespaces: output is OU SC xml file which has no namespace declaration
        Input will be XHTML converted from Word HTML and may also contain MS Office tags
    -->

    <xsl:strip-space elements="*"/>
    <xsl:output method="xml" version="1.0" indent="no"/>

    <!-- Overall organisation into several passes (using modes) allows each to solve a single 
        problem, working from smallest scale to largest:
        - identity      for testing, copies everything unchanged
        - styling       translate html into corresponding SC elements
        - attributing   move attribute spans into true attributes
        - objecting     fixed-format structures/objects such as Figure, olink, CrossRef, Table...
        - listing       create lists from linear sequences of ListItem, or from ul/ol if they exist
        - boxing        create boxes between start/end styles for Box, Example, Reading, InternalSection...
        - questioning   create question/answer/discussion for Activity, Exercise, SAQ,...
        - sectioning    split into Session/Section/Subsections
        - fixing        last chance to fix problems...
    -->

    <!-- starting point of transform -->
    <xsl:template match="/" >
        <!-- to show input: -->
<!--        <xsl:variable name="identity-result">
            <xsl:apply-templates mode="identity" select="/"/>
        </xsl:variable>
-->        
        <xsl:variable name="styling-result">
            <xsl:apply-templates mode="styling" select="h:html/h:body"/>
        </xsl:variable>
        <xsl:variable name="attributing-result">
            <xsl:apply-templates mode="attributing" select="$styling-result"/>
        </xsl:variable>
        <xsl:variable name="objecting-result">
            <xsl:apply-templates mode="objecting" select="$attributing-result"/>
        </xsl:variable>
        <xsl:variable name="listing-result">
            <xsl:apply-templates mode="listing" select="$objecting-result"/>
        </xsl:variable>
        <xsl:variable name="boxing-result">
            <xsl:apply-templates mode="boxing" select="$listing-result"/>
        </xsl:variable>
        <xsl:variable name="questioning-result">
            <xsl:apply-templates mode="questioning" select="$boxing-result"/>
        </xsl:variable>
        <xsl:variable name="sectioning-result">
            <xsl:apply-templates mode="sectioning" select="$questioning-result"/>
        </xsl:variable>
        <xsl:variable name="fixing-result">
            <xsl:apply-templates mode="fixing" select="$sectioning-result"/>
        </xsl:variable>
        
        <xsl:copy-of select="$fixing-result"/>

    </xsl:template>

    <!-- TODO:
        Think about supporting more attributes? Table styles? Activity icons?
        Multipart activities not supported: worth doing or not?
    -->

    <!-- a placeholder image, currently sitting in TM129 Sharepoint -->
    <xsl:variable name="missing-image" select="'https://openuniv.sharepoint.com/sites/tmodules/tm129/lmimages/missing.png'"/>
    

    <!-- Identity ========================================================== -->
    
    <!-- identity transform for testing: copy elements and attributes unchanged -->
    <xsl:template mode="identity" match="* | @* ">
        <xsl:copy>
            <xsl:apply-templates mode="identity" select="@* | node()"/>
        </xsl:copy>
    </xsl:template>

    <xsl:template mode="identity" match="/" >
        <xsl:apply-templates mode="identity"/>
    </xsl:template>
    

    <!-- Styling =========================================================== -->
    <!-- Styling pass does 1:1 renaming of elements and cleans some stuff:
        * turns <p> with named class (=Word para style) into element
        * turns unclassed <p> or MsoNormal into Paragraph
        * turns span with named class (=Word char style) into element 
        * strips as much as possible of html fluff. 
        * strips some elements: div, unnamed spans, spell/grammar errors
        * strips most attributes but preserves a few: a/@href, ol/@start...
        Result will be much simpler, no longer in html namespace, but still flat linear.
    -->

    <!-- Default: para with class (= Word para style) or span with class (= Word char style)
        become an element of same name. 
        This is lower than normal priority (would be 0.5) to ensure special cases will override -->
    <xsl:template mode="styling" match="h:p[@class] | h:span[@class]" priority="0.4">
        <xsl:element name="{@class}">
            <xsl:apply-templates mode="styling" />
        </xsl:element>
    </xsl:template>

    <!-- para with no class or with class=MsoNormal becomes Paragraph -->
    <xsl:template mode="styling" match="h:p[not(@class) or @class='MsoNormal']">
        <Paragraph>
            <xsl:apply-templates mode="styling" />
        </Paragraph>
    </xsl:template>
    
    <!-- Copy remaining elements (stripping namespace). A few specific attribs preserved where needed -->
    <xsl:template mode="styling" match="*">
        <xsl:element name="{local-name()}">
            <xsl:apply-templates mode="styling" select="node() | @href | @start | @type"/>
        </xsl:element>
    </xsl:template>
    
    <!-- table cells: preserve spans, and infer alignment using heuristics -->
    <xsl:template mode="styling" match="h:td">
        <xsl:element name="{local-name()}">
            <!-- colspan/rowspan: will get default value=1 (from schema?) unless ignored -->
            <xsl:apply-templates mode="styling" select="@colspan[. != 1] | @rowspan[. != 1]"/>
            <!-- get cell alignment either from td or from child para -->
            <xsl:if test="@align='center' or descendant::h:p[@align='center']">
                <xsl:attribute name="class" select="'TableCentered'"/>
            </xsl:if>
            <xsl:if test="@align='right' or descendant::h:p[@align='right']">
                <xsl:attribute name="class" select="'TableRight'"/>
            </xsl:if>
            <!-- get decimal from alignment from child para decimal tab stop; also allow
                fudged td style from sc-to-html down conversion -->
            <xsl:if test="@align='decimal' or descendant::h:p[contains(@style, 'decimal')]">
                <xsl:attribute name="class" select="'TableDecimal'"/>
            </xsl:if>
            <!-- content: -->
            <xsl:apply-templates mode="styling" select="node()"/>
        </xsl:element>
    </xsl:template>

    <!-- copy attributes where needed -->
    <xsl:template mode="styling" match="@*">
        <xsl:copy>
            <xsl:apply-templates mode="styling" select="node()"/>
        </xsl:copy>
    </xsl:template>

    <!-- simple name changes -->
    <xsl:template mode="styling" match="h:span[contains(@style, 'small-caps')]">
        <smallCaps>
            <xsl:apply-templates mode="styling"/>
        </smallCaps>
    </xsl:template>

    <xsl:template mode="styling" match="h:span[@class = 'SideNoteHeading']">
        <Heading>
            <xsl:apply-templates mode="styling"/>
        </Heading>
    </xsl:template>
    
    <xsl:template mode="styling" match="h:p[@class = 'MsoCaption']">
        <Caption>
            <xsl:apply-templates mode="styling"/>
        </Caption>
    </xsl:template>
    
    <xsl:template mode="styling" match="h:strong">
        <b>
            <xsl:apply-templates mode="styling"/>
        </b>
    </xsl:template>

    <xsl:template mode="styling" match="h:em">
        <i>
            <xsl:apply-templates mode="styling"/>
        </i>
    </xsl:template>

    <!-- strip office namespace tags; seem to be used to prevent empty html paras but have no content -->
    <xsl:template mode="styling" match="o:*">
        <xsl:apply-templates mode="styling"/>
    </xsl:template>

    <!-- strip indications for spelling and grammar errors -->
    <xsl:template mode="styling" match="h:span[@class='SpellE' or @class='GramE']">
        <xsl:apply-templates mode="styling"/>
    </xsl:template>

    <!-- strip anchor tags, preserving content -->
    <xsl:template mode="styling" match="h:a[@name != '']">
        <xsl:apply-templates mode="styling"/>
    </xsl:template>
    
    <!-- strip other (unnamed) spans, preserving content -->
    <xsl:template mode="styling" match="h:span">
        <xsl:apply-templates mode="styling"/>
    </xsl:template>

    <!-- strip divs, preserving content -->
    <xsl:template mode="styling" match="h:div">
        <xsl:apply-templates mode="styling"/>
    </xsl:template>

    <!-- deal with Word comments and track changes.  
        Word may use html ins and del elements for insertions/deletions: keep ins, lose del contents.
        Comments are turned into links to endnotes. Links have distinctive style; endnotes are in a styled div.
    -->
    <!-- keep insertions, losing tag -->
    <xsl:template mode="styling" match="h:ins">
        <xsl:apply-templates mode="styling"/>
    </xsl:template>
    <!-- lose deleted text and tag -->
    <xsl:template mode="styling" match="h:del"/>
    
    <!-- strip anchor on commented text, preserving content -->
    <xsl:template mode="styling" match="h:a[contains(@style, 'mso-comment-reference')]">
        <xsl:apply-templates mode="styling"/>
    </xsl:template>
    <!-- strip link to comment, deleting link text -->
    <xsl:template mode="styling" match="h:span[@class='MsoCommentReference']" />
    <!-- strip anchor of back link from comment, deleting the inserted text -->
    <xsl:template mode="styling" match="h:a[@class='msocomanchor']" />
    
    
    <!-- strip div used for Word comments, including content -->
    <xsl:template mode="styling" match="h:div[contains(@style, 'mso-element:comment-list')]" />

    

    <!-- lists: heuristic to determine list type from MSO fallback text; return level, type & start -->
    <xsl:template name="getListInfo">
        <xsl:param name="leader"/>
        <xsl:param name="style"/>
        <!-- clue is first char after stripping leading white space *and nbsp* -->
        <xsl:variable name="clue" select="substring(normalize-space(translate($leader, '&#xA0;', '')), 1 ,1)"/>
        <!-- level & margin may be in style; NB if missing take value NaN -->
        <xsl:variable name="_level" select="number(substring-before(substring-after($style, 'level'), ' '))"/>
        <xsl:variable name="_margin" select="number(substring-before(substring-after($style, 'margin-left:'), 'pt'))"/>
        <!-- use level if explicit, otherwise infer from margin if given, else assume 1 -->
        <xsl:variable name="level" select="if ($_level &gt; 1) then 2 
                                            else if ($_level=1) then 1 
                                            else if ($_margin>50) then 2 else 1"/>
        <xsl:choose>
            <xsl:when test="$clue = ''">
                <__listInfo level="{$level}" type="unnumbered" start="0"/>
            </xsl:when>
            <xsl:when test="contains('1234567890', $clue)">
                <__listInfo level="{$level}" type="decimal" 
                    start="{string-length(substring-before('1234567890', $clue))+1}"/>
            </xsl:when>
            <xsl:when test="contains('ivx', $clue)"> <!-- cop out roman-to-decimal conversion! -->
                <__listInfo level="{$level}" type="lower-roman" 
                    start="{if ($clue='x') then 10 else if ($clue='v') then 5 else 1}"/> 
            </xsl:when>
            <xsl:when test="contains('IVX', $clue)">
                <__listInfo level="{$level}" type="upper-roman" 
                    start="{if ($clue='X') then 10 else if ($clue='V') then 5 else 1}"/>
            </xsl:when>
            <xsl:when test="contains('abcdefghijklmn_pqrstuvwxyz', $clue)"> <!-- 'o' may be bullet! -->
                <__listInfo level="{$level}" type="lower-alpha" 
                    start="{string-length(substring-before('abcdefghijklmnopqrstuvwxyz', $clue))+1}"/>
            </xsl:when>
            <xsl:when test="contains('ABCDEFGHIJKLMN_PQRSTUVWXYZ', $clue)">
                <__listInfo level="{$level}" type="upper-alpha" 
                    start="{string-length(substring-before('ABCDEFGHIJKLMNOPQRSTUVWXYZ', $clue))+1}"/>
            </xsl:when>
            <xsl:otherwise>
                <__listInfo level="{$level}" type="bulleted" start="0"/>
            </xsl:otherwise>
        </xsl:choose>
    </xsl:template>
    
    
    <!-- List items. Word may preserve ol/ul/li styling but typically now creates styled <p>. 
        Therefore must recognise a list item from clues in class or style; template is higher
        priority to ensure these paras picked out first. Level of list can be obtained 
        directly from style, but list type and start have to be inferred from a span of 
        fallback leader text provided by Word for browser display (since not using li tags).
        At this stage, convert list items into ListItem or SubListItem, and enclose with 
        BulletedList/NumberedList in later pass. That means @listType and @start added 
        to ListItem/SubListItem for later use after which will be stripped out again.
        NB level clues assume lists are styled carefully in Word: probably ok for single 
        type list or outline type, but changing type of sublist will reset level. 
    -->
    <xsl:template mode="styling" match="h:p[contains(@style, 'mso-list:') or contains(@class, 'MsoList')]" priority="0.6">
        <xsl:variable name="leader" select=".//h:span[contains(@style, 'mso-list:Ignore')]"/>
        <xsl:variable name="info">
            <xsl:call-template name="getListInfo">
                <xsl:with-param name="leader" select="$leader"/>
                <xsl:with-param name="style" select="@style"/>
            </xsl:call-template>
        </xsl:variable>
        <xsl:variable name="level" select="$info/__listInfo/@level"/>
        <xsl:variable name="type"  select="$info/__listInfo/@type"/>
        <xsl:variable name="start" select="$info/__listInfo/@start"/>
        
        <xsl:element name="{if ($level &gt;= 2) then 'SubListItem' else 'ListItem'}">
            <xsl:attribute name="listType" select="$type"/>
            <xsl:attribute name="start" select="$start"/>
            <Paragraph>
                <xsl:apply-templates mode="styling"/>
            </Paragraph>
        </xsl:element>        
    </xsl:template>
    
    <!-- strip out leader text which is used for fallback numbering of lists (eg literal 1., a. etc) -->
    <xsl:template mode="styling" match="h:span[contains(@style, 'mso-list:Ignore')]"/>
    
    
    <!-- start at html body -->
    <xsl:template mode="styling" match="h:body">
        <xsl:apply-templates mode="styling"/>
    </xsl:template>
    

    
    <!-- Attributing ======================================================== -->
    
    <!-- Attributes (currently only id) may occur as spans in element text but need 
        hoisting into element as attribute. For Title & Heading id needs to be in 
        containing structure rather than the heading.
        Assume a syntax like {id="s4.a"} and only a single attribute name/value pair. 
        If later necessary, could split several attrib pairs by separator.
    -->
    
    <xsl:template mode="attributing" match="attribute">
        <xsl:variable name="pair" select="substring-before(substring-after( . , '{'), '}')"/>
        <xsl:variable name="attName" select="normalize-space(substring-before($pair, '='))"/>
        <xsl:variable name="attValue" select="substring-before(substring-after($pair, '=&quot;'), '&quot;')"/>
        <xsl:attribute name="{$attName}" select="$attValue"/>
    </xsl:template>    
    
    <!-- default identity: copy attributes unchanged -->
    <xsl:template mode="attributing" match="@*">
        <xsl:copy-of select="."/>
    </xsl:template>
    
    <!-- default identity: copy elements unchanged, hoisting 'attribute' element into true attribute -->
    <xsl:template mode="attributing" match="*">
        <xsl:copy>
            <xsl:apply-templates mode="attributing" select="@*"/>
            <xsl:apply-templates mode="attributing" select="attribute"/>
            <xsl:apply-templates mode="attributing" select="node()[not(self::attribute)]"/>
        </xsl:copy>
    </xsl:template>
    
    
    <!-- Objecting ========================================================= -->
    
    <!-- default identity: copy elements and attributes unchanged -->
    <xsl:template mode="objecting" match="* | @* " >
        <xsl:copy>
            <xsl:apply-templates mode="objecting" select="@* | node()"/>
        </xsl:copy>
    </xsl:template>
    
    <xsl:template mode="objecting" match="RawXML">
        <xsl:copy-of select="parse-xml(.)"/>
    </xsl:template>
    
    <!-- equations: can be either display or inline, TeX or MathML. 
        Limitation: once in Word, can't use multiple named char styles for same span of text. 
        So can't preserve both <InlineEquation><TeX> and <InlineEquation><MathML>, although 
        can preserve both <Equation> (=para) and <Tex> or <MathML> (=char). 
        So preserve MathML where explicit, assume TeX otherwise; preserve Equation where explict 
        and assume inline otherwise. Default tagging will do most of work, just need to deal with 
        overrides.
        Treat TeX as plain text, MathML as XML.
    -->
    <xsl:template mode="objecting" match="InlineEquation[not(child::TeX) and not(child::MathML)]">
        <InlineEquation>
            <TeX>
                <xsl:value-of select="." />
            </TeX>
        </InlineEquation>
    </xsl:template>
    
    <xsl:template mode="objecting" match="TeX[not(parent::InlineEquation) and not(parent::Equation)]">
        <InlineEquation>
            <TeX>
                <xsl:value-of select="." />
            </TeX>
        </InlineEquation>
    </xsl:template>
    
    <xsl:template mode="objecting" match="MathML[not(parent::InlineEquation) and not(parent::Equation)]">
        <InlineEquation>
            <MathML>
                <xsl:copy-of select="parse-xml(.)"/>
            </MathML>
        </InlineEquation>
    </xsl:template>
    
    <xsl:template mode="objecting" match="MathML">
        <MathML>
            <xsl:copy-of select="parse-xml(.)"/>
        </MathML>
    </xsl:template>
    
    
    <xsl:template mode="objecting" match="ProgramListing">
        <xsl:copy>
            <Paragraph>
                <xsl:apply-templates mode="objecting" select="@* | node()"/>
            </Paragraph>
        </xsl:copy>
    </xsl:template>
    
    <!-- olink, CrossRef and InlineImage use markdown-like syntax for links -->
    <xsl:template mode="objecting" match="olink" >
        <xsl:copy>
            <xsl:variable name="text" select="." />
            <xsl:attribute name="targetdoc" select="substring-before(substring-after($text, '('), '#')"/>
            <xsl:attribute name="targetptr" select="substring-before(substring-after($text, '#'), ')')"/>
            <xsl:value-of select="substring-before(substring-after($text, '['), ']')"/>
        </xsl:copy>
    </xsl:template>
    
    <xsl:template mode="objecting" match="CrossRef">
        <xsl:copy>
            <xsl:variable name="text" select="." />
            <xsl:attribute name="idref" select="substring-before(substring-after($text, ']('), ')')"/>
            <xsl:value-of select="substring-before(substring-after($text, '['), ']')"/>
        </xsl:copy>
    </xsl:template>
    
    <xsl:template mode="objecting" match="InlineFigure">
        <xsl:copy>
            <Image>
                <xsl:variable name="text" select="." />
                <xsl:attribute name="src" select="substring-before(substring-after($text, '('), ')')"/>
                <xsl:attribute name="alt" select="substring-before(substring-after($text, '['), ']')"/>
            </Image>
        </xsl:copy>
    </xsl:template>
    
    <!-- tables -->
    <xsl:template mode="objecting" match="table">
        <Table>
            <!-- if previous TableHead has an id, apply to Table -->
            <xsl:variable name="th_id" select="preceding-sibling::*[1][self::TableHead]/@id"/>
            <xsl:if test="$th_id != ''">
                <xsl:attribute name="id" select="$th_id"/>
            </xsl:if>
            <TableHead>
                <!-- take content from preceding TableHead and/or table/caption -->
                <xsl:apply-templates mode="objecting" select="preceding-sibling::*[1][self::TableHead]/node()"/>
                <xsl:apply-templates mode="objecting" select="caption/node()"/>
            </TableHead>
            <xsl:apply-templates mode="objecting"/>
            <Description>
                <xsl:apply-templates mode="force" select="following-sibling::*[1][self::Description]" />
            </Description>
        </Table>
    </xsl:template>
    
    <!-- TableHead and caption: suppress content since dealt with as part of Table -->
    <xsl:template mode="objecting" match="TableHead" />
    <xsl:template mode="objecting" match="table/caption" />
    

    <!-- multicolumn text: treat as special case of table -->
    <xsl:template mode="objecting" match="MultiColumnBody/table">
        <MultiColumnText>
            <MultiColumnHead>
                <!-- if there is a preceding MultiColumnHead, process its content -->
                <xsl:apply-templates mode="objecting" select="../preceding-sibling::*[1][self::MultiColumnHead]/node()"/>
            </MultiColumnHead>
            <MultiColumnBody>
                <Table> <!-- no caption/heading or description -->
                    <xsl:apply-templates mode="objecting"/>
                </Table>
            </MultiColumnBody>
        </MultiColumnText>
    </xsl:template>

    <!-- MultiColumnText, ...Head, ...Body: look for embedded table and treat as special case of table -->
    <xsl:template mode="objecting" match="MultiColumnText" >
        <xsl:apply-templates mode="objecting"/>
    </xsl:template>
    <xsl:template mode="objecting" match="MultiColumnHead" /> <!-- suppress -->
    <xsl:template mode="objecting" match="MultiColumnBody" >
        <xsl:apply-templates mode="objecting" select="table"/>
    </xsl:template>
    
    
    <!-- Figure is required, other related tags are optional -->
    <xsl:template mode="objecting" match="Figure">
        <xsl:variable name="figSrc" select="following-sibling::*[1][self::FigureSrc]"/>
        <Figure>
            <Image src="{if ($figSrc!='') then $figSrc else $missing-image}" />
            <Caption>
                <xsl:apply-templates mode="force" select="following-sibling::*[position() &lt; 3][self::Caption]" />
            </Caption>
            <Alternative>
                <xsl:apply-templates mode="force" select="following-sibling::*[position() &lt; 4][self::Alternative]" />
            </Alternative>
            <SourceReference>
                <xsl:apply-templates mode="force" select="following-sibling::*[position() &lt; 5][self::SourceReference]" />
            </SourceReference>
            <Description>
                <xsl:apply-templates mode="force" select="following-sibling::*[position() &lt; 6][self::Description][1]" />
            </Description>
        </Figure>
    </xsl:template>
    
    <!-- force templates for out of order processing -->
    <xsl:template mode="force" match="Description">
        <Paragraph>
            <xsl:apply-templates mode="objecting" />
        </Paragraph>
        <!-- if there is a following Description para, tail recurse to it -->
        <xsl:apply-templates mode="force" select="following-sibling::*[1][self::Description]" />
    </xsl:template>
    
    <xsl:template mode="force" match="SourceReference | Caption | Alternative">
        <xsl:apply-templates mode="objecting" />
    </xsl:template>
    
    <!-- Suppress these in normal mode since dealt with by Figure -->
    <xsl:template mode="objecting" match="FigureSrc | Caption | Alternative | SourceReference | Description | img" />


    <!-- Listing =========================================================== -->

    <!-- default identity: copy elements and attributes unchanged -->
    <xsl:template mode="listing" match="* | @* ">
        <xsl:copy>
            <xsl:apply-templates mode="listing" select="@* | node()"/>
        </xsl:copy>
    </xsl:template>

    <!-- depending on origin, possible to find ul/ol/li tags which translate readily to SC, although
        need to be aware of permutations for nesting.
        Unfortunately Word now seems to generate flat lists of styled <p>. These have been converted to 
        ListItem/SubList item but need an enclosing list element. -->
    <xsl:template mode="listing" match="ul">
        <xsl:choose>
            <xsl:when test="@type='none'">
                <UnNumberedList>
                    <xsl:apply-templates mode="listing"/>
                </UnNumberedList>
            </xsl:when>
            <xsl:otherwise>
                <BulletedList>
                    <xsl:apply-templates mode="listing"/>
                </BulletedList>
            </xsl:otherwise>
        </xsl:choose>
    </xsl:template>
    
    <!-- sub lists should be in li of parent list but may occur directly within parent list -->
    <xsl:template mode="listing" match="li/ul | ul/ul | ol/ul">
        <xsl:choose>
            <xsl:when test="@type='none'">
                <UnNumberedSubsidiaryList>
                    <xsl:apply-templates mode="listing"/>
                </UnNumberedSubsidiaryList>
            </xsl:when>
            <xsl:otherwise>
                <BulletedSubsidiaryList>
                    <xsl:apply-templates mode="listing"/>
                </BulletedSubsidiaryList>
            </xsl:otherwise>
        </xsl:choose>
    </xsl:template>
    
    <xsl:template mode="listing" match="ol">
        <NumberedList>
            <xsl:if test="@start">
                <xsl:attribute name="start" select="@start"/>
            </xsl:if>
            <xsl:if test="@type">
                <xsl:attribute name="class" select="if (@type='i') then 'lower-roman' else if (@type='I') then 'upper-roman' else if (@type='a') then 'lower-alpha' else if (@type='A') then 'upper-alpha' else 'decimal'"/>
            </xsl:if>
            <xsl:apply-templates mode="listing"/>
        </NumberedList>
    </xsl:template>
    
    <xsl:template mode="listing" match="li/ol | ul/ol | ol/ol">
        <NumberedSubsidiaryList>
            <xsl:if test="@start">
                <xsl:attribute name="start" select="@start"/>
            </xsl:if>
            <xsl:if test="@type">
                <xsl:attribute name="class" select="if (@type='i') then 'lower-roman' else if (@type='I') then 'upper-roman' else if (@type='a') then 'lower-alpha' else if (@type='A') then 'upper-alpha' else 'decimal'"/>
            </xsl:if>
            <xsl:apply-templates mode="listing"/>
        </NumberedSubsidiaryList>
    </xsl:template>
    
    <xsl:template mode="listing" match="li">
        <ListItem>
            <xsl:apply-templates mode="listing"/>
        </ListItem>
    </xsl:template>

    <xsl:template mode="listing" match="li/ul/li | li/ol/li | ul/ul/li | ul/ol/li | ol/ul/li | ol/ol/li">
        <SubListItem>
            <xsl:apply-templates mode="listing"/>
        </SubListItem>
    </xsl:template>
    
    
    <!-- Code below will create list structure from runs of ListItem/SubListItem. 
        Has issue that sublists are produced between items rather than as part of preceding item.
        Determines the class (numeric/alpha/roman) and start number from ListItem, set in earlier pass. 
    -->
    <xsl:template mode="listing" match="/">
        <xsl:copy>
            <xsl:apply-templates mode="listing" select="@*"/>
            <xsl:for-each-group select="*" group-adjacent="name()='ListItem' or name()='SubListItem'">
                <xsl:choose>
                    <xsl:when test="current-grouping-key()">
                        <!-- check first ListItem for class of list and start -->
                        <xsl:variable name="class" select="current-group()[1]/@listType"/>
                        <xsl:variable name="start" select="current-group()[1]/@start"/>
                        <!-- make appropriate element to contain the ListItems -->
                        <xsl:element name="{if ($class='unnumbered') then 'UnNumberedList' else if ($class='bulleted') then 'BulletedList' else 'NumberedList'}">
                            <xsl:if test="($class != 'bulleted') and ($class != 'unnumbered')">
                                <xsl:attribute name="class" select="$class"/>
                                <xsl:attribute name="start" select="$start"/>
                            </xsl:if>
                            <xsl:for-each-group select="current-group()" group-adjacent="name()='SubListItem'">
                                <xsl:choose>
                                    <xsl:when test="current-grouping-key()">
                                        <!-- check first SubListItem for class and start -->
                                        <xsl:variable name="class" select="current-group()[1]/@listType"/>
                                        <xsl:variable name="start" select="current-group()[1]/@start"/>
                                        <!-- make appropriate element to contain the SubListItems -->
                                        <xsl:element name="{if ($class='bulleted') then 'BulletedSubsidiaryList' else if ($class='unnumbered') then 'UnNumberedSubsidiaryList' else 'NumberedSubsidiaryList'}">
                                            <xsl:if test="($class != 'bulleted') and ($class != 'unnumbered')">
                                                <xsl:attribute name="class" select="$class"/>
                                                <xsl:attribute name="start" select="$start"/>
                                            </xsl:if>
                                            <xsl:apply-templates mode="listing" select="current-group()"/>
                                        </xsl:element>
                                    </xsl:when>
                                    <xsl:otherwise> <!-- not a sublist, a run of ListItems -->
                                        <xsl:apply-templates mode="listing" select="current-group()"/>
                                    </xsl:otherwise>
                                </xsl:choose>
                            </xsl:for-each-group>
                        </xsl:element>
                    </xsl:when>
                    <xsl:otherwise> <!-- not a list, a run of other elements -->
                        <xsl:apply-templates mode="listing" select="current-group()"/>
                    </xsl:otherwise>
                </xsl:choose>
            </xsl:for-each-group>
        </xsl:copy>
    </xsl:template>
    
    <!-- now strip attribs temporarily added to ListItem and SubListItem -->
    <xsl:template mode="listing" match="ListItem/@* | SubListItem/@*" />
    
    

    <!-- Boxing ============================================================ -->
    <!-- Deals with box-like structures that require matching xxxHead and closing xxxEnd paras: 
         Box, CaseStudy, Dialogue, Example, Extract, Quote, Reading, StudyNote, Verse, InternalSection, KeyPoints, Activity, Exercise, ITQ, SAQ
         but in fact triggered by any element containing 'Head' (not Heading or TableHead!)
        Note: use of xxxHead is consistent with existing LDS conversion 
        https://learn3.open.ac.uk/mod/oucontent/view.php?id=185740&extra=tablelandscape_idm78
    -->

    <!-- default identity: copy elements and attributes unchanged -->
    <xsl:template mode="boxing" match="* | @* ">
        <xsl:copy>
            <xsl:apply-templates mode="boxing" select="@* | node()"/>
        </xsl:copy>
    </xsl:template>

    <!-- Generic processing of box-like structures, eg Box, StudyNote,... and also Activity, SAQ... 
        Use for-each-group interating over Unit to structure boxes, selecting group 
        starting-with xxxHead and ending-with xxxEnd 
        Nb for-each-group returns *every* node including groups that don't match criterion so 
        retest always needed. -->
    <xsl:template mode="boxing" match="/" name="buildBoxes">
        <xsl:copy>
            <xsl:for-each-group select="*" group-starting-with="*[contains(name(), 'Head') and not(contains(name(), 'Heading')) and not(name() = 'TableHead') and not(name() = 'MultiColumnHead')]">
                <xsl:variable name="boxHead" select="name(current-group()[1])"/>
                <xsl:variable name="boxType" select="substring-before($boxHead, 'Head')"/>
                <xsl:variable name="boxEnd"  select="concat($boxType, 'End')"/>
                <!-- may get initial group that doesn't start-with 'Head' so boxType='' -->
                <xsl:choose>
                    <xsl:when test="$boxType != ''">
                        <xsl:for-each-group select="current-group()" group-ending-with="*[name() = $boxEnd]">
                            <xsl:choose>
                                <!-- check again, may get final group not in box -->
                                <xsl:when test="current-group()[1][name() = $boxHead]">
                                    <xsl:element name="{$boxType}">
                                        <xsl:apply-templates mode="boxing" select="@*"/>
                                        <xsl:apply-templates mode="boxing" select="current-group()"/>
                                    </xsl:element>
                                </xsl:when>
                                <xsl:otherwise> <!-- content after box -->
                                    <xsl:apply-templates mode="boxing" select="current-group()"/>
                                </xsl:otherwise>
                            </xsl:choose>
                        </xsl:for-each-group>
                    </xsl:when>
                    <xsl:otherwise> <!-- content before box -->
                        <xsl:apply-templates mode="boxing" select="current-group()"/>
                    </xsl:otherwise>
                </xsl:choose>
            </xsl:for-each-group>
        </xsl:copy>
    </xsl:template>
    
    <xsl:template mode="boxing" match="*[contains(name(), 'Head') and not(contains(name(), 'Heading')) and not(name() = 'TableHead') and not(name() = 'MultiColumnHead')]">
        <xsl:if test=". != '&#x00a0;'"> <!-- nbsp added by Word or sc-to-html for otherwise empty heading -->
            <Heading>
                <xsl:apply-templates mode="boxing" />
            </Heading>
        </xsl:if>
    </xsl:template>
    
    <xsl:template mode="boxing" match="*[contains(name(), 'End') and not(contains(name(), 'Append'))]"/>



    <!-- Questioning ======================================================= -->
    
    <!-- boxing has already dealt with Head/End and made containing structure, so 
        this pass only needs to build Question, Interaction, Answer, Discussion substructure 
        inside Activity or similar -->

    <!-- default identity: copy elements and attributes unchanged -->
    <xsl:template mode="questioning" match="* | @* ">
        <xsl:copy>
            <xsl:apply-templates mode="questioning" select="@* | node()"/>
        </xsl:copy>
    </xsl:template>

    <xsl:template mode="questioning" match="Activity | Exercise | ITQ | SAQ">
        <xsl:copy>
            <xsl:apply-templates mode="questioning" select="@*"/>
            <xsl:apply-templates mode="questioning" select="Heading"/>
            <xsl:apply-templates mode="questioning" select="Timing"/>
            <xsl:for-each-group select="*[not(self::Heading) and not(self::Timing)]" group-starting-with="Interaction | Answer | Discussion">
                <xsl:choose>
                    <xsl:when test="current-group()[1][self::Interaction]">
                        <Interaction>
                            <xsl:apply-templates mode="questioning" select="current-group()"/>
                        </Interaction>
                    </xsl:when>
                    <xsl:when test="current-group()[1][self::Answer]">
                        <Answer>
                            <xsl:apply-templates mode="questioning" select="current-group()"/>
                        </Answer>
                    </xsl:when>
                    <xsl:when test="current-group()[1][self::Discussion]">
                        <Discussion>
                            <xsl:apply-templates mode="questioning" select="current-group()"/>
                        </Discussion>
                    </xsl:when>
                    <xsl:otherwise> <!-- must be question -->
                        <Question>
                            <xsl:apply-templates mode="questioning" select="current-group()"/>
                        </Question>
                    </xsl:otherwise>
                </xsl:choose> 
            </xsl:for-each-group>
        </xsl:copy>
    </xsl:template>

    <!-- suppress since dealt with above -->
    <xsl:template mode="questioning" match="Question | Interaction | Answer | Discussion" />
        
    
    <!-- Sectioning ======================================================== -->

    <!-- Build nested Session/Section structure from flat file using h1/h2/h3...
        Code uses for-each-group starting-with to build structure. Note that iteration 
        must be done in a single pass but subgrouping achieved within current-group(). 
        Code is split into several template calls simply to avoid very deep indentation; 
        there is no recursion. -->

    <!-- default identity: copy elements and attributes unchanged -->
    <xsl:template mode="sectioning" match="* | @* ">
        <xsl:copy>
            <xsl:apply-templates mode="sectioning" select="@* | node()"/>
        </xsl:copy>
    </xsl:template>

    <xsl:template name="buildUnit">
        <xsl:for-each-group select="*" group-starting-with="h1">
            <xsl:choose>
                <xsl:when test="current-group()[1][self::h1]">
                    <xsl:call-template name="buildSession"/>
                </xsl:when>
                <xsl:otherwise>
                    <!-- content in parent but not in new child  -->
                    <xsl:apply-templates mode="sectioning" select="current-group()"/>
                </xsl:otherwise>
            </xsl:choose>
        </xsl:for-each-group>
    </xsl:template>
    
    <xsl:template name="buildSession">
        <Session>
            <xsl:apply-templates mode="sectioning" select="@*"/>
            <xsl:for-each-group select="current-group()" group-starting-with="h2">
                <xsl:choose>
                    <xsl:when test="current-group()[1][self::h2]">
                        <xsl:call-template name="buildSection"/>
                    </xsl:when>
                    <xsl:otherwise>
                        <!-- content in parent but not in new child  -->
                        <xsl:apply-templates mode="sectioning" select="current-group()"/>
                    </xsl:otherwise>
                </xsl:choose>
            </xsl:for-each-group>
        </Session>
    </xsl:template>

    <xsl:template name="buildSection">
        <Section>
            <xsl:apply-templates mode="sectioning" select="@*"/>
            <xsl:for-each-group select="current-group()" group-starting-with="h3">
                <xsl:choose>
                    <xsl:when test="current-group()[1][self::h3]">
                        <xsl:call-template name="buildSubSection"/>
                    </xsl:when>
                    <xsl:otherwise>
                        <!-- content in parent but not in new child  -->
                        <xsl:apply-templates mode="sectioning" select="current-group()"/>
                    </xsl:otherwise>
                </xsl:choose>
            </xsl:for-each-group>
        </Section>
    </xsl:template>

    <xsl:template name="buildSubSection">
        <SubSection>
            <xsl:apply-templates mode="sectioning" select="@*"/>
            <xsl:for-each-group select="current-group()" group-starting-with="h4">
                <xsl:choose>
                    <xsl:when test="current-group()[1][self::h4]">
                        <xsl:call-template name="buildSubSubSection"/>
                    </xsl:when>
                    <xsl:otherwise>
                        <!-- content in parent but not in new child  -->
                        <xsl:apply-templates mode="sectioning" select="current-group()"/>
                    </xsl:otherwise>
                </xsl:choose>
            </xsl:for-each-group>
        </SubSection>
    </xsl:template>

    <xsl:template name="buildSubSubSection">
        <SubSubSection>
            <xsl:apply-templates mode="sectioning" select="@*"/>
            <xsl:apply-templates mode="sectioning" select="current-group()"/>
        </SubSubSection>
    </xsl:template>
    
    
    <!-- start process of building sections at root -->
    <xsl:template mode="sectioning" match="/">
        <xsl:call-template name="buildUnit"/>
    </xsl:template>

    <!-- sectioning has already built structure so just convert heading text into Title -->
    <xsl:template mode="sectioning" match="h1 | h2 | h3">
        <Title>
            <xsl:apply-templates mode="sectioning" select="node()[not(self::id)]"/>
        </Title>
    </xsl:template>
    
    <xsl:template mode="sectioning" match="h4">
        <Heading>
            <xsl:apply-templates mode="sectioning" select="node()[not(self::id)]"/>
        </Heading>
    </xsl:template>


    <!-- Fixing ============================================================ -->
    
    <!-- A chance to fix outstanding issues... 
        - Sublists need hoisting into previous list item, rather than between items
        - Where list items contatin br, split into paragraphs
    -->
    
    <!-- default identity: copy elements and attributes unchanged -->
    <xsl:template mode="fixing fixing-force" match="* | @* ">
        <xsl:copy>
            <xsl:apply-templates mode="fixing" select="@* | node()"/>
        </xsl:copy>
    </xsl:template>
    
    <!-- find list items immediately followed by sublist -->
    <xsl:template mode="fixing" match="ListItem[following-sibling::*[1][self::BulletedSubsidiaryList | self::NumberedSubsidiaryList | self::UnNumberedSubsidiaryList]]">
        <ListItem>
            <xsl:apply-templates mode="fixing" select="node()"/>
            <!-- force copy out of order with special mode -->
            <xsl:apply-templates mode="fixing-force" select="following-sibling::*[1][self::BulletedSubsidiaryList | self::NumberedSubsidiaryList | self::UnNumberedSubsidiaryList]" />
        </ListItem>
    </xsl:template>
    
    <!-- suppress sublists not inside ListItem since dealt with above -->
    <xsl:template mode="fixing" match="(BulletedSubsidiaryList | NumberedSubsidiaryList | UnNumberedSubsidiaryList)[preceding-sibling::*[1][self::ListItem]]"/>
    
    <!-- Split ListItem and SubListItem paragaphs containing br -->
    <xsl:template mode="fixing" match="Paragraph[(parent::ListItem or parent::SubListItem) and child::br]">
        <xsl:for-each-group select="node()" group-ending-with="br">
            <Paragraph>
                <xsl:apply-templates mode="fixing" select="current-group()[not(self::br)]"/>
            </Paragraph>
        </xsl:for-each-group>
    </xsl:template>
    
    
    
</xsl:stylesheet>
