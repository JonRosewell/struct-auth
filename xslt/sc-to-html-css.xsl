<?xml version="1.0" encoding="UTF-8"?>
<xsl:stylesheet xmlns:xsl="http://www.w3.org/1999/XSL/Transform" version="3.0">
    <xsl:output method="html"/>

    <!-- Structured Authoring: MS Word to OU Structured Content XML
        MS Word to OU structured content conversion, designed to replace OU IT/LDS 
        customisation for oXygen.
        Jon Rosewell, Jan 2025
        https://github.com/JonRosewell/struct-auth 
        sc-to-html-css.xsl: include to style XML to Word conversion
    -->

    <!-- This is intended for inclusion into sc-to-html.xsl. It builds a CSS style sheet 
        for the html page produced by sc-to-html which in turn builds the Word stylesheet 
        when imported into Word. Every tag/style needs a minimal entry to preserve case, even 
        if not visually styled, hence tedious length. -->
    
    <xsl:template name="buildCSS">
        <style>
<!--        Start with full list of SC tags so that case is preserved on import to Word stylesheet 
            Some not used because:
            - they are html elements (b, i,...)
            - they are replaced by html (Session by h1, Section by h2,...)
            - they are box-like and are replaced by xxHead/xxEnd pairs 
            - they are meta data and not supported: front-matter details,... 
            - they are print-related: TOC,...
            A trade-off: completeness vs overload 
-->            
<!--        Following styles are not used:    
            p.Acknowledgements  { mso-style-name: Acknowledgements; }
            p.Activity  { mso-style-name: Activity; } 
            p.Address  { mso-style-name: Address; }
            p.AddressLine  { mso-style-name: AddressLine; }
            p.Appendices  { mso-style-name: Appendices; }
            p.Appendix  { mso-style-name: Appendix; }
            p.AudioReaderNote  { mso-style-name: AudioReaderNote; }
            p.b  { mso-style-name: b; } 
            p.BackMatter  { mso-style-name: BackMatter; }
            p.Box  { mso-style-name: Box; } 
            p.br  { mso-style-name: br; }
            p.BritishLibraryData  { mso-style-name: BritishLibraryData; }
            p.CaseStudy  { mso-style-name: CaseStudy; } 
            p.Conclusion  { mso-style-name: Conclusion; }
            p.CoPublished  { mso-style-name: CoPublished; }
            p.CopublisherAddress  { mso-style-name: CopublisherAddress; }
            p.Copyright  { mso-style-name: Copyright; }
            p.CourseTeam  { mso-style-name: CourseTeam; }
            p.CourseTitle  { mso-style-name: CourseTitle; }
            p.Cover  { mso-style-name: Cover; }
            p.Covers  { mso-style-name: Covers; }
            p.Dialogue  { mso-style-name: Dialogue; }
            p.Edited  { mso-style-name: Edited; }
            p.Edition  { mso-style-name: Edition; }
            p.Example  { mso-style-name: Example; }
            p.Exercise  { mso-style-name: Exercise; }
            p.Extract  { mso-style-name: Extract; }
            p.FirstPublished  { mso-style-name: FirstPublished; }
            p.font  { mso-style-name: font; }
            p.FrontMatter  { mso-style-name: FrontMatter; }
            p.FurtherReading  { mso-style-name: FurtherReading; }
            p.GeneralInfo  { mso-style-name: GeneralInfo; }
            p.HalfTitleVerso  { mso-style-name: HalfTitleVerso; }
            p.i  { mso-style-name: i; }
            p.Imprint  { mso-style-name: Imprint; }
            p.Index  { mso-style-name: Index; }
            p.Index1  { mso-style-name: Index1; }
            p.Index2  { mso-style-name: Index2; }
            p.Index3  { mso-style-name: Index3; }
            p.InternalSection  { mso-style-name: InternalSection; }
            p.ISBN  { mso-style-name: ISBN; }
            p.Item  { mso-style-name: Item; }
            p.ItemAcknowledgement  { mso-style-name: ItemAcknowledgement; }
            p.ItemID  { mso-style-name: ItemID; }
            p.ItemRef  { mso-style-name: ItemRef; }
            p.ItemRights  { mso-style-name: ItemRights; }
            p.ItemTitle  { mso-style-name: ItemTitle; }
            p.ITQ  { mso-style-name: ITQ; }
            p.KeyPoints  { mso-style-name: KeyPoints; }
            p.LearningOutcome  { mso-style-name: LearningOutcome; }
            p.LearningOutcomes  { mso-style-name: LearningOutcomes; }
            p.LibraryofCongressData  { mso-style-name: LibraryofCongressData; }
            p.ListItem  { mso-style-name: ListItem; }
            p.Logo  { mso-style-name: Logo; }
            p.Matching  { mso-style-name: Matching; }
            p.math  { mso-style-name: math; }
            p.MediaContent  { mso-style-name: MediaContent; }
            p.meta  { mso-style-name: meta; }
            p.Multipart  { mso-style-name: Multipart; }
            p.MultipleChoice  { mso-style-name: MultipleChoice; }
            p.OUCourseInfo  { mso-style-name: OUCourseInfo; }
            p.OUWebAddress  { mso-style-name: OUWebAddress; }
            span.OwnerRef  { mso-style-name: OwnerRef; }
            p.PageNumber  { mso-style-name: PageNumber; }
            p.Parameter  { mso-style-name: Parameter; }
            p.Parameters  { mso-style-name: Parameters; }
            p.Part  { mso-style-name: Part; }
            p.Preface  { mso-style-name: Preface; }
            p.Printed  { mso-style-name: Printed; }
            p.Promotion  { mso-style-name: Promotion; }
            p.Quote  { mso-style-name: Quote; }
            p.Reading  { mso-style-name: Reading; }
            p.Rights  { mso-style-name: Rights; }
            p.SAQ  { mso-style-name: SAQ; }
            p.Section  { mso-style-name: Section; }
            p.Session  { mso-style-name: Session; }
            p.SingleChoice  { mso-style-name: SingleChoice; }
            p.Standard  { mso-style-name: Standard; }
            p.StudyNote  { mso-style-name: StudyNote; }
            p.sub  { mso-style-name: sub; }
            p.SubListItem  { mso-style-name: SubListItem; }
            p.sup  { mso-style-name: sup; }
            p.sym  { mso-style-name: sym; }
            p.td  { mso-style-name: td; }
            p.th  { mso-style-name: th; }
            p.TOC  { mso-style-name: TOC; }
            p.TOC1  { mso-style-name: TOC1; }
            p.TOC2  { mso-style-name: TOC2; }
            p.TOC3  { mso-style-name: TOC3; }
            p.Total  { mso-style-name: Total; }
            p.Typeset  { mso-style-name: Typeset; }
            p.u  { mso-style-name: u; }
            p.Verse  { mso-style-name: Verse; }
-->
            
            <!-- following styles are used: preserve case of names in Word style sheet -->
            p.Alternative  { mso-style-name: Alternative; }
            p.Answer  { mso-style-name: Answer; }
            span.AuthorComment  { mso-style-name: AuthorComment; }
            p.BulletedList  { mso-style-name: BulletedList; }
            p.BulletedSubsidiaryList  { mso-style-name: BulletedSubsidiaryList; }
            p.ByLine  { mso-style-name: ByLine; }
            p.Caption  { mso-style-name: Caption; }
            p.Chemistry  { mso-style-name: Chemistry; }
            span.ComputerCode  { mso-style-name: ComputerCode; }
            p.ComputerDisplay  { mso-style-name: ComputerDisplay; }
            span.ComputerUI  { mso-style-name: ComputerUI; }
            p.CourseCode  { mso-style-name: CourseCode; }
            span.CrossRef  { mso-style-name: CrossRef; }
            p.Definition  { mso-style-name: Definition; }
            p.Description  { mso-style-name: Description; }
            p.Discussion  { mso-style-name: Discussion; }
            span.EditorComment  { mso-style-name: EditorComment; }
            p.Equation  { mso-style-name: Equation; }
            p.Figure  { mso-style-name: Figure; }
            p.footnote  { mso-style-name: footnote; }
            p.FreeResponse  { mso-style-name: FreeResponse; }
            p.FreeResponseDisplay  { mso-style-name: FreeResponseDisplay; }
            p.Glossary  { mso-style-name: Glossary; }
            p.GlossaryItem  { mso-style-name: GlossaryItem; }
            span.GlossaryTerm  { mso-style-name: GlossaryTerm; }
            p.Heading  { mso-style-name: Heading; }
            span.Hours  { mso-style-name: Hours; }
            p.Icon  { mso-style-name: Icon; }
            p.Image  { mso-style-name: Image; }
            span.IndexTerm  { mso-style-name: IndexTerm; }
            span.InlineChemistry  { mso-style-name: InlineChemistry; }
            span.InlineEquation  { mso-style-name: InlineEquation; }
            span.InlineFigure  { mso-style-name: InlineFigure; }
            span.InlinePageNumber  { mso-style-name: InlinePageNumber; }
            p.InPageActivity  { mso-style-name: InPageActivity; }
            p.Instructions  { mso-style-name: Instructions; }
            p.Interaction  { mso-style-name: Interaction; }
            p.Introduction  { mso-style-name: Introduction; }
            p.KeyPoint  { mso-style-name: KeyPoint; }
            span.Label { mso-style-name: Label; }
            span.language  { mso-style-name: language; }
            span.MathML  { mso-style-name: MathML; }
            span.Minutes  { mso-style-name: Minutes; }
            p.MultiColumnBody  { mso-style-name: MultiColumnBody; }
            p.MultiColumnHead  { mso-style-name: MultiColumnHead; }
            p.MultiColumnText  { mso-style-name: MultiColumnText; }
            span.Number  { mso-style-name: Number; }
            p.NumberedList  { mso-style-name: NumberedList; }
            p.NumberedSubsidiaryList  { mso-style-name: NumberedSubsidiaryList; }
            span.olink  { mso-style-name: olink; }
            p.Paragraph  { mso-style-name: Paragraph; }
            p.ProgramListing  { mso-style-name: ProgramListing; }
            p.Proof  { mso-style-name: Proof; }
            p.Question  { mso-style-name: Question; }
            p.Reference  { mso-style-name: Reference; }
            p.References  { mso-style-name: References; }
            p.Remark  { mso-style-name: Remark; }
            p.RevealMore  { mso-style-name: RevealMore; }
            span.SecondVoice  { mso-style-name: SecondVoice; }
            span.SideNote  { mso-style-name: SideNote; }
            span.SideNoteParagraph  { mso-style-name: SideNoteParagraph; }
            span.smallcaps  { mso-style-name: smallCaps; }
            p.SourceReference  { mso-style-name: SourceReference; }
            p.Speaker  { mso-style-name: Speaker; }
            p.SubHeading  { mso-style-name: SubHeading; }
            p.SubSection  { mso-style-name: SubSection; }
            p.SubSubHeading  { mso-style-name: SubSubHeading; }
            p.SubSubSection  { mso-style-name: SubSubSection; }
            p.Summary  { mso-style-name: Summary; }
            p.Table  { mso-style-name: Table; }
            p.TableFootnote  { mso-style-name: TableFootnote; }
            p.TableHead  { mso-style-name: TableHead; }
            p.Term  { mso-style-name: Term; }
            span.TeX  { mso-style-name: TeX; }
            p.Timing  { mso-style-name: Timing; }
            p.Title  { mso-style-name: Title; }
            p.Transcript  { mso-style-name: Transcript; }
            p.Unit  { mso-style-name: Unit; }
            p.UnitID  { mso-style-name: UnitID; }
            p.UnitTitle  { mso-style-name: UnitTitle; }
            p.UnNumberedList  { mso-style-name: UnNumberedList; }
            p.UnNumberedSubsidiaryList  { mso-style-name: UnNumberedSubsidiaryList; }
            p.VoiceRecorder  { mso-style-name: VoiceRecorder; }
            
<!--            jpr additions: some special purpose, most split of box-like into xxHead and xxEnd pairs -->
            span.attribute { mso-style-name: attribute; }
            p.RawXML { mso-style-name: RawXML; } 
            p.FigureSrc { mso-style-name: FigureSrc; } 
            span.SideNoteHeading  { mso-style-name: SideNoteHeading; }
            
            p.BoxHead  { mso-style-name: BoxHead; }
            p.CaseStudyHead  { mso-style-name: CaseStudyHead; }
            p.DialogueHead  { mso-style-name: DialogueHead; }
            p.ExampleHead  { mso-style-name: ExampleHead; }
            p.ExtractHead  { mso-style-name: ExtractHead; }
            p.KeyPointsHead  { mso-style-name: KeyPointsHead; }
            p.QuoteHead  { mso-style-name: QuoteHead; }
            p.ReadingHead  { mso-style-name: ReadingHead; }
            p.StudyNoteHead  { mso-style-name: StudyNoteHead; }
            p.VerseHead  { mso-style-name: VerseHead; }
            p.InternalSectionHead  { mso-style-name: InternalSectionHead; }
            p.ActivityHead  { mso-style-name: ActivityHead; }
            p.ExerciseHead  { mso-style-name: ExerciseHead; }
            p.ITQHead  { mso-style-name: ITQHead; }
            p.SAQHead  { mso-style-name: SAQHead; }
            p.ListHead  { mso-style-name: ListHead; }
            p.SubListHead  { mso-style-name: SubListHead; }
            
            p.BoxEnd  { mso-style-name: BoxEnd; }
            p.CaseStudyEnd  { mso-style-name: CaseStudyEnd; }
            p.DialogueEnd  { mso-style-name: DialogueEnd; }
            p.ExampleEnd  { mso-style-name: ExampleEnd; }
            p.ExtractEnd  { mso-style-name: ExtractEnd; }
            p.KeyPointsEnd  { mso-style-name: KeyPointsEnd; }
            p.QuoteEnd  { mso-style-name: QuoteEnd; }
            p.ReadingEnd  { mso-style-name: ReadingEnd; }
            p.StudyNoteEnd  { mso-style-name: StudyNoteEnd; }
            p.VerseEnd  { mso-style-name: VerseEnd; }
            p.InternalSectionEnd  { mso-style-name: InternalSectionEnd; }
            p.ActivityEnd  { mso-style-name: ActivityEnd; }
            p.ExerciseEnd  { mso-style-name: ExerciseEnd; }
            p.ITQEnd  { mso-style-name: ITQEnd; }
            p.SAQEnd  { mso-style-name: SAQEnd; }
            p.ListEnd  { mso-style-name: ListEnd; }            
            p.SubListEnd  { mso-style-name: SubListEnd; }            
            
<!--            appearance -->
            p, li, div, p.MsoNormal, li.MsoNormal, div.MsoNormal, table { font-size:11.0pt; font-family:"Calibri",sans-serif; }
            h1, h2, h3, h4, h5, h6 { font-family:"Calibri Light",sans-serif; color:#2F5496; }
            h1 { font-size: 18pt; } 
            h2 { font-size: 15pt; margin-left: 11pt; } 
            h3 { font-size: 13pt; margin-left: 22pt; } 
            h4 { font-size: 11pt; margin-left: 33pt; font-style: italic; } 
            table, th, td { border: solid windowtext .5pt; border-collapse: collapse; margin: 4pt }
            p.TableHead { border-bottom: solid #9CC2E5 1.5pt; font-weight: bold; }
            
            span.EditorComment { color: magenta; } 
            span.AuthorComment { color: darkorange; } 
            span.SecondVoice { color: #0070C0; } 
            span.ComputerUI { font-family: "Segoe UI", sans-serif; background: #FFFACD; mso-no-proof: yes; } 
            span.ComputerCode { font-family: "Ubuntu Mono", monospace; mso-no-proof: yes; } 
            span.GlossaryTerm { font-weight: bold; } 
            span.olink { color: blue; text-decoration: underline; mso-no-proof: yes; } 
            span.CrossRef { color: blue; text-decoration: underline; mso-no-proof: yes; }
            span.InlineFigure { color: blue; text-decoration: underline; mso-no-proof: yes; }
            span.attribute { color: lightgrey; mso-no-proof: yes; font-size: 8pt; }
            span.Number { background: #F2F2F2; font-weight: bold; mso-no-proof: yes;}
            span.SideNote, span.SideNoteHeading, span.SideNoteParagraph { background: #FFFF99; } 
            span.SideNoteHeading { font-weight: bold; } 
            span.Label { font-weight: bold; } 
            
            p.UnitTitle { font-size: 24pt; }
            span.TeX, span.MathML, span.InlineEquation, p.Equation { font-family: "Times New Roman", serif; color: purple; mso-no-proof: yes; } 
            p.Equation { font-size: 9pt; text-align: center;  }
            p.RawXML { font-size: 9pt; font-family: "Courier New", monospace; color: red; mso-no-proof: yes; } 
            p.ProgramListing { margin: 0pt; margin-left: 1cm; font-family: "Ubuntu Mono", monospace; mso-no-proof: yes; } 
            
            p.Figure { mso-style-next: FigureSrc; color: grey; } 
            p.FigureSrc { mso-style-next: Caption; font-style: italic; color: blue; } 
            p.Caption { mso-style-next: SourceReference; font-weight: bold; }
            p.SourceReference { mso-style-next: Description; font-size: 9.0pt; font-style: italic; text-align: right; } 
            p.Description { mso-style-next: Normal; color: teal; } 
            p.Alternative  { color: teal; font-style: italic; }
            
            p.BoxHead, p.BoxEnd, p.CaseStudyHead, p.CaseStudyEnd, p.ExampleHead, p.ExampleEnd, p.ExtractHead, p.ExtractEnd, .InternalSectionHead, p.InternalSectionEnd, p.KeyPointsHead, p.KeyPointsEnd, p.QuoteHead, p.QuoteEnd, p.ReadingHead, p.ReadingEnd, p.StudyNoteHead, p.StudyNoteEnd, p.ListHead, p.ListEnd { mso-style-next: Normal; } 
            p.BoxHead, p.BoxEnd  { margin: 6.0pt; background: #DEEAF6; border: solid windowtext .5pt; } 
            p.BoxHead { border-bottom: none; font-weight: bold; } 
            p.BoxEnd { border-top: none; } 
            p.StudyNoteHead, p.StudyNoteEnd { margin: 6.0pt; background: #E2EFD9; border: solid windowtext .5pt; } 
            p.StudyNoteHead { border-bottom: none; font-weight: bold; } 
            p.StudyNoteEnd { border-top: none; } 
            p.ExtractHead, p.ExtractEnd, p.KeyPointsHead, p.KeyPointsEnd, p.ReadingHead, p.ReadingEnd, p.CaseStudyHead, p.CaseStudyEnd { margin: 6.0pt; background: #FBE4D5; border: solid windowtext .5pt; } 
            p.ExampleHead, p.ExampleEnd { margin: 6.0pt; background: #D5F4F4; border: solid windowtext .5pt; } 
            p.ExampleHead, p.ExtractHead, p.KeyPointsHead, p.ReadingHead, p.CaseStudyHead { border-bottom: none;  font-weight: bold; } 
            p.ExampleEnd, p.ExtractEnd, p.KeyPointsEnd, p.ReadingEnd, p.CaseStudyEnd { border-top: none;  }
            p.InternalSectionHead, p.InternalSectionEnd { border: solid #F2F2F2 5pt; font-weight: bold; } 
            p.InternalSectionHead { border-bottom: none; } 
            p.InternalSectionEnd { border-top: none; } 
            p.QuoteHead, p.QuoteEnd { margin: 6.0pt; } 
            p.QuoteHead { border-bottom: solid lightgrey .5pt; font-weight: bold; } 
            p.QuoteEnd { border-top: solid lightgrey .5pt; } 
            p.ListHead, p.ListEnd, p.SubListHead, p.SubListEnd { font-size: 8pt; color:lightgrey; } 
            p.ListHead { border-bottom: dashed lightgrey 1pt; } 
            p.ListEnd { border-top: dashed lightgrey 1pt; } 
            p.SubListHead { border-bottom: dotted lightgrey 1pt; } 
            p.SubListEnd { border-top: dotted lightgrey 1pt; } 
            
            p.ActivityHead, p.Interaction, p.Answer, p.Discussion, p.ActivityEnd, p.SAQHead, p.SAQEnd, p.ITQHead, p.ITQEnd, p.ExerciseHead, p.ExerciseEnd { mso-style-next: Normal; color: #4472C4; background: #DEEAF6; } 
            p.ActivityHead, p.SAQHead, p.ITQHead, p.ExerciseHead { font-weight: bold; border-top: solid windowtext .5pt; }
            p.ActivityEnd, p.SAQEnd, p.ITQEnd, p.ExerciseEnd { font-weight: bold; border-bottom: solid windowtext .5pt; }
            p.Interaction, p.Answer, p.Discussion { font-style: italic; } 
            p.Timing { font-style: italic; font-size: 10pt;  } 
            p.RevealMore  { background: #F2CEED; border: solid lightgrey .5pt;  }
            
            p.Heading { mso-style-next: Normal; font-size: 12pt; font-weight: bold; color: red; } 
            p.SubHeading { mso-style-next: Normal; font-weight: bold; } 
            p.SubSubHeading { mso-style-next: Normal; font-style: italic; }
            
            p.ByLine { font-style: italic; }

        </style>
    </xsl:template>

</xsl:stylesheet>
