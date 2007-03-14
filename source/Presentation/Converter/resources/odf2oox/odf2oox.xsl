<?xml version="1.0" encoding="UTF-8"?>
<xsl:stylesheet version="1.0" 
	xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" 
 xmlns:xsl="http://www.w3.org/1999/XSL/Transform" 
  xmlns:odf="urn:odf"
  xmlns:pzip="urn:cleverage:xmlns:post-processings:zip"
  xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" 
  xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"
  xmlns:number="urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0"
  xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0"
  exclude-result-prefixes="odf style text number draw">

	<xsl:import href="docprops.xsl"/>
	

  <xsl:strip-space elements="*"/>
  <xsl:preserve-space elements="text:p text:span number:text"/>
  
  <xsl:param name="outputFile"/>
  <xsl:output method="xml" encoding="UTF-8"/>


  <xsl:variable name="app-version">1.00</xsl:variable>

  <!-- existence of docProps/custom.xml file -->
  <xsl:variable name="docprops-custom-file"
    select="count(document('meta.xml')/office:document-meta/office:meta/meta:user-defined)"
    xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
    xmlns:meta="urn:oasis:names:tc:opendocument:xmlns:meta:1.0"/>

	<xsl:template match="/odf:source">
		<xsl:processing-instruction name="mso-application">progid="ppt.Document"</xsl:processing-instruction>

		<pzip:archive pzip:target="{$outputFile}">

			<!-- sections preformatting -->
			<xsl:call-template name="dummy"/>

			<!-- Document core properties -->
			<pzip:entry pzip:target="docProps/core.xml">
				<xsl:call-template name="docprops-core"/>
			</pzip:entry>
			Document app properties
			<pzip:entry pzip:target="docProps/app.xml">
				<xsl:call-template name="docprops-app"/>
			</pzip:entry>

			Document custom properties
			<xsl:if test="$docprops-custom-file > 0">
				<pzip:entry pzip:target="docProps/custom.xml">
					<xsl:call-template name="docprops-custom"/>
				</pzip:entry>
			</xsl:if>
			<!--Main content-->
			<pzip:entry pzip:target="ppt/presentation.xml">
				<xsl:call-template name="tmpPresentation"/>
			</pzip:entry>
			<!--Numbering (lists)-->
			<pzip:entry pzip:target="ppt/presProps.xml">
				<xsl:call-template name="tmpPresProp"/>
			</pzip:entry>
			<!--numbering relationships item-->
			<pzip:entry pzip:target="ppt/_rels/presentation.xml.rels">
				<xsl:call-template name="tmpPresRel"/>
			</pzip:entry>
			<!--footnotes-->
			<pzip:entry pzip:target="ppt/tableStyles.xml">
				<xsl:call-template name="tmpTablestyle"/>
			</pzip:entry>
			<!--endnotes-->
			<pzip:entry pzip:target="ppt/viewProps.xml">
				<xsl:call-template name="tmpViewProps"/>
			</pzip:entry>
			<!--content types-->
			<pzip:entry pzip:target="[Content_Types].xml">
				<xsl:call-template name="tmpContentType"/>
			</pzip:entry>
			<!--package relationship item-->
			<pzip:entry pzip:target="_rels/.rels">
				<xsl:call-template name="tmpRel"/>
			</pzip:entry>
			<!--SLIDE LAYOUT FILES -->
			<pzip:entry pzip:target="ppt/slideLayouts/slideLayout1.xml">
				<xsl:call-template name="dummy"/>
			</pzip:entry>
			<pzip:entry pzip:target="ppt/slideLayouts/slideLayout2.xml">
				<xsl:call-template name="dummy"/>
			</pzip:entry>
			<pzip:entry pzip:target="ppt/slideLayouts/slideLayout3.xml">
				<xsl:call-template name="dummy"/>
			</pzip:entry>
			<pzip:entry pzip:target="ppt/slideLayouts/slideLayout4.xml">
				<xsl:call-template name="dummy"/>
			</pzip:entry>
			<pzip:entry pzip:target="ppt/slideLayouts/slideLayout5.xml">
				<xsl:call-template name="dummy"/>
			</pzip:entry>
			<pzip:entry pzip:target="ppt/slideLayouts/slideLayout6.xml">
				<xsl:call-template name="dummy"/>
			</pzip:entry>
			<pzip:entry pzip:target="ppt/slideLayouts/slideLayout7.xml">
				<xsl:call-template name="layout7"/>
			</pzip:entry>
			<pzip:entry pzip:target="ppt/slideLayouts/slideLayout8.xml">
				<xsl:call-template name="dummy"/>
			</pzip:entry>
			<pzip:entry pzip:target="ppt/slideLayouts/slideLayout9.xml">
				<xsl:call-template name="dummy"/>
			</pzip:entry>
			<pzip:entry pzip:target="ppt/slideLayouts/slideLayout10.xml">
				<xsl:call-template name="dummy"/>
			</pzip:entry>
			<pzip:entry pzip:target="ppt/slideLayouts/slideLayout11.xml">
				<xsl:call-template name="dummy"/>
			</pzip:entry>
			<pzip:entry pzip:target="ppt/slideLayouts/_rels/slideLayout1.xml">
				<xsl:call-template name="dummy"/>
			</pzip:entry>
			<pzip:entry pzip:target="ppt/slideLayouts/_rels/slideLayout2.xml">
				<xsl:call-template name="dummy"/>
			</pzip:entry>
			<pzip:entry pzip:target="ppt/slideLayouts/_rels/slideLayout3.xml">
				<xsl:call-template name="dummy"/>
			</pzip:entry>
			<pzip:entry pzip:target="ppt/slideLayouts/_rels/slideLayout4.xml">
				<xsl:call-template name="dummy"/>
			</pzip:entry>
			<pzip:entry pzip:target="ppt/slideLayouts/_rels/slideLayout5.xml">
				<xsl:call-template name="dummy"/>
			</pzip:entry>
			<pzip:entry pzip:target="ppt/slideLayouts/_rels/slideLayout6.xml">
				<xsl:call-template name="dummy"/>
			</pzip:entry>
			<pzip:entry pzip:target="ppt/slideLayouts/_rels/slideLayout7.xml">
				<xsl:call-template name="layout7Rel"/>
			</pzip:entry>
			<pzip:entry pzip:target="ppt/slideLayouts/_rels/slideLayout8.xml">
				<xsl:call-template name="dummy"/>
			</pzip:entry>
			<pzip:entry pzip:target="ppt/slideLayouts/_rels/slideLayout9.xml">
				<xsl:call-template name="dummy"/>
			</pzip:entry>
			<pzip:entry pzip:target="ppt/slideLayouts/_rels/slideLayout10.xml">
				<xsl:call-template name="dummy"/>
			</pzip:entry>
			<pzip:entry pzip:target="ppt/slideLayouts/_rels/slideLayout11.xml">
				<xsl:call-template name="dummy"/>
			</pzip:entry>
			<!--THEME-->
			<pzip:entry pzip:target="ppt/theme/theme1.xml">
				<xsl:call-template name="tmpTheme1"/>
			</pzip:entry>
			<pzip:entry pzip:target="ppt/theme/theme2.xml">
				<xsl:call-template name="tmpTheme2"/>
			</pzip:entry>
			
			<!--Slide List -->
			<pzip:entry pzip:target="ppt/slides/slide1.xml">
				<xsl:call-template name="tmpSlide"/>
			</pzip:entry>
			<pzip:entry pzip:target="ppt/slides/_rels/slide1.xml.rels">
				<xsl:call-template name="tmpSlide1rel"/>
			</pzip:entry>
			<!-- SLIDE MASTER -->
			<pzip:entry pzip:target="ppt/slideMasters/slideMaster1.xml.xml">
				<xsl:call-template name="SlideMaster"/>
			</pzip:entry>

			<pzip:entry pzip:target="ppt/slideMasters/_rels/slideMaster1.xml.rels">
				<xsl:call-template name="tmpSlideMasterRel"/>
			</pzip:entry>						  

	    </pzip:archive>
   </xsl:template>
	<xsl:template name="dummy">
		<body>
			<xsl:value-of select ="Pradeep"/>
		</body>
	</xsl:template>

	<!--for slide 1-->
	<xsl:template name="tmpSlide">
			<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" 
				xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" 
				xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
				<p:cSld>
					<p:spTree>
						<p:nvGrpSpPr>
							<p:cNvPr id="1" name=""/>
							<p:cNvGrpSpPr/>
							<p:nvPr/>
						</p:nvGrpSpPr>
						<p:grpSpPr>
							<a:xfrm>
								<a:off x="0" y="0"/>
								<a:ext cx="0" cy="0"/>
								<a:chOff x="0" y="0"/>
								<a:chExt cx="0" cy="0"/>
							</a:xfrm>
						</p:grpSpPr>
					</p:spTree>
				</p:cSld>
				<p:clrMapOvr>
					<a:masterClrMapping/>
				</p:clrMapOvr>
				<p:transition spd="med"/>
				<p:timing>
					<p:tnLst>
						<p:par>
							<p:cTn id="1" dur="indefinite" nodeType="tmRoot"/>
						</p:par>
					</p:tnLst>
				</p:timing>
			</p:sld>
		</xsl:template>


	<xsl:template name="tmpSlide1rel">		
		<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
			<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide" Target="../notesSlides/notesSlide1.xml"/>
			<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout7.xml"/>
		</Relationships>		
	</xsl:template >
	
	<!--slide layout 7-->
	<xsl:template name ="layout7">		
		<p:sldLayout xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" type="blank" preserve="1">
			<p:cSld name="Blank">
				<p:spTree>
					<p:nvGrpSpPr>
						<p:cNvPr id="1" name=""/>
						<p:cNvGrpSpPr/>
						<p:nvPr/>
					</p:nvGrpSpPr>
					<p:grpSpPr>
						<a:xfrm>
							<a:off x="0" y="0"/>
							<a:ext cx="0" cy="0"/>
							<a:chOff x="0" y="0"/>
							<a:chExt cx="0" cy="0"/>
						</a:xfrm>
					</p:grpSpPr>
					<p:sp>
						<p:nvSpPr>
							<p:cNvPr id="2" name="Rectangle 3"/>
							<p:cNvSpPr>
								<a:spLocks noGrp="1" noChangeArrowheads="1"/>
							</p:cNvSpPr>
							<p:nvPr>
								<p:ph type="dt" idx="10"/>
							</p:nvPr>
						</p:nvSpPr>
						<p:spPr>
							<a:ln/>
						</p:spPr>
						<p:txBody>
							<a:bodyPr/>
							<a:lstStyle>
								<a:lvl1pPr>
									<a:defRPr/>
								</a:lvl1pPr>
							</a:lstStyle>
							<a:p>
								<a:pPr>
									<a:defRPr/>
								</a:pPr>
								<a:endParaRPr lang="en-GB"/>
							</a:p>
						</p:txBody>
					</p:sp>
					<p:sp>
						<p:nvSpPr>
							<p:cNvPr id="3" name="Rectangle 4"/>
							<p:cNvSpPr>
								<a:spLocks noGrp="1" noChangeArrowheads="1"/>
							</p:cNvSpPr>
							<p:nvPr>
								<p:ph type="ftr" idx="11"/>
							</p:nvPr>
						</p:nvSpPr>
						<p:spPr>
							<a:ln/>
						</p:spPr>
						<p:txBody>
							<a:bodyPr/>
							<a:lstStyle>
								<a:lvl1pPr>
									<a:defRPr/>
								</a:lvl1pPr>
							</a:lstStyle>
							<a:p>
								<a:pPr>
									<a:defRPr/>
								</a:pPr>
								<a:endParaRPr lang="en-GB"/>
							</a:p>
						</p:txBody>
					</p:sp>
					<p:sp>
						<p:nvSpPr>
							<p:cNvPr id="4" name="Rectangle 5"/>
							<p:cNvSpPr>
								<a:spLocks noGrp="1" noChangeArrowheads="1"/>
							</p:cNvSpPr>
							<p:nvPr>
								<p:ph type="sldNum" idx="12"/>
							</p:nvPr>
						</p:nvSpPr>
						<p:spPr>
							<a:ln/>
						</p:spPr>
						<p:txBody>
							<a:bodyPr/>
							<a:lstStyle>
								<a:lvl1pPr>
									<a:defRPr/>
								</a:lvl1pPr>
							</a:lstStyle>
							<a:p>
								<a:pPr>
									<a:defRPr/>
								</a:pPr>
								<a:fld id="7AEF4067-DCF4-47CB-A5C7-3E4B40F6E446" type="slidenum">
									<a:rPr lang="en-GB"/>
									<a:pPr>
										<a:defRPr/>
									</a:pPr>
									<a:t>‹#›</a:t>
								</a:fld>
								<a:endParaRPr lang="en-GB"/>
							</a:p>
						</p:txBody>
					</p:sp>
				</p:spTree>
			</p:cSld>
			<p:clrMapOvr>
				<a:masterClrMapping/>
			</p:clrMapOvr>
		</p:sldLayout>
	</xsl:template>

	<xsl:template name ="layout7Rel">
		<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
			<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="../slideMasters/slideMaster1.xml"/>
		</Relationships>
	</xsl:template>

	<xsl:template name ="tmpPresentation">		
		<p:presentation xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" 
			xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" 
			xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" 
			strictFirstAndLastChars="0" saveSubsetFonts="1">
			<p:sldMasterIdLst>
				<p:sldMasterId id="2147483648" r:id="rId1"/>
			</p:sldMasterIdLst>
			<p:notesMasterIdLst>
				<p:notesMasterId r:id="rId3"/>
			</p:notesMasterIdLst>
			<p:sldIdLst>
				<p:sldId id="256" r:id="rId2"/>
			</p:sldIdLst>
			<p:sldSz cx="10080625" cy="7559675"/>
			<p:notesSz cx="7772400" cy="10058400"/>
			<p:defaultTextStyle>
				<a:defPPr>
					<a:defRPr lang="en-GB"/>
				</a:defPPr>
				<a:lvl1pPr algn="l" defTabSz="457200" rtl="0" fontAlgn="base" hangingPunct="0">
					<a:lnSpc>
						<a:spcPct val="93000"/>
					</a:lnSpc>
					<a:spcBef>
						<a:spcPct val="0"/>
					</a:spcBef>
					<a:spcAft>
						<a:spcPct val="0"/>
					</a:spcAft>
					<a:buClr>
						<a:srgbClr val="000000"/>
					</a:buClr>
					<a:buSzPct val="45000"/>
					<a:buFont typeface="StarSymbol" charset="0"/>
					<a:defRPr kern="1200">
						<a:solidFill>
							<a:schemeClr val="tx1"/>
						</a:solidFill>
						<a:latin typeface="Arial" charset="0"/>
						<a:ea typeface="+mn-ea"/>
						<a:cs typeface="Arial Unicode MS" charset="0"/>
					</a:defRPr>
				</a:lvl1pPr>
				<a:lvl2pPr marL="431800" indent="-215900" algn="l" defTabSz="457200" rtl="0" fontAlgn="base" hangingPunct="0">
					<a:lnSpc>
						<a:spcPct val="93000"/>
					</a:lnSpc>
					<a:spcBef>
						<a:spcPct val="0"/>
					</a:spcBef>
					<a:spcAft>
						<a:spcPct val="0"/>
					</a:spcAft>
					<a:buClr>
						<a:srgbClr val="000000"/>
					</a:buClr>
					<a:buSzPct val="45000"/>
					<a:buFont typeface="StarSymbol" charset="0"/>
					<a:defRPr kern="1200">
						<a:solidFill>
							<a:schemeClr val="tx1"/>
						</a:solidFill>
						<a:latin typeface="Arial" charset="0"/>
						<a:ea typeface="+mn-ea"/>
						<a:cs typeface="Arial Unicode MS" charset="0"/>
					</a:defRPr>
				</a:lvl2pPr>
				<a:lvl3pPr marL="647700" indent="-215900" algn="l" defTabSz="457200" rtl="0" fontAlgn="base" hangingPunct="0">
					<a:lnSpc>
						<a:spcPct val="93000"/>
					</a:lnSpc>
					<a:spcBef>
						<a:spcPct val="0"/>
					</a:spcBef>
					<a:spcAft>
						<a:spcPct val="0"/>
					</a:spcAft>
					<a:buClr>
						<a:srgbClr val="000000"/>
					</a:buClr>
					<a:buSzPct val="45000"/>
					<a:buFont typeface="StarSymbol" charset="0"/>
					<a:defRPr kern="1200">
						<a:solidFill>
							<a:schemeClr val="tx1"/>
						</a:solidFill>
						<a:latin typeface="Arial" charset="0"/>
						<a:ea typeface="+mn-ea"/>
						<a:cs typeface="Arial Unicode MS" charset="0"/>
					</a:defRPr>
				</a:lvl3pPr>
				<a:lvl4pPr marL="863600" indent="-215900" algn="l" defTabSz="457200" rtl="0" fontAlgn="base" hangingPunct="0">
					<a:lnSpc>
						<a:spcPct val="93000"/>
					</a:lnSpc>
					<a:spcBef>
						<a:spcPct val="0"/>
					</a:spcBef>
					<a:spcAft>
						<a:spcPct val="0"/>
					</a:spcAft>
					<a:buClr>
						<a:srgbClr val="000000"/>
					</a:buClr>
					<a:buSzPct val="45000"/>
					<a:buFont typeface="StarSymbol" charset="0"/>
					<a:defRPr kern="1200">
						<a:solidFill>
							<a:schemeClr val="tx1"/>
						</a:solidFill>
						<a:latin typeface="Arial" charset="0"/>
						<a:ea typeface="+mn-ea"/>
						<a:cs typeface="Arial Unicode MS" charset="0"/>
					</a:defRPr>
				</a:lvl4pPr>
				<a:lvl5pPr marL="1079500" indent="-215900" algn="l" defTabSz="457200" rtl="0" fontAlgn="base" hangingPunct="0">
					<a:lnSpc>
						<a:spcPct val="93000"/>
					</a:lnSpc>
					<a:spcBef>
						<a:spcPct val="0"/>
					</a:spcBef>
					<a:spcAft>
						<a:spcPct val="0"/>
					</a:spcAft>
					<a:buClr>
						<a:srgbClr val="000000"/>
					</a:buClr>
					<a:buSzPct val="45000"/>
					<a:buFont typeface="StarSymbol" charset="0"/>
					<a:defRPr kern="1200">
						<a:solidFill>
							<a:schemeClr val="tx1"/>
						</a:solidFill>
						<a:latin typeface="Arial" charset="0"/>
						<a:ea typeface="+mn-ea"/>
						<a:cs typeface="Arial Unicode MS" charset="0"/>
					</a:defRPr>
				</a:lvl5pPr>
				<a:lvl6pPr marL="2286000" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1">
					<a:defRPr kern="1200">
						<a:solidFill>
							<a:schemeClr val="tx1"/>
						</a:solidFill>
						<a:latin typeface="Arial" charset="0"/>
						<a:ea typeface="+mn-ea"/>
						<a:cs typeface="Arial Unicode MS" charset="0"/>
					</a:defRPr>
				</a:lvl6pPr>
				<a:lvl7pPr marL="2743200" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1">
					<a:defRPr kern="1200">
						<a:solidFill>
							<a:schemeClr val="tx1"/>
						</a:solidFill>
						<a:latin typeface="Arial" charset="0"/>
						<a:ea typeface="+mn-ea"/>
						<a:cs typeface="Arial Unicode MS" charset="0"/>
					</a:defRPr>
				</a:lvl7pPr>
				<a:lvl8pPr marL="3200400" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1">
					<a:defRPr kern="1200">
						<a:solidFill>
							<a:schemeClr val="tx1"/>
						</a:solidFill>
						<a:latin typeface="Arial" charset="0"/>
						<a:ea typeface="+mn-ea"/>
						<a:cs typeface="Arial Unicode MS" charset="0"/>
					</a:defRPr>
				</a:lvl8pPr>
				<a:lvl9pPr marL="3657600" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1">
					<a:defRPr kern="1200">
						<a:solidFill>
							<a:schemeClr val="tx1"/>
						</a:solidFill>
						<a:latin typeface="Arial" charset="0"/>
						<a:ea typeface="+mn-ea"/>
						<a:cs typeface="Arial Unicode MS" charset="0"/>
					</a:defRPr>
				</a:lvl9pPr>
			</p:defaultTextStyle>
		</p:presentation>
		
	</xsl:template>

	<xsl:template name ="tmpViewProps">		
		<p:viewPr xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" 
			xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" 
			xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
			<p:normalViewPr showOutlineIcons="0">
				<p:restoredLeft sz="15620"/>
				<p:restoredTop sz="94660"/>
			</p:normalViewPr>
			<p:slideViewPr>
				<p:cSldViewPr>
					<p:cViewPr varScale="1">
						<p:scale>
							<a:sx n="63" d="100"/>
							<a:sy n="63" d="100"/>
						</p:scale>
						<p:origin x="-594" y="-108"/>
					</p:cViewPr>
					<p:guideLst>
						<p:guide orient="horz" pos="2160"/>
						<p:guide pos="2880"/>
					</p:guideLst>
				</p:cSldViewPr>
			</p:slideViewPr>
			<p:outlineViewPr>
				<p:cViewPr varScale="1">
					<p:scale>
						<a:sx n="170" d="200"/>
						<a:sy n="170" d="200"/>
					</p:scale>
					<p:origin x="-780" y="-84"/>
				</p:cViewPr>
			</p:outlineViewPr>
			<p:notesTextViewPr>
				<p:cViewPr>
					<p:scale>
						<a:sx n="100" d="100"/>
						<a:sy n="100" d="100"/>
					</p:scale>
					<p:origin x="0" y="0"/>
				</p:cViewPr>
			</p:notesTextViewPr>
			<p:notesViewPr>
				<p:cSldViewPr>
					<p:cViewPr varScale="1">
						<p:scale>
							<a:sx n="59" d="100"/>
							<a:sy n="59" d="100"/>
						</p:scale>
						<p:origin x="-1752" y="-72"/>
					</p:cViewPr>
					<p:guideLst>
						<p:guide orient="horz" pos="2880"/>
						<p:guide pos="2160"/>
					</p:guideLst>
				</p:cSldViewPr>
			</p:notesViewPr>
			<p:gridSpacing cx="78028800" cy="78028800"/>
		</p:viewPr>
	</xsl:template>

	<xsl:template name ="tmpContentType">		
		<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
			<Override PartName="/ppt/slideLayouts/slideLayout7.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/>
			<Override PartName="/ppt/slideLayouts/slideLayout8.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/>
			<Override PartName="/ppt/notesSlides/notesSlide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml"/>
			<Override PartName="/ppt/slideMasters/slideMaster1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"/>
			<Override PartName="/ppt/presProps.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presProps+xml"/>
			<Override PartName="/ppt/slideLayouts/slideLayout4.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/>
			<Override PartName="/ppt/slideLayouts/slideLayout5.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/>
			<Override PartName="/ppt/slideLayouts/slideLayout6.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/>
			<Override PartName="/ppt/theme/theme2.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>
			<Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>
			<Override PartName="/ppt/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>
			<Override PartName="/ppt/slideLayouts/slideLayout2.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/>
			<Override PartName="/ppt/slideLayouts/slideLayout3.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/>
			<Default Extension="jpeg" ContentType="image/jpeg"/>
			<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
			<Default Extension="xml" ContentType="application/xml"/>
			<Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>
			<Override PartName="/ppt/notesMasters/notesMaster1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.notesMaster+xml"/>
			<Override PartName="/ppt/slideLayouts/slideLayout1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/>
			<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
			<Override PartName="/ppt/tableStyles.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.tableStyles+xml"/>
			<Override PartName="/ppt/slideLayouts/slideLayout11.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/>
			<Override PartName="/ppt/slideLayouts/slideLayout10.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/>
			<Override PartName="/ppt/viewProps.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.viewProps+xml"/>
			<Override PartName="/ppt/slideLayouts/slideLayout9.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/>
			<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
		</Types>
	</xsl:template>

	<xsl:template name ="tmpPresProp">		
		<p:presentationPr xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
			<p:showPr showNarration="1">
				<p:present/>
				<p:sldAll/>
				<p:penClr>
					<a:schemeClr val="tx1"/>
				</p:penClr>
			</p:showPr>
		</p:presentationPr>
	</xsl:template>

	<xsl:template name ="tmpTablestyle">
		<a:tblStyleLst xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" 
			def="5C22544A-7EE6-4342-B048-85BDC9FD1C3A"/>
	</xsl:template>

	<xsl:template name ="SlideMaster">
		<p:sldMaster xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" 
			xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" 
			xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
			<p:cSld>
				<p:bg>
					<p:bgPr>
						<a:solidFill>
							<a:srgbClr val="FFFFFF"/>
						</a:solidFill>
						<a:effectLst/>
					</p:bgPr>
				</p:bg>
				<p:spTree>
					<p:nvGrpSpPr>
						<p:cNvPr id="1" name=""/>
						<p:cNvGrpSpPr/>
						<p:nvPr/>
					</p:nvGrpSpPr>
					<p:grpSpPr>
						<a:xfrm>
							<a:off x="0" y="0"/>
							<a:ext cx="0" cy="0"/>
							<a:chOff x="0" y="0"/>
							<a:chExt cx="0" cy="0"/>
						</a:xfrm>
					</p:grpSpPr>
					<p:sp>
						<p:nvSpPr>
							<p:cNvPr id="1026" name="Rectangle 1"/>
							<p:cNvSpPr>
								<a:spLocks noGrp="1" noChangeArrowheads="1"/>
							</p:cNvSpPr>
							<p:nvPr>
								<p:ph type="title"/>
							</p:nvPr>
						</p:nvSpPr>
						<p:spPr bwMode="auto">
							<a:xfrm>
								<a:off x="503238" y="301625"/>
								<a:ext cx="9070975" cy="1260475"/>
							</a:xfrm>
							<a:prstGeom prst="rect">
								<a:avLst/>
							</a:prstGeom>
							<a:noFill/>
							<a:ln w="9525">
								<a:noFill/>
								<a:round/>
								<a:headEnd/>
								<a:tailEnd/>
							</a:ln>
						</p:spPr>
						<p:txBody>
							<a:bodyPr vert="horz" wrap="square" lIns="0" tIns="0" rIns="0" bIns="0" numCol="1" anchor="ctr" anchorCtr="0" compatLnSpc="1">
								<a:prstTxWarp prst="textNoShape">
									<a:avLst/>
								</a:prstTxWarp>
							</a:bodyPr>
							<a:lstStyle/>
							<a:p>
								<a:pPr lvl="0"/>
								<a:r>
									<a:rPr lang="en-GB" smtClean="0"/>
									<a:t>Click to edit the title text format</a:t>
								</a:r>
							</a:p>
						</p:txBody>
					</p:sp>
					<p:sp>
						<p:nvSpPr>
							<p:cNvPr id="1027" name="Rectangle 2"/>
							<p:cNvSpPr>
								<a:spLocks noGrp="1" noChangeArrowheads="1"/>
							</p:cNvSpPr>
							<p:nvPr>
								<p:ph type="body" idx="1"/>
							</p:nvPr>
						</p:nvSpPr>
						<p:spPr bwMode="auto">
							<a:xfrm>
								<a:off x="503238" y="1768475"/>
								<a:ext cx="9070975" cy="4987925"/>
							</a:xfrm>
							<a:prstGeom prst="rect">
								<a:avLst/>
							</a:prstGeom>
							<a:noFill/>
							<a:ln w="9525">
								<a:noFill/>
								<a:round/>
								<a:headEnd/>
								<a:tailEnd/>
							</a:ln>
						</p:spPr>
						<p:txBody>
							<a:bodyPr vert="horz" wrap="square" lIns="0" tIns="0" rIns="0" bIns="0" numCol="1" anchor="t" anchorCtr="0" compatLnSpc="1">
								<a:prstTxWarp prst="textNoShape">
									<a:avLst/>
								</a:prstTxWarp>
							</a:bodyPr>
							<a:lstStyle/>
							<a:p>
								<a:pPr lvl="0"/>
								<a:r>
									<a:rPr lang="en-GB" smtClean="0"/>
									<a:t>Click to edit the outline text format</a:t>
								</a:r>
							</a:p>
							<a:p>
								<a:pPr lvl="1"/>
								<a:r>
									<a:rPr lang="en-GB" smtClean="0"/>
									<a:t>Second Outline Level</a:t>
								</a:r>
							</a:p>
							<a:p>
								<a:pPr lvl="2"/>
								<a:r>
									<a:rPr lang="en-GB" smtClean="0"/>
									<a:t>Third Outline Level</a:t>
								</a:r>
							</a:p>
							<a:p>
								<a:pPr lvl="3"/>
								<a:r>
									<a:rPr lang="en-GB" smtClean="0"/>
									<a:t>Fourth Outline Level</a:t>
								</a:r>
							</a:p>
							<a:p>
								<a:pPr lvl="4"/>
								<a:r>
									<a:rPr lang="en-GB" smtClean="0"/>
									<a:t>Fifth Outline Level</a:t>
								</a:r>
							</a:p>
							<a:p>
								<a:pPr lvl="4"/>
								<a:r>
									<a:rPr lang="en-GB" smtClean="0"/>
									<a:t>Sixth Outline Level</a:t>
								</a:r>
							</a:p>
							<a:p>
								<a:pPr lvl="4"/>
								<a:r>
									<a:rPr lang="en-GB" smtClean="0"/>
									<a:t>Seventh Outline Level</a:t>
								</a:r>
							</a:p>
							<a:p>
								<a:pPr lvl="4"/>
								<a:r>
									<a:rPr lang="en-GB" smtClean="0"/>
									<a:t>Eighth Outline Level</a:t>
								</a:r>
							</a:p>
							<a:p>
								<a:pPr lvl="4"/>
								<a:r>
									<a:rPr lang="en-GB" smtClean="0"/>
									<a:t>Ninth Outline Level</a:t>
								</a:r>
							</a:p>
						</p:txBody>
					</p:sp>
					<p:sp>
						<p:nvSpPr>
							<p:cNvPr id="2" name="Rectangle 3"/>
							<p:cNvSpPr>
								<a:spLocks noGrp="1" noChangeArrowheads="1"/>
							</p:cNvSpPr>
							<p:nvPr>
								<p:ph type="dt"/>
							</p:nvPr>
						</p:nvSpPr>
						<p:spPr bwMode="auto">
							<a:xfrm>
								<a:off x="503238" y="6886575"/>
								<a:ext cx="2346325" cy="520700"/>
							</a:xfrm>
							<a:prstGeom prst="rect">
								<a:avLst/>
							</a:prstGeom>
							<a:noFill/>
							<a:ln w="9525">
								<a:noFill/>
								<a:round/>
								<a:headEnd/>
								<a:tailEnd/>
							</a:ln>
							<a:effectLst/>
						</p:spPr>
						<p:txBody>
							<a:bodyPr vert="horz" wrap="square" lIns="0" tIns="0" rIns="0" bIns="0" numCol="1" anchor="t" anchorCtr="0" compatLnSpc="1">
								<a:prstTxWarp prst="textNoShape">
									<a:avLst/>
								</a:prstTxWarp>
							</a:bodyPr>
							<a:lstStyle>
								<a:lvl1pPr>
									<a:lnSpc>
										<a:spcPct val="95000"/>
									</a:lnSpc>
									<a:tabLst>
										<a:tab pos="723900" algn="l"/>
										<a:tab pos="1447800" algn="l"/>
										<a:tab pos="2171700" algn="l"/>
									</a:tabLst>
									<a:defRPr sz="1400" smtClean="0">
										<a:solidFill>
											<a:srgbClr val="000000"/>
										</a:solidFill>
										<a:latin typeface="Times New Roman" pitchFamily="16" charset="0"/>
									</a:defRPr>
								</a:lvl1pPr>
							</a:lstStyle>
							<a:p>
								<a:pPr>
									<a:defRPr/>
								</a:pPr>
								<a:endParaRPr lang="en-GB"/>
							</a:p>
						</p:txBody>
					</p:sp>
					<p:sp>
						<p:nvSpPr>
							<p:cNvPr id="1028" name="Rectangle 4"/>
							<p:cNvSpPr>
								<a:spLocks noGrp="1" noChangeArrowheads="1"/>
							</p:cNvSpPr>
							<p:nvPr>
								<p:ph type="ftr"/>
							</p:nvPr>
						</p:nvSpPr>
						<p:spPr bwMode="auto">
							<a:xfrm>
								<a:off x="3448050" y="6886575"/>
								<a:ext cx="3194050" cy="520700"/>
							</a:xfrm>
							<a:prstGeom prst="rect">
								<a:avLst/>
							</a:prstGeom>
							<a:noFill/>
							<a:ln w="9525">
								<a:noFill/>
								<a:round/>
								<a:headEnd/>
								<a:tailEnd/>
							</a:ln>
							<a:effectLst/>
						</p:spPr>
						<p:txBody>
							<a:bodyPr vert="horz" wrap="square" lIns="0" tIns="0" rIns="0" bIns="0" numCol="1" anchor="t" anchorCtr="0" compatLnSpc="1">
								<a:prstTxWarp prst="textNoShape">
									<a:avLst/>
								</a:prstTxWarp>
							</a:bodyPr>
							<a:lstStyle>
								<a:lvl1pPr algn="ctr">
									<a:lnSpc>
										<a:spcPct val="95000"/>
									</a:lnSpc>
									<a:tabLst>
										<a:tab pos="723900" algn="l"/>
										<a:tab pos="1447800" algn="l"/>
										<a:tab pos="2171700" algn="l"/>
										<a:tab pos="2895600" algn="l"/>
									</a:tabLst>
									<a:defRPr sz="1400" smtClean="0">
										<a:solidFill>
											<a:srgbClr val="000000"/>
										</a:solidFill>
										<a:latin typeface="Times New Roman" pitchFamily="16" charset="0"/>
									</a:defRPr>
								</a:lvl1pPr>
							</a:lstStyle>
							<a:p>
								<a:pPr>
									<a:defRPr/>
								</a:pPr>
								<a:endParaRPr lang="en-GB"/>
							</a:p>
						</p:txBody>
					</p:sp>
					<p:sp>
						<p:nvSpPr>
							<p:cNvPr id="1029" name="Rectangle 5"/>
							<p:cNvSpPr>
								<a:spLocks noGrp="1" noChangeArrowheads="1"/>
							</p:cNvSpPr>
							<p:nvPr>
								<p:ph type="sldNum"/>
							</p:nvPr>
						</p:nvSpPr>
						<p:spPr bwMode="auto">
							<a:xfrm>
								<a:off x="7227888" y="6886575"/>
								<a:ext cx="2346325" cy="520700"/>
							</a:xfrm>
							<a:prstGeom prst="rect">
								<a:avLst/>
							</a:prstGeom>
							<a:noFill/>
							<a:ln w="9525">
								<a:noFill/>
								<a:round/>
								<a:headEnd/>
								<a:tailEnd/>
							</a:ln>
							<a:effectLst/>
						</p:spPr>
						<p:txBody>
							<a:bodyPr vert="horz" wrap="square" lIns="0" tIns="0" rIns="0" bIns="0" numCol="1" anchor="t" anchorCtr="0" compatLnSpc="1">
								<a:prstTxWarp prst="textNoShape">
									<a:avLst/>
								</a:prstTxWarp>
							</a:bodyPr>
							<a:lstStyle>
								<a:lvl1pPr algn="r">
									<a:lnSpc>
										<a:spcPct val="95000"/>
									</a:lnSpc>
									<a:tabLst>
										<a:tab pos="723900" algn="l"/>
										<a:tab pos="1447800" algn="l"/>
										<a:tab pos="2171700" algn="l"/>
									</a:tabLst>
									<a:defRPr sz="1400" smtClean="0">
										<a:solidFill>
											<a:srgbClr val="000000"/>
										</a:solidFill>
										<a:latin typeface="Times New Roman" pitchFamily="16" charset="0"/>
									</a:defRPr>
								</a:lvl1pPr>
							</a:lstStyle>
							<a:p>
								<a:pPr>
									<a:defRPr/>
								</a:pPr>
								<a:fld id="154800B4-7A45-4719-B4BD-6F8C5D2C98D3" type="slidenum">
									<a:rPr lang="en-GB"/>
									<a:pPr>
										<a:defRPr/>
									</a:pPr>
									<a:t>‹#›</a:t>
								</a:fld>
								<a:endParaRPr lang="en-GB"/>
							</a:p>
						</p:txBody>
					</p:sp>
				</p:spTree>
			</p:cSld>
			<p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/>
			<p:sldLayoutIdLst>
				<p:sldLayoutId id="2147483649" r:id="rId1"/>
				<p:sldLayoutId id="2147483650" r:id="rId2"/>
				<p:sldLayoutId id="2147483651" r:id="rId3"/>
				<p:sldLayoutId id="2147483652" r:id="rId4"/>
				<p:sldLayoutId id="2147483653" r:id="rId5"/>
				<p:sldLayoutId id="2147483654" r:id="rId6"/>
				<p:sldLayoutId id="2147483655" r:id="rId7"/>
				<p:sldLayoutId id="2147483656" r:id="rId8"/>
				<p:sldLayoutId id="2147483657" r:id="rId9"/>
				<p:sldLayoutId id="2147483658" r:id="rId10"/>
				<p:sldLayoutId id="2147483659" r:id="rId11"/>
			</p:sldLayoutIdLst>
			<p:txStyles>
				<p:titleStyle>
					<a:lvl1pPr algn="ctr" defTabSz="457200" rtl="0" eaLnBrk="0" fontAlgn="base" hangingPunct="0">
						<a:lnSpc>
							<a:spcPct val="93000"/>
						</a:lnSpc>
						<a:spcBef>
							<a:spcPct val="0"/>
						</a:spcBef>
						<a:spcAft>
							<a:spcPct val="0"/>
						</a:spcAft>
						<a:buClr>
							<a:srgbClr val="000000"/>
						</a:buClr>
						<a:buSzPct val="45000"/>
						<a:buFont typeface="StarSymbol" charset="0"/>
						<a:defRPr sz="4400">
							<a:solidFill>
								<a:srgbClr val="000000"/>
							</a:solidFill>
							<a:latin typeface="+mj-lt"/>
							<a:ea typeface="+mj-ea"/>
							<a:cs typeface="+mj-cs"/>
						</a:defRPr>
					</a:lvl1pPr>
					<a:lvl2pPr algn="ctr" defTabSz="457200" rtl="0" eaLnBrk="0" fontAlgn="base" hangingPunct="0">
						<a:lnSpc>
							<a:spcPct val="93000"/>
						</a:lnSpc>
						<a:spcBef>
							<a:spcPct val="0"/>
						</a:spcBef>
						<a:spcAft>
							<a:spcPct val="0"/>
						</a:spcAft>
						<a:buClr>
							<a:srgbClr val="000000"/>
						</a:buClr>
						<a:buSzPct val="45000"/>
						<a:buFont typeface="StarSymbol" charset="0"/>
						<a:defRPr sz="4400">
							<a:solidFill>
								<a:srgbClr val="000000"/>
							</a:solidFill>
							<a:latin typeface="Arial" charset="0"/>
							<a:cs typeface="Arial Unicode MS" charset="0"/>
						</a:defRPr>
					</a:lvl2pPr>
					<a:lvl3pPr algn="ctr" defTabSz="457200" rtl="0" eaLnBrk="0" fontAlgn="base" hangingPunct="0">
						<a:lnSpc>
							<a:spcPct val="93000"/>
						</a:lnSpc>
						<a:spcBef>
							<a:spcPct val="0"/>
						</a:spcBef>
						<a:spcAft>
							<a:spcPct val="0"/>
						</a:spcAft>
						<a:buClr>
							<a:srgbClr val="000000"/>
						</a:buClr>
						<a:buSzPct val="45000"/>
						<a:buFont typeface="StarSymbol" charset="0"/>
						<a:defRPr sz="4400">
							<a:solidFill>
								<a:srgbClr val="000000"/>
							</a:solidFill>
							<a:latin typeface="Arial" charset="0"/>
							<a:cs typeface="Arial Unicode MS" charset="0"/>
						</a:defRPr>
					</a:lvl3pPr>
					<a:lvl4pPr algn="ctr" defTabSz="457200" rtl="0" eaLnBrk="0" fontAlgn="base" hangingPunct="0">
						<a:lnSpc>
							<a:spcPct val="93000"/>
						</a:lnSpc>
						<a:spcBef>
							<a:spcPct val="0"/>
						</a:spcBef>
						<a:spcAft>
							<a:spcPct val="0"/>
						</a:spcAft>
						<a:buClr>
							<a:srgbClr val="000000"/>
						</a:buClr>
						<a:buSzPct val="45000"/>
						<a:buFont typeface="StarSymbol" charset="0"/>
						<a:defRPr sz="4400">
							<a:solidFill>
								<a:srgbClr val="000000"/>
							</a:solidFill>
							<a:latin typeface="Arial" charset="0"/>
							<a:cs typeface="Arial Unicode MS" charset="0"/>
						</a:defRPr>
					</a:lvl4pPr>
					<a:lvl5pPr algn="ctr" defTabSz="457200" rtl="0" eaLnBrk="0" fontAlgn="base" hangingPunct="0">
						<a:lnSpc>
							<a:spcPct val="93000"/>
						</a:lnSpc>
						<a:spcBef>
							<a:spcPct val="0"/>
						</a:spcBef>
						<a:spcAft>
							<a:spcPct val="0"/>
						</a:spcAft>
						<a:buClr>
							<a:srgbClr val="000000"/>
						</a:buClr>
						<a:buSzPct val="45000"/>
						<a:buFont typeface="StarSymbol" charset="0"/>
						<a:defRPr sz="4400">
							<a:solidFill>
								<a:srgbClr val="000000"/>
							</a:solidFill>
							<a:latin typeface="Arial" charset="0"/>
							<a:cs typeface="Arial Unicode MS" charset="0"/>
						</a:defRPr>
					</a:lvl5pPr>
					<a:lvl6pPr marL="1536700" indent="-215900" algn="ctr" defTabSz="457200" rtl="0" fontAlgn="base" hangingPunct="0">
						<a:lnSpc>
							<a:spcPct val="93000"/>
						</a:lnSpc>
						<a:spcBef>
							<a:spcPct val="0"/>
						</a:spcBef>
						<a:spcAft>
							<a:spcPct val="0"/>
						</a:spcAft>
						<a:buClr>
							<a:srgbClr val="000000"/>
						</a:buClr>
						<a:buSzPct val="45000"/>
						<a:buFont typeface="StarSymbol" charset="0"/>
						<a:defRPr sz="4400">
							<a:solidFill>
								<a:srgbClr val="000000"/>
							</a:solidFill>
							<a:latin typeface="Arial" charset="0"/>
							<a:cs typeface="Arial Unicode MS" charset="0"/>
						</a:defRPr>
					</a:lvl6pPr>
					<a:lvl7pPr marL="1993900" indent="-215900" algn="ctr" defTabSz="457200" rtl="0" fontAlgn="base" hangingPunct="0">
						<a:lnSpc>
							<a:spcPct val="93000"/>
						</a:lnSpc>
						<a:spcBef>
							<a:spcPct val="0"/>
						</a:spcBef>
						<a:spcAft>
							<a:spcPct val="0"/>
						</a:spcAft>
						<a:buClr>
							<a:srgbClr val="000000"/>
						</a:buClr>
						<a:buSzPct val="45000"/>
						<a:buFont typeface="StarSymbol" charset="0"/>
						<a:defRPr sz="4400">
							<a:solidFill>
								<a:srgbClr val="000000"/>
							</a:solidFill>
							<a:latin typeface="Arial" charset="0"/>
							<a:cs typeface="Arial Unicode MS" charset="0"/>
						</a:defRPr>
					</a:lvl7pPr>
					<a:lvl8pPr marL="2451100" indent="-215900" algn="ctr" defTabSz="457200" rtl="0" fontAlgn="base" hangingPunct="0">
						<a:lnSpc>
							<a:spcPct val="93000"/>
						</a:lnSpc>
						<a:spcBef>
							<a:spcPct val="0"/>
						</a:spcBef>
						<a:spcAft>
							<a:spcPct val="0"/>
						</a:spcAft>
						<a:buClr>
							<a:srgbClr val="000000"/>
						</a:buClr>
						<a:buSzPct val="45000"/>
						<a:buFont typeface="StarSymbol" charset="0"/>
						<a:defRPr sz="4400">
							<a:solidFill>
								<a:srgbClr val="000000"/>
							</a:solidFill>
							<a:latin typeface="Arial" charset="0"/>
							<a:cs typeface="Arial Unicode MS" charset="0"/>
						</a:defRPr>
					</a:lvl8pPr>
					<a:lvl9pPr marL="2908300" indent="-215900" algn="ctr" defTabSz="457200" rtl="0" fontAlgn="base" hangingPunct="0">
						<a:lnSpc>
							<a:spcPct val="93000"/>
						</a:lnSpc>
						<a:spcBef>
							<a:spcPct val="0"/>
						</a:spcBef>
						<a:spcAft>
							<a:spcPct val="0"/>
						</a:spcAft>
						<a:buClr>
							<a:srgbClr val="000000"/>
						</a:buClr>
						<a:buSzPct val="45000"/>
						<a:buFont typeface="StarSymbol" charset="0"/>
						<a:defRPr sz="4400">
							<a:solidFill>
								<a:srgbClr val="000000"/>
							</a:solidFill>
							<a:latin typeface="Arial" charset="0"/>
							<a:cs typeface="Arial Unicode MS" charset="0"/>
						</a:defRPr>
					</a:lvl9pPr>
				</p:titleStyle>
				<p:bodyStyle>
					<a:lvl1pPr marL="431800" indent="-323850" algn="l" defTabSz="457200" rtl="0" eaLnBrk="0" fontAlgn="base" hangingPunct="0">
						<a:lnSpc>
							<a:spcPct val="93000"/>
						</a:lnSpc>
						<a:spcBef>
							<a:spcPct val="0"/>
						</a:spcBef>
						<a:spcAft>
							<a:spcPts val="1425"/>
						</a:spcAft>
						<a:buClr>
							<a:srgbClr val="000000"/>
						</a:buClr>
						<a:buSzPct val="45000"/>
						<a:buFont typeface="StarSymbol" charset="0"/>
						<a:buChar char="●"/>
						<a:defRPr sz="3200">
							<a:solidFill>
								<a:srgbClr val="000000"/>
							</a:solidFill>
							<a:latin typeface="+mn-lt"/>
							<a:ea typeface="+mn-ea"/>
							<a:cs typeface="+mn-cs"/>
						</a:defRPr>
					</a:lvl1pPr>
					<a:lvl2pPr marL="863600" indent="-287338" algn="l" defTabSz="457200" rtl="0" eaLnBrk="0" fontAlgn="base" hangingPunct="0">
						<a:lnSpc>
							<a:spcPct val="93000"/>
						</a:lnSpc>
						<a:spcBef>
							<a:spcPct val="0"/>
						</a:spcBef>
						<a:spcAft>
							<a:spcPts val="1138"/>
						</a:spcAft>
						<a:buClr>
							<a:srgbClr val="000000"/>
						</a:buClr>
						<a:buSzPct val="75000"/>
						<a:buFont typeface="Symbol" charset="2"/>
						<a:buChar char=""/>
						<a:defRPr sz="2800">
							<a:solidFill>
								<a:srgbClr val="000000"/>
							</a:solidFill>
							<a:latin typeface="+mn-lt"/>
							<a:cs typeface="+mn-cs"/>
						</a:defRPr>
					</a:lvl2pPr>
					<a:lvl3pPr marL="1295400" indent="-215900" algn="l" defTabSz="457200" rtl="0" eaLnBrk="0" fontAlgn="base" hangingPunct="0">
						<a:lnSpc>
							<a:spcPct val="93000"/>
						</a:lnSpc>
						<a:spcBef>
							<a:spcPct val="0"/>
						</a:spcBef>
						<a:spcAft>
							<a:spcPts val="850"/>
						</a:spcAft>
						<a:buClr>
							<a:srgbClr val="000000"/>
						</a:buClr>
						<a:buSzPct val="45000"/>
						<a:buFont typeface="StarSymbol" charset="0"/>
						<a:buChar char="●"/>
						<a:defRPr sz="2400">
							<a:solidFill>
								<a:srgbClr val="000000"/>
							</a:solidFill>
							<a:latin typeface="+mn-lt"/>
							<a:cs typeface="+mn-cs"/>
						</a:defRPr>
					</a:lvl3pPr>
					<a:lvl4pPr marL="1727200" indent="-215900" algn="l" defTabSz="457200" rtl="0" eaLnBrk="0" fontAlgn="base" hangingPunct="0">
						<a:lnSpc>
							<a:spcPct val="93000"/>
						</a:lnSpc>
						<a:spcBef>
							<a:spcPct val="0"/>
						</a:spcBef>
						<a:spcAft>
							<a:spcPts val="575"/>
						</a:spcAft>
						<a:buClr>
							<a:srgbClr val="000000"/>
						</a:buClr>
						<a:buSzPct val="75000"/>
						<a:buFont typeface="Symbol" charset="2"/>
						<a:buChar char=""/>
						<a:defRPr sz="2000">
							<a:solidFill>
								<a:srgbClr val="000000"/>
							</a:solidFill>
							<a:latin typeface="+mn-lt"/>
							<a:cs typeface="+mn-cs"/>
						</a:defRPr>
					</a:lvl4pPr>
					<a:lvl5pPr marL="2159000" indent="-215900" algn="l" defTabSz="457200" rtl="0" eaLnBrk="0" fontAlgn="base" hangingPunct="0">
						<a:lnSpc>
							<a:spcPct val="93000"/>
						</a:lnSpc>
						<a:spcBef>
							<a:spcPct val="0"/>
						</a:spcBef>
						<a:spcAft>
							<a:spcPts val="288"/>
						</a:spcAft>
						<a:buClr>
							<a:srgbClr val="000000"/>
						</a:buClr>
						<a:buSzPct val="45000"/>
						<a:buFont typeface="StarSymbol" charset="0"/>
						<a:buChar char="●"/>
						<a:defRPr sz="2000">
							<a:solidFill>
								<a:srgbClr val="000000"/>
							</a:solidFill>
							<a:latin typeface="+mn-lt"/>
							<a:cs typeface="+mn-cs"/>
						</a:defRPr>
					</a:lvl5pPr>
					<a:lvl6pPr marL="2616200" indent="-215900" algn="l" defTabSz="457200" rtl="0" fontAlgn="base" hangingPunct="0">
						<a:lnSpc>
							<a:spcPct val="93000"/>
						</a:lnSpc>
						<a:spcBef>
							<a:spcPct val="0"/>
						</a:spcBef>
						<a:spcAft>
							<a:spcPts val="288"/>
						</a:spcAft>
						<a:buClr>
							<a:srgbClr val="000000"/>
						</a:buClr>
						<a:buSzPct val="45000"/>
						<a:buFont typeface="StarSymbol" charset="0"/>
						<a:buChar char="●"/>
						<a:defRPr sz="2000">
							<a:solidFill>
								<a:srgbClr val="000000"/>
							</a:solidFill>
							<a:latin typeface="+mn-lt"/>
							<a:cs typeface="+mn-cs"/>
						</a:defRPr>
					</a:lvl6pPr>
					<a:lvl7pPr marL="3073400" indent="-215900" algn="l" defTabSz="457200" rtl="0" fontAlgn="base" hangingPunct="0">
						<a:lnSpc>
							<a:spcPct val="93000"/>
						</a:lnSpc>
						<a:spcBef>
							<a:spcPct val="0"/>
						</a:spcBef>
						<a:spcAft>
							<a:spcPts val="288"/>
						</a:spcAft>
						<a:buClr>
							<a:srgbClr val="000000"/>
						</a:buClr>
						<a:buSzPct val="45000"/>
						<a:buFont typeface="StarSymbol" charset="0"/>
						<a:buChar char="●"/>
						<a:defRPr sz="2000">
							<a:solidFill>
								<a:srgbClr val="000000"/>
							</a:solidFill>
							<a:latin typeface="+mn-lt"/>
							<a:cs typeface="+mn-cs"/>
						</a:defRPr>
					</a:lvl7pPr>
					<a:lvl8pPr marL="3530600" indent="-215900" algn="l" defTabSz="457200" rtl="0" fontAlgn="base" hangingPunct="0">
						<a:lnSpc>
							<a:spcPct val="93000"/>
						</a:lnSpc>
						<a:spcBef>
							<a:spcPct val="0"/>
						</a:spcBef>
						<a:spcAft>
							<a:spcPts val="288"/>
						</a:spcAft>
						<a:buClr>
							<a:srgbClr val="000000"/>
						</a:buClr>
						<a:buSzPct val="45000"/>
						<a:buFont typeface="StarSymbol" charset="0"/>
						<a:buChar char="●"/>
						<a:defRPr sz="2000">
							<a:solidFill>
								<a:srgbClr val="000000"/>
							</a:solidFill>
							<a:latin typeface="+mn-lt"/>
							<a:cs typeface="+mn-cs"/>
						</a:defRPr>
					</a:lvl8pPr>
					<a:lvl9pPr marL="3987800" indent="-215900" algn="l" defTabSz="457200" rtl="0" fontAlgn="base" hangingPunct="0">
						<a:lnSpc>
							<a:spcPct val="93000"/>
						</a:lnSpc>
						<a:spcBef>
							<a:spcPct val="0"/>
						</a:spcBef>
						<a:spcAft>
							<a:spcPts val="288"/>
						</a:spcAft>
						<a:buClr>
							<a:srgbClr val="000000"/>
						</a:buClr>
						<a:buSzPct val="45000"/>
						<a:buFont typeface="StarSymbol" charset="0"/>
						<a:buChar char="●"/>
						<a:defRPr sz="2000">
							<a:solidFill>
								<a:srgbClr val="000000"/>
							</a:solidFill>
							<a:latin typeface="+mn-lt"/>
							<a:cs typeface="+mn-cs"/>
						</a:defRPr>
					</a:lvl9pPr>
				</p:bodyStyle>
				<p:otherStyle>
					<a:defPPr>
						<a:defRPr lang="en-US"/>
					</a:defPPr>
					<a:lvl1pPr marL="0" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1">
						<a:defRPr sz="1800" kern="1200">
							<a:solidFill>
								<a:schemeClr val="tx1"/>
							</a:solidFill>
							<a:latin typeface="+mn-lt"/>
							<a:ea typeface="+mn-ea"/>
							<a:cs typeface="+mn-cs"/>
						</a:defRPr>
					</a:lvl1pPr>
					<a:lvl2pPr marL="457200" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1">
						<a:defRPr sz="1800" kern="1200">
							<a:solidFill>
								<a:schemeClr val="tx1"/>
							</a:solidFill>
							<a:latin typeface="+mn-lt"/>
							<a:ea typeface="+mn-ea"/>
							<a:cs typeface="+mn-cs"/>
						</a:defRPr>
					</a:lvl2pPr>
					<a:lvl3pPr marL="914400" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1">
						<a:defRPr sz="1800" kern="1200">
							<a:solidFill>
								<a:schemeClr val="tx1"/>
							</a:solidFill>
							<a:latin typeface="+mn-lt"/>
							<a:ea typeface="+mn-ea"/>
							<a:cs typeface="+mn-cs"/>
						</a:defRPr>
					</a:lvl3pPr>
					<a:lvl4pPr marL="1371600" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1">
						<a:defRPr sz="1800" kern="1200">
							<a:solidFill>
								<a:schemeClr val="tx1"/>
							</a:solidFill>
							<a:latin typeface="+mn-lt"/>
							<a:ea typeface="+mn-ea"/>
							<a:cs typeface="+mn-cs"/>
						</a:defRPr>
					</a:lvl4pPr>
					<a:lvl5pPr marL="1828800" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1">
						<a:defRPr sz="1800" kern="1200">
							<a:solidFill>
								<a:schemeClr val="tx1"/>
							</a:solidFill>
							<a:latin typeface="+mn-lt"/>
							<a:ea typeface="+mn-ea"/>
							<a:cs typeface="+mn-cs"/>
						</a:defRPr>
					</a:lvl5pPr>
					<a:lvl6pPr marL="2286000" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1">
						<a:defRPr sz="1800" kern="1200">
							<a:solidFill>
								<a:schemeClr val="tx1"/>
							</a:solidFill>
							<a:latin typeface="+mn-lt"/>
							<a:ea typeface="+mn-ea"/>
							<a:cs typeface="+mn-cs"/>
						</a:defRPr>
					</a:lvl6pPr>
					<a:lvl7pPr marL="2743200" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1">
						<a:defRPr sz="1800" kern="1200">
							<a:solidFill>
								<a:schemeClr val="tx1"/>
							</a:solidFill>
							<a:latin typeface="+mn-lt"/>
							<a:ea typeface="+mn-ea"/>
							<a:cs typeface="+mn-cs"/>
						</a:defRPr>
					</a:lvl7pPr>
					<a:lvl8pPr marL="3200400" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1">
						<a:defRPr sz="1800" kern="1200">
							<a:solidFill>
								<a:schemeClr val="tx1"/>
							</a:solidFill>
							<a:latin typeface="+mn-lt"/>
							<a:ea typeface="+mn-ea"/>
							<a:cs typeface="+mn-cs"/>
						</a:defRPr>
					</a:lvl8pPr>
					<a:lvl9pPr marL="3657600" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1">
						<a:defRPr sz="1800" kern="1200">
							<a:solidFill>
								<a:schemeClr val="tx1"/>
							</a:solidFill>
							<a:latin typeface="+mn-lt"/>
							<a:ea typeface="+mn-ea"/>
							<a:cs typeface="+mn-cs"/>
						</a:defRPr>
					</a:lvl9pPr>
				</p:otherStyle>
			</p:txStyles>
		</p:sldMaster>
	</xsl:template>

	<xsl:template name ="tmpSlideMasterRel">		
		<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
			<Relationship Id="rId8" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout8.xml"/>
			<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout3.xml"/>
			<Relationship Id="rId7" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout7.xml"/>
			<Relationship Id="rId12" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="../theme/theme1.xml"/>
			<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout2.xml"/>
			<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>
			<Relationship Id="rId6" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout6.xml"/>
			<Relationship Id="rId11" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout11.xml"/>
			<Relationship Id="rId5" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout5.xml"/>
			<Relationship Id="rId10" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout10.xml"/>
			<Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout4.xml"/>
			<Relationship Id="rId9" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout9.xml"/>
		</Relationships>
	</xsl:template>

	<xsl:template name ="tmpPresRel">		
		<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
			<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesMaster" Target="notesMasters/notesMaster1.xml"/>
			<Relationship Id="rId7" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/tableStyles" Target="tableStyles.xml"/>
			<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/>
			<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="slideMasters/slideMaster1.xml"/>
			<Relationship Id="rId6" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/>
			<Relationship Id="rId5" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/viewProps" Target="viewProps.xml"/>
			<Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/presProps" Target="presProps.xml"/>
		</Relationships>
	</xsl:template>

	<xsl:template name ="tmpTheme1">		
		<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Office Theme">
			<a:themeElements>
				<a:clrScheme name="Office Theme 1">
					<a:dk1>
						<a:srgbClr val="000000"/>
					</a:dk1>
					<a:lt1>
						<a:srgbClr val="FFFFFF"/>
					</a:lt1>
					<a:dk2>
						<a:srgbClr val="000000"/>
					</a:dk2>
					<a:lt2>
						<a:srgbClr val="808080"/>
					</a:lt2>
					<a:accent1>
						<a:srgbClr val="00CC99"/>
					</a:accent1>
					<a:accent2>
						<a:srgbClr val="3333CC"/>
					</a:accent2>
					<a:accent3>
						<a:srgbClr val="FFFFFF"/>
					</a:accent3>
					<a:accent4>
						<a:srgbClr val="000000"/>
					</a:accent4>
					<a:accent5>
						<a:srgbClr val="AAE2CA"/>
					</a:accent5>
					<a:accent6>
						<a:srgbClr val="2D2DB9"/>
					</a:accent6>
					<a:hlink>
						<a:srgbClr val="CCCCFF"/>
					</a:hlink>
					<a:folHlink>
						<a:srgbClr val="B2B2B2"/>
					</a:folHlink>
				</a:clrScheme>
				<a:fontScheme name="Office Theme">
					<a:majorFont>
						<a:latin typeface="Arial"/>
						<a:ea typeface=""/>
						<a:cs typeface="Arial Unicode MS"/>
					</a:majorFont>
					<a:minorFont>
						<a:latin typeface="Arial"/>
						<a:ea typeface=""/>
						<a:cs typeface="Arial Unicode MS"/>
					</a:minorFont>
				</a:fontScheme>
				<a:fmtScheme name="Office">
					<a:fillStyleLst>
						<a:solidFill>
							<a:schemeClr val="phClr"/>
						</a:solidFill>
						<a:gradFill rotWithShape="1">
							<a:gsLst>
								<a:gs pos="0">
									<a:schemeClr val="phClr">
										<a:tint val="50000"/>
										<a:satMod val="300000"/>
									</a:schemeClr>
								</a:gs>
								<a:gs pos="35000">
									<a:schemeClr val="phClr">
										<a:tint val="37000"/>
										<a:satMod val="300000"/>
									</a:schemeClr>
								</a:gs>
								<a:gs pos="100000">
									<a:schemeClr val="phClr">
										<a:tint val="15000"/>
										<a:satMod val="350000"/>
									</a:schemeClr>
								</a:gs>
							</a:gsLst>
							<a:lin ang="16200000" scaled="1"/>
						</a:gradFill>
						<a:gradFill rotWithShape="1">
							<a:gsLst>
								<a:gs pos="0">
									<a:schemeClr val="phClr">
										<a:shade val="51000"/>
										<a:satMod val="130000"/>
									</a:schemeClr>
								</a:gs>
								<a:gs pos="80000">
									<a:schemeClr val="phClr">
										<a:shade val="93000"/>
										<a:satMod val="130000"/>
									</a:schemeClr>
								</a:gs>
								<a:gs pos="100000">
									<a:schemeClr val="phClr">
										<a:shade val="94000"/>
										<a:satMod val="135000"/>
									</a:schemeClr>
								</a:gs>
							</a:gsLst>
							<a:lin ang="16200000" scaled="0"/>
						</a:gradFill>
					</a:fillStyleLst>
					<a:lnStyleLst>
						<a:ln w="9525" cap="flat" cmpd="sng" algn="ctr">
							<a:solidFill>
								<a:schemeClr val="phClr">
									<a:shade val="95000"/>
									<a:satMod val="105000"/>
								</a:schemeClr>
							</a:solidFill>
							<a:prstDash val="solid"/>
						</a:ln>
						<a:ln w="25400" cap="flat" cmpd="sng" algn="ctr">
							<a:solidFill>
								<a:schemeClr val="phClr"/>
							</a:solidFill>
							<a:prstDash val="solid"/>
						</a:ln>
						<a:ln w="38100" cap="flat" cmpd="sng" algn="ctr">
							<a:solidFill>
								<a:schemeClr val="phClr"/>
							</a:solidFill>
							<a:prstDash val="solid"/>
						</a:ln>
					</a:lnStyleLst>
					<a:effectStyleLst>
						<a:effectStyle>
							<a:effectLst>
								<a:outerShdw blurRad="40000" dist="20000" dir="5400000" rotWithShape="0">
									<a:srgbClr val="000000">
										<a:alpha val="38000"/>
									</a:srgbClr>
								</a:outerShdw>
							</a:effectLst>
						</a:effectStyle>
						<a:effectStyle>
							<a:effectLst>
								<a:outerShdw blurRad="40000" dist="23000" dir="5400000" rotWithShape="0">
									<a:srgbClr val="000000">
										<a:alpha val="35000"/>
									</a:srgbClr>
								</a:outerShdw>
							</a:effectLst>
						</a:effectStyle>
						<a:effectStyle>
							<a:effectLst>
								<a:outerShdw blurRad="40000" dist="23000" dir="5400000" rotWithShape="0">
									<a:srgbClr val="000000">
										<a:alpha val="35000"/>
									</a:srgbClr>
								</a:outerShdw>
							</a:effectLst>
							<a:scene3d>
								<a:camera prst="orthographicFront">
									<a:rot lat="0" lon="0" rev="0"/>
								</a:camera>
								<a:lightRig rig="threePt" dir="t">
									<a:rot lat="0" lon="0" rev="1200000"/>
								</a:lightRig>
							</a:scene3d>
							<a:sp3d>
								<a:bevelT w="63500" h="25400"/>
							</a:sp3d>
						</a:effectStyle>
					</a:effectStyleLst>
					<a:bgFillStyleLst>
						<a:solidFill>
							<a:schemeClr val="phClr"/>
						</a:solidFill>
						<a:gradFill rotWithShape="1">
							<a:gsLst>
								<a:gs pos="0">
									<a:schemeClr val="phClr">
										<a:tint val="40000"/>
										<a:satMod val="350000"/>
									</a:schemeClr>
								</a:gs>
								<a:gs pos="40000">
									<a:schemeClr val="phClr">
										<a:tint val="45000"/>
										<a:shade val="99000"/>
										<a:satMod val="350000"/>
									</a:schemeClr>
								</a:gs>
								<a:gs pos="100000">
									<a:schemeClr val="phClr">
										<a:shade val="20000"/>
										<a:satMod val="255000"/>
									</a:schemeClr>
								</a:gs>
							</a:gsLst>
							<a:path path="circle">
								<a:fillToRect l="50000" t="-80000" r="50000" b="180000"/>
							</a:path>
						</a:gradFill>
						<a:gradFill rotWithShape="1">
							<a:gsLst>
								<a:gs pos="0">
									<a:schemeClr val="phClr">
										<a:tint val="80000"/>
										<a:satMod val="300000"/>
									</a:schemeClr>
								</a:gs>
								<a:gs pos="100000">
									<a:schemeClr val="phClr">
										<a:shade val="30000"/>
										<a:satMod val="200000"/>
									</a:schemeClr>
								</a:gs>
							</a:gsLst>
							<a:path path="circle">
								<a:fillToRect l="50000" t="50000" r="50000" b="50000"/>
							</a:path>
						</a:gradFill>
					</a:bgFillStyleLst>
				</a:fmtScheme>
			</a:themeElements>
			<a:objectDefaults>
				<a:spDef>
					<a:spPr bwMode="auto">
						<a:xfrm>
							<a:off x="0" y="0"/>
							<a:ext cx="1" cy="1"/>
						</a:xfrm>
						<a:custGeom>
							<a:avLst/>
							<a:gdLst/>
							<a:ahLst/>
							<a:cxnLst/>
							<a:rect l="0" t="0" r="0" b="0"/>
							<a:pathLst/>
						</a:custGeom>
						<a:solidFill>
							<a:srgbClr val="00B8FF"/>
						</a:solidFill>
						<a:ln w="9525" cap="flat" cmpd="sng" algn="ctr">
							<a:solidFill>
								<a:schemeClr val="tx1"/>
							</a:solidFill>
							<a:prstDash val="solid"/>
							<a:round/>
							<a:headEnd type="none" w="med" len="med"/>
							<a:tailEnd type="none" w="med" len="med"/>
						</a:ln>
						<a:effectLst/>
					</a:spPr>
					<a:bodyPr vert="horz" wrap="square" lIns="91440" tIns="45720" rIns="91440" bIns="45720" numCol="1" anchor="t" anchorCtr="0" compatLnSpc="1">
						<a:prstTxWarp prst="textNoShape">
							<a:avLst/>
						</a:prstTxWarp>
					</a:bodyPr>
					<a:lstStyle>
						<a:defPPr marL="0" marR="0" indent="0" algn="l" defTabSz="457200" rtl="0" eaLnBrk="1" fontAlgn="base" latinLnBrk="0" hangingPunct="0">
							<a:lnSpc>
								<a:spcPct val="93000"/>
							</a:lnSpc>
							<a:spcBef>
								<a:spcPct val="0"/>
							</a:spcBef>
							<a:spcAft>
								<a:spcPct val="0"/>
							</a:spcAft>
							<a:buClr>
								<a:srgbClr val="000000"/>
							</a:buClr>
							<a:buSzPct val="45000"/>
							<a:buFont typeface="StarSymbol" charset="0"/>
							<a:buNone/>
							<a:tabLst/>
							<a:defRPr kumimoji="0" lang="en-GB" sz="1800" b="0" i="0" u="none" strike="noStrike" cap="none" normalizeH="0" baseline="0" smtClean="0">
								<a:ln>
									<a:noFill/>
								</a:ln>
								<a:effectLst/>
								<a:latin typeface="Arial" charset="0"/>
								<a:cs typeface="Arial Unicode MS" charset="0"/>
							</a:defRPr>
						</a:defPPr>
					</a:lstStyle>
				</a:spDef>
				<a:lnDef>
					<a:spPr bwMode="auto">
						<a:xfrm>
							<a:off x="0" y="0"/>
							<a:ext cx="1" cy="1"/>
						</a:xfrm>
						<a:custGeom>
							<a:avLst/>
							<a:gdLst/>
							<a:ahLst/>
							<a:cxnLst/>
							<a:rect l="0" t="0" r="0" b="0"/>
							<a:pathLst/>
						</a:custGeom>
						<a:solidFill>
							<a:srgbClr val="00B8FF"/>
						</a:solidFill>
						<a:ln w="9525" cap="flat" cmpd="sng" algn="ctr">
							<a:solidFill>
								<a:schemeClr val="tx1"/>
							</a:solidFill>
							<a:prstDash val="solid"/>
							<a:round/>
							<a:headEnd type="none" w="med" len="med"/>
							<a:tailEnd type="none" w="med" len="med"/>
						</a:ln>
						<a:effectLst/>
					</a:spPr>
					<a:bodyPr vert="horz" wrap="square" lIns="91440" tIns="45720" rIns="91440" bIns="45720" numCol="1" anchor="t" anchorCtr="0" compatLnSpc="1">
						<a:prstTxWarp prst="textNoShape">
							<a:avLst/>
						</a:prstTxWarp>
					</a:bodyPr>
					<a:lstStyle>
						<a:defPPr marL="0" marR="0" indent="0" algn="l" defTabSz="457200" rtl="0" eaLnBrk="1" fontAlgn="base" latinLnBrk="0" hangingPunct="0">
							<a:lnSpc>
								<a:spcPct val="93000"/>
							</a:lnSpc>
							<a:spcBef>
								<a:spcPct val="0"/>
							</a:spcBef>
							<a:spcAft>
								<a:spcPct val="0"/>
							</a:spcAft>
							<a:buClr>
								<a:srgbClr val="000000"/>
							</a:buClr>
							<a:buSzPct val="45000"/>
							<a:buFont typeface="StarSymbol" charset="0"/>
							<a:buNone/>
							<a:tabLst/>
							<a:defRPr kumimoji="0" lang="en-GB" sz="1800" b="0" i="0" u="none" strike="noStrike" cap="none" normalizeH="0" baseline="0" smtClean="0">
								<a:ln>
									<a:noFill/>
								</a:ln>
								<a:effectLst/>
								<a:latin typeface="Arial" charset="0"/>
								<a:cs typeface="Arial Unicode MS" charset="0"/>
							</a:defRPr>
						</a:defPPr>
					</a:lstStyle>
				</a:lnDef>
			</a:objectDefaults>
			<a:extraClrSchemeLst>
				<a:extraClrScheme>
					<a:clrScheme name="Office Theme 1">
						<a:dk1>
							<a:srgbClr val="000000"/>
						</a:dk1>
						<a:lt1>
							<a:srgbClr val="FFFFFF"/>
						</a:lt1>
						<a:dk2>
							<a:srgbClr val="000000"/>
						</a:dk2>
						<a:lt2>
							<a:srgbClr val="808080"/>
						</a:lt2>
						<a:accent1>
							<a:srgbClr val="00CC99"/>
						</a:accent1>
						<a:accent2>
							<a:srgbClr val="3333CC"/>
						</a:accent2>
						<a:accent3>
							<a:srgbClr val="FFFFFF"/>
						</a:accent3>
						<a:accent4>
							<a:srgbClr val="000000"/>
						</a:accent4>
						<a:accent5>
							<a:srgbClr val="AAE2CA"/>
						</a:accent5>
						<a:accent6>
							<a:srgbClr val="2D2DB9"/>
						</a:accent6>
						<a:hlink>
							<a:srgbClr val="CCCCFF"/>
						</a:hlink>
						<a:folHlink>
							<a:srgbClr val="B2B2B2"/>
						</a:folHlink>
					</a:clrScheme>
					<a:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/>
				</a:extraClrScheme>
				<a:extraClrScheme>
					<a:clrScheme name="Office Theme 2">
						<a:dk1>
							<a:srgbClr val="000000"/>
						</a:dk1>
						<a:lt1>
							<a:srgbClr val="FFFFFF"/>
						</a:lt1>
						<a:dk2>
							<a:srgbClr val="0000FF"/>
						</a:dk2>
						<a:lt2>
							<a:srgbClr val="FFFF00"/>
						</a:lt2>
						<a:accent1>
							<a:srgbClr val="FF9900"/>
						</a:accent1>
						<a:accent2>
							<a:srgbClr val="00FFFF"/>
						</a:accent2>
						<a:accent3>
							<a:srgbClr val="AAAAFF"/>
						</a:accent3>
						<a:accent4>
							<a:srgbClr val="DADADA"/>
						</a:accent4>
						<a:accent5>
							<a:srgbClr val="FFCAAA"/>
						</a:accent5>
						<a:accent6>
							<a:srgbClr val="00E7E7"/>
						</a:accent6>
						<a:hlink>
							<a:srgbClr val="FF0000"/>
						</a:hlink>
						<a:folHlink>
							<a:srgbClr val="969696"/>
						</a:folHlink>
					</a:clrScheme>
					<a:clrMap bg1="dk2" tx1="lt1" bg2="dk1" tx2="lt2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/>
				</a:extraClrScheme>
				<a:extraClrScheme>
					<a:clrScheme name="Office Theme 3">
						<a:dk1>
							<a:srgbClr val="000000"/>
						</a:dk1>
						<a:lt1>
							<a:srgbClr val="FFFFCC"/>
						</a:lt1>
						<a:dk2>
							<a:srgbClr val="808000"/>
						</a:dk2>
						<a:lt2>
							<a:srgbClr val="666633"/>
						</a:lt2>
						<a:accent1>
							<a:srgbClr val="339933"/>
						</a:accent1>
						<a:accent2>
							<a:srgbClr val="800000"/>
						</a:accent2>
						<a:accent3>
							<a:srgbClr val="FFFFE2"/>
						</a:accent3>
						<a:accent4>
							<a:srgbClr val="000000"/>
						</a:accent4>
						<a:accent5>
							<a:srgbClr val="ADCAAD"/>
						</a:accent5>
						<a:accent6>
							<a:srgbClr val="730000"/>
						</a:accent6>
						<a:hlink>
							<a:srgbClr val="0033CC"/>
						</a:hlink>
						<a:folHlink>
							<a:srgbClr val="FFCC66"/>
						</a:folHlink>
					</a:clrScheme>
					<a:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/>
				</a:extraClrScheme>
				<a:extraClrScheme>
					<a:clrScheme name="Office Theme 4">
						<a:dk1>
							<a:srgbClr val="000000"/>
						</a:dk1>
						<a:lt1>
							<a:srgbClr val="FFFFFF"/>
						</a:lt1>
						<a:dk2>
							<a:srgbClr val="000000"/>
						</a:dk2>
						<a:lt2>
							<a:srgbClr val="333333"/>
						</a:lt2>
						<a:accent1>
							<a:srgbClr val="DDDDDD"/>
						</a:accent1>
						<a:accent2>
							<a:srgbClr val="808080"/>
						</a:accent2>
						<a:accent3>
							<a:srgbClr val="FFFFFF"/>
						</a:accent3>
						<a:accent4>
							<a:srgbClr val="000000"/>
						</a:accent4>
						<a:accent5>
							<a:srgbClr val="EBEBEB"/>
						</a:accent5>
						<a:accent6>
							<a:srgbClr val="737373"/>
						</a:accent6>
						<a:hlink>
							<a:srgbClr val="4D4D4D"/>
						</a:hlink>
						<a:folHlink>
							<a:srgbClr val="EAEAEA"/>
						</a:folHlink>
					</a:clrScheme>
					<a:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/>
				</a:extraClrScheme>
				<a:extraClrScheme>
					<a:clrScheme name="Office Theme 5">
						<a:dk1>
							<a:srgbClr val="000000"/>
						</a:dk1>
						<a:lt1>
							<a:srgbClr val="FFFFFF"/>
						</a:lt1>
						<a:dk2>
							<a:srgbClr val="000000"/>
						</a:dk2>
						<a:lt2>
							<a:srgbClr val="808080"/>
						</a:lt2>
						<a:accent1>
							<a:srgbClr val="FFCC66"/>
						</a:accent1>
						<a:accent2>
							<a:srgbClr val="0000FF"/>
						</a:accent2>
						<a:accent3>
							<a:srgbClr val="FFFFFF"/>
						</a:accent3>
						<a:accent4>
							<a:srgbClr val="000000"/>
						</a:accent4>
						<a:accent5>
							<a:srgbClr val="FFE2B8"/>
						</a:accent5>
						<a:accent6>
							<a:srgbClr val="0000E7"/>
						</a:accent6>
						<a:hlink>
							<a:srgbClr val="CC00CC"/>
						</a:hlink>
						<a:folHlink>
							<a:srgbClr val="C0C0C0"/>
						</a:folHlink>
					</a:clrScheme>
					<a:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/>
				</a:extraClrScheme>
				<a:extraClrScheme>
					<a:clrScheme name="Office Theme 6">
						<a:dk1>
							<a:srgbClr val="000000"/>
						</a:dk1>
						<a:lt1>
							<a:srgbClr val="FFFFFF"/>
						</a:lt1>
						<a:dk2>
							<a:srgbClr val="000000"/>
						</a:dk2>
						<a:lt2>
							<a:srgbClr val="808080"/>
						</a:lt2>
						<a:accent1>
							<a:srgbClr val="C0C0C0"/>
						</a:accent1>
						<a:accent2>
							<a:srgbClr val="0066FF"/>
						</a:accent2>
						<a:accent3>
							<a:srgbClr val="FFFFFF"/>
						</a:accent3>
						<a:accent4>
							<a:srgbClr val="000000"/>
						</a:accent4>
						<a:accent5>
							<a:srgbClr val="DCDCDC"/>
						</a:accent5>
						<a:accent6>
							<a:srgbClr val="005CE7"/>
						</a:accent6>
						<a:hlink>
							<a:srgbClr val="FF0000"/>
						</a:hlink>
						<a:folHlink>
							<a:srgbClr val="009900"/>
						</a:folHlink>
					</a:clrScheme>
					<a:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/>
				</a:extraClrScheme>
				<a:extraClrScheme>
					<a:clrScheme name="Office Theme 7">
						<a:dk1>
							<a:srgbClr val="000000"/>
						</a:dk1>
						<a:lt1>
							<a:srgbClr val="FFFFFF"/>
						</a:lt1>
						<a:dk2>
							<a:srgbClr val="000000"/>
						</a:dk2>
						<a:lt2>
							<a:srgbClr val="808080"/>
						</a:lt2>
						<a:accent1>
							<a:srgbClr val="3399FF"/>
						</a:accent1>
						<a:accent2>
							<a:srgbClr val="99FFCC"/>
						</a:accent2>
						<a:accent3>
							<a:srgbClr val="FFFFFF"/>
						</a:accent3>
						<a:accent4>
							<a:srgbClr val="000000"/>
						</a:accent4>
						<a:accent5>
							<a:srgbClr val="ADCAFF"/>
						</a:accent5>
						<a:accent6>
							<a:srgbClr val="8AE7B9"/>
						</a:accent6>
						<a:hlink>
							<a:srgbClr val="CC00CC"/>
						</a:hlink>
						<a:folHlink>
							<a:srgbClr val="B2B2B2"/>
						</a:folHlink>
					</a:clrScheme>
					<a:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/>
				</a:extraClrScheme>
			</a:extraClrSchemeLst>
		</a:theme>
	</xsl:template>

	<xsl:template name ="tmpTheme2">
		<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Office Theme">
			<a:themeElements>
				<a:clrScheme name="">
					<a:dk1>
						<a:srgbClr val="000000"/>
					</a:dk1>
					<a:lt1>
						<a:srgbClr val="FFFFFF"/>
					</a:lt1>
					<a:dk2>
						<a:srgbClr val="000000"/>
					</a:dk2>
					<a:lt2>
						<a:srgbClr val="808080"/>
					</a:lt2>
					<a:accent1>
						<a:srgbClr val="00CC99"/>
					</a:accent1>
					<a:accent2>
						<a:srgbClr val="3333CC"/>
					</a:accent2>
					<a:accent3>
						<a:srgbClr val="FFFFFF"/>
					</a:accent3>
					<a:accent4>
						<a:srgbClr val="000000"/>
					</a:accent4>
					<a:accent5>
						<a:srgbClr val="AAE2CA"/>
					</a:accent5>
					<a:accent6>
						<a:srgbClr val="2D2DB9"/>
					</a:accent6>
					<a:hlink>
						<a:srgbClr val="CCCCFF"/>
					</a:hlink>
					<a:folHlink>
						<a:srgbClr val="B2B2B2"/>
					</a:folHlink>
				</a:clrScheme>
				<a:fontScheme name="Office">
					<a:majorFont>
						<a:latin typeface="Calibri"/>
						<a:ea typeface=""/>
						<a:cs typeface=""/>
						<a:font script="Jpan" typeface="ＭＳ Ｐゴシック"/>
						<a:font script="Hang" typeface="맑은 고딕"/>
						<a:font script="Hans" typeface="宋体"/>
						<a:font script="Hant" typeface="新細明體"/>
						<a:font script="Arab" typeface="Times New Roman"/>
						<a:font script="Hebr" typeface="Times New Roman"/>
						<a:font script="Thai" typeface="Angsana New"/>
						<a:font script="Ethi" typeface="Nyala"/>
						<a:font script="Beng" typeface="Vrinda"/>
						<a:font script="Gujr" typeface="Shruti"/>
						<a:font script="Khmr" typeface="MoolBoran"/>
						<a:font script="Knda" typeface="Tunga"/>
						<a:font script="Guru" typeface="Raavi"/>
						<a:font script="Cans" typeface="Euphemia"/>
						<a:font script="Cher" typeface="Plantagenet Cherokee"/>
						<a:font script="Yiii" typeface="Microsoft Yi Baiti"/>
						<a:font script="Tibt" typeface="Microsoft Himalaya"/>
						<a:font script="Thaa" typeface="MV Boli"/>
						<a:font script="Deva" typeface="Mangal"/>
						<a:font script="Telu" typeface="Gautami"/>
						<a:font script="Taml" typeface="Latha"/>
						<a:font script="Syrc" typeface="Estrangelo Edessa"/>
						<a:font script="Orya" typeface="Kalinga"/>
						<a:font script="Mlym" typeface="Kartika"/>
						<a:font script="Laoo" typeface="DokChampa"/>
						<a:font script="Sinh" typeface="Iskoola Pota"/>
						<a:font script="Mong" typeface="Mongolian Baiti"/>
						<a:font script="Viet" typeface="Times New Roman"/>
						<a:font script="Uigh" typeface="Microsoft Uighur"/>
					</a:majorFont>
					<a:minorFont>
						<a:latin typeface="Calibri"/>
						<a:ea typeface=""/>
						<a:cs typeface=""/>
						<a:font script="Jpan" typeface="ＭＳ Ｐゴシック"/>
						<a:font script="Hang" typeface="맑은 고딕"/>
						<a:font script="Hans" typeface="宋体"/>
						<a:font script="Hant" typeface="新細明體"/>
						<a:font script="Arab" typeface="Arial"/>
						<a:font script="Hebr" typeface="Arial"/>
						<a:font script="Thai" typeface="Cordia New"/>
						<a:font script="Ethi" typeface="Nyala"/>
						<a:font script="Beng" typeface="Vrinda"/>
						<a:font script="Gujr" typeface="Shruti"/>
						<a:font script="Khmr" typeface="DaunPenh"/>
						<a:font script="Knda" typeface="Tunga"/>
						<a:font script="Guru" typeface="Raavi"/>
						<a:font script="Cans" typeface="Euphemia"/>
						<a:font script="Cher" typeface="Plantagenet Cherokee"/>
						<a:font script="Yiii" typeface="Microsoft Yi Baiti"/>
						<a:font script="Tibt" typeface="Microsoft Himalaya"/>
						<a:font script="Thaa" typeface="MV Boli"/>
						<a:font script="Deva" typeface="Mangal"/>
						<a:font script="Telu" typeface="Gautami"/>
						<a:font script="Taml" typeface="Latha"/>
						<a:font script="Syrc" typeface="Estrangelo Edessa"/>
						<a:font script="Orya" typeface="Kalinga"/>
						<a:font script="Mlym" typeface="Kartika"/>
						<a:font script="Laoo" typeface="DokChampa"/>
						<a:font script="Sinh" typeface="Iskoola Pota"/>
						<a:font script="Mong" typeface="Mongolian Baiti"/>
						<a:font script="Viet" typeface="Arial"/>
						<a:font script="Uigh" typeface="Microsoft Uighur"/>
					</a:minorFont>
				</a:fontScheme>
				<a:fmtScheme name="Office">
					<a:fillStyleLst>
						<a:solidFill>
							<a:schemeClr val="phClr"/>
						</a:solidFill>
						<a:gradFill rotWithShape="1">
							<a:gsLst>
								<a:gs pos="0">
									<a:schemeClr val="phClr">
										<a:tint val="50000"/>
										<a:satMod val="300000"/>
									</a:schemeClr>
								</a:gs>
								<a:gs pos="35000">
									<a:schemeClr val="phClr">
										<a:tint val="37000"/>
										<a:satMod val="300000"/>
									</a:schemeClr>
								</a:gs>
								<a:gs pos="100000">
									<a:schemeClr val="phClr">
										<a:tint val="15000"/>
										<a:satMod val="350000"/>
									</a:schemeClr>
								</a:gs>
							</a:gsLst>
							<a:lin ang="16200000" scaled="1"/>
						</a:gradFill>
						<a:gradFill rotWithShape="1">
							<a:gsLst>
								<a:gs pos="0">
									<a:schemeClr val="phClr">
										<a:shade val="51000"/>
										<a:satMod val="130000"/>
									</a:schemeClr>
								</a:gs>
								<a:gs pos="80000">
									<a:schemeClr val="phClr">
										<a:shade val="93000"/>
										<a:satMod val="130000"/>
									</a:schemeClr>
								</a:gs>
								<a:gs pos="100000">
									<a:schemeClr val="phClr">
										<a:shade val="94000"/>
										<a:satMod val="135000"/>
									</a:schemeClr>
								</a:gs>
							</a:gsLst>
							<a:lin ang="16200000" scaled="0"/>
						</a:gradFill>
					</a:fillStyleLst>
					<a:lnStyleLst>
						<a:ln w="9525" cap="flat" cmpd="sng" algn="ctr">
							<a:solidFill>
								<a:schemeClr val="phClr">
									<a:shade val="95000"/>
									<a:satMod val="105000"/>
								</a:schemeClr>
							</a:solidFill>
							<a:prstDash val="solid"/>
						</a:ln>
						<a:ln w="25400" cap="flat" cmpd="sng" algn="ctr">
							<a:solidFill>
								<a:schemeClr val="phClr"/>
							</a:solidFill>
							<a:prstDash val="solid"/>
						</a:ln>
						<a:ln w="38100" cap="flat" cmpd="sng" algn="ctr">
							<a:solidFill>
								<a:schemeClr val="phClr"/>
							</a:solidFill>
							<a:prstDash val="solid"/>
						</a:ln>
					</a:lnStyleLst>
					<a:effectStyleLst>
						<a:effectStyle>
							<a:effectLst>
								<a:outerShdw blurRad="40000" dist="20000" dir="5400000" rotWithShape="0">
									<a:srgbClr val="000000">
										<a:alpha val="38000"/>
									</a:srgbClr>
								</a:outerShdw>
							</a:effectLst>
						</a:effectStyle>
						<a:effectStyle>
							<a:effectLst>
								<a:outerShdw blurRad="40000" dist="23000" dir="5400000" rotWithShape="0">
									<a:srgbClr val="000000">
										<a:alpha val="35000"/>
									</a:srgbClr>
								</a:outerShdw>
							</a:effectLst>
						</a:effectStyle>
						<a:effectStyle>
							<a:effectLst>
								<a:outerShdw blurRad="40000" dist="23000" dir="5400000" rotWithShape="0">
									<a:srgbClr val="000000">
										<a:alpha val="35000"/>
									</a:srgbClr>
								</a:outerShdw>
							</a:effectLst>
							<a:scene3d>
								<a:camera prst="orthographicFront">
									<a:rot lat="0" lon="0" rev="0"/>
								</a:camera>
								<a:lightRig rig="threePt" dir="t">
									<a:rot lat="0" lon="0" rev="1200000"/>
								</a:lightRig>
							</a:scene3d>
							<a:sp3d>
								<a:bevelT w="63500" h="25400"/>
							</a:sp3d>
						</a:effectStyle>
					</a:effectStyleLst>
					<a:bgFillStyleLst>
						<a:solidFill>
							<a:schemeClr val="phClr"/>
						</a:solidFill>
						<a:gradFill rotWithShape="1">
							<a:gsLst>
								<a:gs pos="0">
									<a:schemeClr val="phClr">
										<a:tint val="40000"/>
										<a:satMod val="350000"/>
									</a:schemeClr>
								</a:gs>
								<a:gs pos="40000">
									<a:schemeClr val="phClr">
										<a:tint val="45000"/>
										<a:shade val="99000"/>
										<a:satMod val="350000"/>
									</a:schemeClr>
								</a:gs>
								<a:gs pos="100000">
									<a:schemeClr val="phClr">
										<a:shade val="20000"/>
										<a:satMod val="255000"/>
									</a:schemeClr>
								</a:gs>
							</a:gsLst>
							<a:path path="circle">
								<a:fillToRect l="50000" t="-80000" r="50000" b="180000"/>
							</a:path>
						</a:gradFill>
						<a:gradFill rotWithShape="1">
							<a:gsLst>
								<a:gs pos="0">
									<a:schemeClr val="phClr">
										<a:tint val="80000"/>
										<a:satMod val="300000"/>
									</a:schemeClr>
								</a:gs>
								<a:gs pos="100000">
									<a:schemeClr val="phClr">
										<a:shade val="30000"/>
										<a:satMod val="200000"/>
									</a:schemeClr>
								</a:gs>
							</a:gsLst>
							<a:path path="circle">
								<a:fillToRect l="50000" t="50000" r="50000" b="50000"/>
							</a:path>
						</a:gradFill>
					</a:bgFillStyleLst>
				</a:fmtScheme>
			</a:themeElements>
			<a:objectDefaults/>
			<a:extraClrSchemeLst/>
		</a:theme>
	</xsl:template>

	<xsl:template name ="tmpRel">
		<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
			<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
			<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/thumbnail" Target="docProps/thumbnail.jpeg"/>
			<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/>
			<Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
		</Relationships>
	</xsl:template>
	</xsl:stylesheet>
