<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<!--
Copyright (c) 2007, Sonata Software Limited
* All rights reserved.
*
* Redistribution and use in source and binary forms, with or without
* modification, are permitted provided that the following conditions are met:
*
*     * Redistributions of source code must retain the above copyright
*       notice, this list of conditions and the following disclaimer.
*     * Redistributions in binary form must reproduce the above copyright
*       notice, this list of conditions and the following disclaimer in the
*       documentation and/or other materials provided with the distribution.
*     * Neither the name of Sonata Software Limited nor the names of its contributors
*       may be used to endorse or promote products derived from this software
*       without specific prior written permission.
*
* THIS SOFTWARE IS PROVIDED BY THE REGENTS AND CONTRIBUTORS ``AS IS'' AND ANY
* EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE IMPLIED
* WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE
* DISCLAIMED. IN NO EVENT SHALL THE REGENTS AND CONTRIBUTORS BE LIABLE FOR ANY
* DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES
* (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES;
* LOSS OF USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND
* ON ANY THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT
* (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE OF THIS
* SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE
-->
<xsl:stylesheet version="1.0"
  xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
  xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
  xmlns:odf="urn:odf"
  xmlns:pzip="urn:cleverage:xmlns:post-processings:zip"
  xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0"
  xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"
  xmlns:number="urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0"
  xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0"
  xmlns:page="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0"
  exclude-result-prefixes="odf style text number draw page">

	<xsl:import href="docprops.xsl"/>
	<xsl:import href ="slides.xsl"/>
	<xsl:import href="presentation.xsl"/>
	<xsl:import href ="theme.xsl"/>
	<xsl:import href ="slideMasters.xsl"/>
	<xsl:import href="slideLayouts.xsl"/>

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
		<xsl:processing-instruction name="mso-application">progid="ppt.Document</xsl:processing-instruction>

		<pzip:archive pzip:target="{$outputFile}">

			<!-- Document core properties -->
			<pzip:entry pzip:target="docProps/core.xml">
				<xsl:call-template name="docprops-core"/>
			</pzip:entry>
			<!--Document app properties-->
			<pzip:entry pzip:target="docProps/app.xml">
				<xsl:call-template name="docprops-app"/>
			</pzip:entry>
			<pzip:entry pzip:target="ppt/presProps.xml">
				<xsl:call-template name="presProps"/>
			</pzip:entry>
			<pzip:entry pzip:target="ppt/tableStyles.xml">
				<xsl:call-template name="tabStyles"/>
			</pzip:entry>
			<pzip:entry pzip:target="ppt/viewProps.xml">
				<xsl:call-template name="InsetViewPropsDotXML"/>
			</pzip:entry>
			<pzip:entry pzip:target="ppt/presentation.xml">
				<xsl:call-template name="presentation"/>
			</pzip:entry>
			<pzip:entry pzip:target="ppt/_rels/presentation.xml.rels">
				<xsl:call-template name="presentationRel"/>
			</pzip:entry>
			<pzip:entry pzip:target="[Content_Types].xml">
				<xsl:call-template name="content"/>
			</pzip:entry>
			<xsl:for-each  select ="document('styles.xml')//style:master-page">
				<xsl:variable name ="SlideMaster">
					<xsl:value-of select ="concat('slideMaster',position())"/>
				</xsl:variable>
				<pzip:entry pzip:target="{concat('ppt/theme/theme',position(),'.xml')}">
					<xsl:call-template name="theme"/>
				</pzip:entry>
				<pzip:entry pzip:target="{concat('ppt/slideMasters/',$SlideMaster,'.xml')}">
					<!-- Check the below template name -->
					<xsl:call-template name="slideMasters">
						<xsl:with-param name="slideMasterName" select="@style:name"/>
						<xsl:with-param name ="smId" select ="position()"/>
					</xsl:call-template>
				</pzip:entry>
				<pzip:entry pzip:target="{concat('ppt/slideMasters/_rels/',$SlideMaster,'.xml.rels')}">
					<xsl:call-template name="slideMaster1Rel">
						<xsl:with-param name ="StartLayoutNo" select ="(position() + (10 *(position() - 1 )))"/>
						<xsl:with-param name ="ThemeId" select ="position()"/>
					</xsl:call-template >
				</pzip:entry>
				<xsl:call-template name ="CreateLaouts">
					<xsl:with-param  name ="LayOutNo">
						<xsl:value-of select ="(position() + (10 *(position() - 1 )))"/>
					</xsl:with-param>
					<xsl:with-param name ="SlideMasterNo" select ="position()"/>
          <xsl:with-param name="slideMasterName" select="@style:name"/>
				</xsl:call-template>
			</xsl:for-each>

			<xsl:for-each select ="document('content.xml')
				/office:document-content/office:body/office:presentation/draw:page">
				<xsl:variable name ="SlideName">
					<xsl:value-of select ="concat(concat('Slide',position()),'.xml')"/>
				</xsl:variable>
				<pzip:entry pzip:target="{concat('ppt/slides/',$SlideName)}">
					<!--<xsl:call-template name="slides" />-->
					<xsl:apply-templates select ="." mode="slide" >
						<xsl:with-param name ="pageNo" select ="position()"/>
					</xsl:apply-templates >
				</pzip:entry>
				<pzip:entry pzip:target="{concat(concat('ppt/slides/_rels/',$SlideName), '.rels')}">
					<xsl:call-template name="slidesRel"/>
				</pzip:entry>
			</xsl:for-each>
			<!--<pzip:entry pzip:target="ppt/slides/_rels/slide1.xml.rels">
				<xsl:call-template name="slidesRel"/>
			</pzip:entry>-->
			<pzip:entry pzip:target="_rels/.rels">
				<xsl:call-template name="Rels"/>
			</pzip:entry>
		</pzip:archive>
	</xsl:template>
	<xsl:template name ="CreateLaouts">
		<xsl:param name ="LayOutNo"/>
		<xsl:param name ="SlideMasterNo"/>
    <xsl:param name="slideMasterName"/>
    
		<pzip:entry pzip:target="{concat('ppt/slideLayouts/slideLayout',$LayOutNo,'.xml')}">
      <xsl:call-template name="InsertSlideLayout1">
        <xsl:with-param name ="MasterName" select ="$slideMasterName"/>
      </xsl:call-template >
		</pzip:entry>
		<pzip:entry pzip:target="{concat('ppt/slideLayouts/slideLayout',$LayOutNo +1 ,'.xml')}">
      <xsl:call-template name="InsertSlideLayout2">
        <xsl:with-param name ="MasterName" select ="$slideMasterName"/>
      </xsl:call-template>
		</pzip:entry>
		<pzip:entry pzip:target="{concat('ppt/slideLayouts/slideLayout',$LayOutNo +2 ,'.xml')}">
			<xsl:call-template name="InsertSlideLayout3">
        <xsl:with-param name ="MasterName" select ="$slideMasterName"/>
      </xsl:call-template>
		</pzip:entry>
		<pzip:entry pzip:target="{concat('ppt/slideLayouts/slideLayout',$LayOutNo +3 ,'.xml')}">
			<xsl:call-template name="InsertSlideLayout4">
        <xsl:with-param name ="MasterName" select ="$slideMasterName"/>
      </xsl:call-template>
		</pzip:entry>
		<pzip:entry pzip:target="{concat('ppt/slideLayouts/slideLayout',$LayOutNo +4 ,'.xml')}">
			<xsl:call-template name="InsertSlideLayout5">
        <xsl:with-param name ="MasterName" select ="$slideMasterName"/>
      </xsl:call-template>
		</pzip:entry>
		<pzip:entry pzip:target="{concat('ppt/slideLayouts/slideLayout',$LayOutNo +5 ,'.xml')}">
			<xsl:call-template name="InsertSlideLayout6">
        <xsl:with-param name ="MasterName" select ="$slideMasterName"/>
      </xsl:call-template>
		</pzip:entry>
		<pzip:entry pzip:target="{concat('ppt/slideLayouts/slideLayout',$LayOutNo +6 ,'.xml')}">
			<xsl:call-template name="InsertSlideLayout7">
        <xsl:with-param name ="MasterName" select ="$slideMasterName"/>
      </xsl:call-template>
		</pzip:entry>
		<pzip:entry pzip:target="{concat('ppt/slideLayouts/slideLayout',$LayOutNo +7 ,'.xml')}">
			<xsl:call-template name="InsertSlideLayout8">
        <xsl:with-param name ="MasterName" select ="$slideMasterName"/>
      </xsl:call-template>
		</pzip:entry>
		<pzip:entry pzip:target="{concat('ppt/slideLayouts/slideLayout',$LayOutNo +8 ,'.xml')}">
			<xsl:call-template name="InsertSlideLayout9">
        <xsl:with-param name ="MasterName" select ="$slideMasterName"/>
      </xsl:call-template>
		</pzip:entry>
		<pzip:entry pzip:target="{concat('ppt/slideLayouts/slideLayout',$LayOutNo +9 ,'.xml')}">
			<xsl:call-template name="InsertSlideLayout10">
        <xsl:with-param name ="MasterName" select ="$slideMasterName"/>
      </xsl:call-template>
		</pzip:entry>
		<pzip:entry pzip:target="{concat('ppt/slideLayouts/slideLayout',$LayOutNo +10 ,'.xml')}">
			<xsl:call-template name="InsertSlideLayout11">
        <xsl:with-param name ="MasterName" select ="$slideMasterName"/>
      </xsl:call-template>
		</pzip:entry>
		<!-- For layouts -->
		<pzip:entry pzip:target="{concat('ppt/slideLayouts/_rels/slideLayout',$LayOutNo ,'.xml.rels')}">
			<xsl:call-template name="InsertLayoutRel">
				<xsl:with-param name ="slideMaasterNo" select ="$SlideMasterNo"/>
			</xsl:call-template >
		</pzip:entry>
		<pzip:entry pzip:target="{concat('ppt/slideLayouts/_rels/slideLayout',$LayOutNo +1 ,'.xml.rels')}">
			<xsl:call-template name="InsertLayoutRel">
				<xsl:with-param name ="slideMaasterNo" select ="$SlideMasterNo"/>
			</xsl:call-template >
		</pzip:entry>
		<pzip:entry pzip:target="{concat('ppt/slideLayouts/_rels/slideLayout',$LayOutNo +2 ,'.xml.rels')}">
			<xsl:call-template name="InsertLayoutRel">
				<xsl:with-param name ="slideMaasterNo" select ="$SlideMasterNo"/>
			</xsl:call-template >
		</pzip:entry>
		<pzip:entry pzip:target="{concat('ppt/slideLayouts/_rels/slideLayout',$LayOutNo +3,'.xml.rels')}">
			<xsl:call-template name="InsertLayoutRel">
				<xsl:with-param name ="slideMaasterNo" select ="$SlideMasterNo"/>
			</xsl:call-template >
		</pzip:entry>
		<pzip:entry pzip:target="{concat('ppt/slideLayouts/_rels/slideLayout',$LayOutNo +4,'.xml.rels')}">
			<xsl:call-template name="InsertLayoutRel">
				<xsl:with-param name ="slideMaasterNo" select ="$SlideMasterNo"/>
			</xsl:call-template >
		</pzip:entry>
		<pzip:entry pzip:target="{concat('ppt/slideLayouts/_rels/slideLayout',$LayOutNo+5 ,'.xml.rels')}">
			<xsl:call-template name="InsertLayoutRel">
				<xsl:with-param name ="slideMaasterNo" select ="$SlideMasterNo"/>
			</xsl:call-template >
		</pzip:entry>
		<pzip:entry pzip:target="{concat('ppt/slideLayouts/_rels/slideLayout',$LayOutNo+6 ,'.xml.rels')}">
			<xsl:call-template name="InsertLayoutRel">
				<xsl:with-param name ="slideMaasterNo" select ="$SlideMasterNo"/>
			</xsl:call-template >
		</pzip:entry>
		<pzip:entry pzip:target="{concat('ppt/slideLayouts/_rels/slideLayout',$LayOutNo +7 ,'.xml.rels')}">
			<xsl:call-template name="InsertLayoutRel">
				<xsl:with-param name ="slideMaasterNo" select ="$SlideMasterNo"/>
			</xsl:call-template >
		</pzip:entry>
		<pzip:entry pzip:target="{concat('ppt/slideLayouts/_rels/slideLayout',$LayOutNo +8,'.xml.rels')}">
			<xsl:call-template name="InsertLayoutRel">
				<xsl:with-param name ="slideMaasterNo" select ="$SlideMasterNo"/>
			</xsl:call-template >
		</pzip:entry>
		<pzip:entry pzip:target="{concat('ppt/slideLayouts/_rels/slideLayout',$LayOutNo+9 ,'.xml.rels')}">
			<xsl:call-template name="InsertLayoutRel">
				<xsl:with-param name ="slideMaasterNo" select ="$SlideMasterNo"/>
			</xsl:call-template >
		</pzip:entry>
		<pzip:entry pzip:target="{concat('ppt/slideLayouts/_rels/slideLayout',$LayOutNo+10 ,'.xml.rels')}">
			<xsl:call-template name="InsertLayoutRel">
				<xsl:with-param name ="slideMaasterNo" select ="$SlideMasterNo"/>
			</xsl:call-template >
		</pzip:entry>
	</xsl:template>
	<xsl:template name ="content" >
		<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
			<!-- Added By Vijayeta,Extensions of the images,to be added in package-->
			<Default Extension="jpeg" ContentType="image/jpeg"/>
			<Default Extension="jpg" ContentType="image/jpeg"/>
			<Default Extension="gif" ContentType="image/gif"/>
			<Default Extension="png" ContentType="image/png"/>
			<!-- Added By Vijayeta,Extensions of the images,to be added in package-->
			<!--<Override PartName="/ppt/slideMasters/slideMaster1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"/>-->

			<!-- @@Slide master code begins Pradeep Nemadi-->
			<xsl:for-each select ="document('styles.xml')//office:master-styles/style:master-page ">
				<Override>
					<xsl:attribute name ="PartName">
						<xsl:value-of select ="concat('/ppt/theme/theme',position(),'.xml')"/>
					</xsl:attribute>
					<xsl:attribute name ="ContentType">
						<xsl:value-of select ="'application/vnd.openxmlformats-officedocument.theme+xml'"/>
					</xsl:attribute>
				</Override >
				<Override>
					<xsl:attribute name ="PartName">
						<xsl:value-of select ="concat('/ppt/slideMasters/slideMaster',position(),'.xml')"/>
					</xsl:attribute>
					<xsl:attribute name ="ContentType">
						<xsl:value-of select ="'application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml'"/>
					</xsl:attribute>
				</Override>
				<xsl:variable name ="LayoutStartNo">
					<xsl:value-of select ="(position() + (10 *(position() - 1 )))"/>
				</xsl:variable >
				<Override>
					<xsl:attribute name ="PartName">
						<xsl:value-of select ="concat('/ppt/slideLayouts/slideLayout',$LayoutStartNo,'.xml')"/>
					</xsl:attribute>
					<xsl:attribute name ="ContentType" >
						<xsl:value-of select ="'application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml'"/>
					</xsl:attribute>
				</Override >
				<Override>
					<xsl:attribute name ="PartName">
						<xsl:value-of select ="concat('/ppt/slideLayouts/slideLayout',$LayoutStartNo+1,'.xml')"/>
					</xsl:attribute>
					<xsl:attribute name ="ContentType" >
						<xsl:value-of select ="'application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml'"/>
					</xsl:attribute>
				</Override >

				<Override>
					<xsl:attribute name ="PartName">
						<xsl:value-of select ="concat('/ppt/slideLayouts/slideLayout',$LayoutStartNo+2,'.xml')"/>
					</xsl:attribute>
					<xsl:attribute name ="ContentType" >
						<xsl:value-of select ="'application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml'"/>
					</xsl:attribute>
				</Override >

				<Override>
					<xsl:attribute name ="PartName">
						<xsl:value-of select ="concat('/ppt/slideLayouts/slideLayout',$LayoutStartNo+3,'.xml')"/>
					</xsl:attribute>
					<xsl:attribute name ="ContentType" >
						<xsl:value-of select ="'application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml'"/>
					</xsl:attribute>
				</Override >

				<Override>
					<xsl:attribute name ="PartName">
						<xsl:value-of select ="concat('/ppt/slideLayouts/slideLayout',$LayoutStartNo+4,'.xml')"/>
					</xsl:attribute>
					<xsl:attribute name ="ContentType" >
						<xsl:value-of select ="'application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml'"/>
					</xsl:attribute>
				</Override >
				<Override>
					<xsl:attribute name ="PartName">
						<xsl:value-of select ="concat('/ppt/slideLayouts/slideLayout',$LayoutStartNo+5,'.xml')"/>
					</xsl:attribute>
					<xsl:attribute name ="ContentType" >
						<xsl:value-of select ="'application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml'"/>
					</xsl:attribute>
				</Override >
				<Override>
					<xsl:attribute name ="PartName">
						<xsl:value-of select ="concat('/ppt/slideLayouts/slideLayout',$LayoutStartNo+6,'.xml')"/>
					</xsl:attribute>
					<xsl:attribute name ="ContentType" >
						<xsl:value-of select ="'application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml'"/>
					</xsl:attribute>
				</Override >
				<Override>
					<xsl:attribute name ="PartName">
						<xsl:value-of select ="concat('/ppt/slideLayouts/slideLayout',$LayoutStartNo+7,'.xml')"/>
					</xsl:attribute>
					<xsl:attribute name ="ContentType" >
						<xsl:value-of select ="'application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml'"/>
					</xsl:attribute>
				</Override >
				<Override>
					<xsl:attribute name ="PartName">
						<xsl:value-of select ="concat('/ppt/slideLayouts/slideLayout',$LayoutStartNo+8,'.xml')"/>
					</xsl:attribute>
					<xsl:attribute name ="ContentType" >
						<xsl:value-of select ="'application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml'"/>
					</xsl:attribute>
				</Override >
				<Override>
					<xsl:attribute name ="PartName">
						<xsl:value-of select ="concat('/ppt/slideLayouts/slideLayout',$LayoutStartNo+9,'.xml')"/>
					</xsl:attribute>
					<xsl:attribute name ="ContentType" >
						<xsl:value-of select ="'application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml'"/>
					</xsl:attribute>
				</Override >
				<Override>
					<xsl:attribute name ="PartName">
						<xsl:value-of select ="concat('/ppt/slideLayouts/slideLayout',$LayoutStartNo+10,'.xml')"/>
					</xsl:attribute>
					<xsl:attribute name ="ContentType" >
						<xsl:value-of select ="'application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml'"/>
					</xsl:attribute>
				</Override >

			</xsl:for-each>
			<!-- @@Slide master code ends Pradeep Nemadi-->





			<!-- @@ write link to slide layouts end-->

			<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
			<Default Extension="xml" ContentType="application/xml"/>
			<Default Extension="wav" ContentType="audio/wav"/>
			<Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>
			<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
			<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
			<Override PartName="/ppt/tableStyles.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.tableStyles+xml"/>
			<Override PartName="/ppt/viewProps.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.viewProps+xml"/>
			<Override PartName="/ppt/presProps.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presProps+xml"/>

			<xsl:for-each select ="document('content.xml')
			/office:document-content/office:body/office:presentation/draw:page">
				<Override >
					<xsl:attribute name ="PartName">
						<xsl:value-of select ="concat(concat('/ppt/slides/slide',position()),'.xml')"/>
					</xsl:attribute>
					<xsl:attribute name ="ContentType">
						<xsl:value-of select ="'application/vnd.openxmlformats-officedocument.presentationml.slide+xml'"/>
					</xsl:attribute>
				</Override>
			</xsl:for-each>
		</Types>
	</xsl:template>
	<xsl:template name ="Rels">
		<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
			<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
			<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
			<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/>
		</Relationships>
	</xsl:template>
	<!-- InsetViewPropsDotXML Added by Lohith- Temp fix -->
	<xsl:template name="InsetViewPropsDotXML">
		<p:viewPr xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
			<p:normalViewPr showOutlineIcons="0">
				<p:restoredLeft sz="15620"/>
				<p:restoredTop sz="94660"/>
			</p:normalViewPr>
			<p:slideViewPr>
				<p:cSldViewPr>
					<p:cViewPr varScale="1">
						<p:scale>
							<a:sx n="59" d="100"/>
							<a:sy n="59" d="100"/>
						</p:scale>
						<p:origin x="-780" y="-78"/>
					</p:cViewPr>
					<p:guideLst>
						<p:guide orient="horz" pos="2160"/>
						<p:guide pos="2880"/>
					</p:guideLst>
				</p:cSldViewPr>
			</p:slideViewPr>
			<p:notesTextViewPr>
				<p:cViewPr>
					<p:scale>
						<a:sx n="100" d="100"/>
						<a:sy n="100" d="100"/>
					</p:scale>
					<p:origin x="0" y="0"/>
				</p:cViewPr>
			</p:notesTextViewPr>
			<p:gridSpacing cx="78028800" cy="78028800"/>
		</p:viewPr>

	</xsl:template>
</xsl:stylesheet>