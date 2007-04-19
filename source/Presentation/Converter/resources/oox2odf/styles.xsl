<?xml version="1.0" encoding="UTF-8" ?>
<!--
   Pradeep Nemadi
-->
<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
	xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" 
  xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" 
  xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" 
  xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" 
  xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" 
  xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0" 
  xmlns:xlink="http://www.w3.org/1999/xlink" 
  xmlns:dc="http://purl.org/dc/elements/1.1/" 
  xmlns:meta="urn:oasis:names:tc:opendocument:xmlns:meta:1.0" 
  xmlns:number="urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0" 
  xmlns:presentation="urn:oasis:names:tc:opendocument:xmlns:presentation:1.0" 
  xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0" 
  xmlns:chart="urn:oasis:names:tc:opendocument:xmlns:chart:1.0" 
  xmlns:dr3d="urn:oasis:names:tc:opendocument:xmlns:dr3d:1.0" 
  xmlns:math="http://www.w3.org/1998/Math/MathML" 
  xmlns:form="urn:oasis:names:tc:opendocument:xmlns:form:1.0" 
  xmlns:script="urn:oasis:names:tc:opendocument:xmlns:script:1.0" 
  xmlns:ooo="http://openoffice.org/2004/office" 
  xmlns:ooow="http://openoffice.org/2004/writer" 
  xmlns:oooc="http://openoffice.org/2004/calc" 
  xmlns:dom="http://www.w3.org/2001/xml-events" 
  xmlns:smil="urn:oasis:names:tc:opendocument:xmlns:smil-compatible:1.0" 
  xmlns:anim="urn:oasis:names:tc:opendocument:xmlns:animation:1.0" 
  xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" 
  xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" 
  xmlns:rels="http://schemas.openxmlformats.org/package/2006/relationships" >


	<!--<xsl:key name="StyleId" match="w:style" use="@w:styleId"/>
  <xsl:key name="default-styles"
    match="w:style[@w:default = 1 or @w:default = 'true' or @w:default = 'on']" use="@w:type"/>-->

	<xsl:template name="styles">
		<office:document-styles>
			<office:font-face-decls>
			</office:font-face-decls>
			<xsl:text>document styles</xsl:text>
			<office:styles>
				<xsl:call-template name="InsertShapeStyles"/>
			</office:styles>
			<xsl:text>automatic styles</xsl:text>
			<office:automatic-styles>
				<xsl:call-template name="InsertSlideSize"/>
				<xsl:call-template name="InsertNotesSize"/>
				<style:style style:name="pr1" style:family="presentation" style:parent-style-name="Default-backgroundobjects">
					<style:graphic-properties draw:stroke="none" draw:fill="none" draw:fill-color="#ffffff" draw:auto-grow-height="false" fo:min-height="1.449cm"/>
				</style:style>
			</office:automatic-styles>
			<xsl:text>master styles</xsl:text>
			<office:master-styles>
				<xsl:call-template name="InsertMasterStylesDefinition"/>
			</office:master-styles>
		</office:document-styles>

	</xsl:template>
	<xsl:template name="InsertShapeStyles">

		<xsl:variable name ="triangle" select ="1" />
		<xsl:for-each select ="document('ppt/presentation.xml')/p:presentation/p:sldIdLst/p:sldId">
			
			<xsl:variable name ="SlideId">
				<xsl:value-of  select ="concat(concat('slide',position()),'.xml')" />
			</xsl:variable>
			<xsl:for-each select ="document(concat('ppt/slides/',$SlideId))/p:sld/p:cSld/p:spTree/p:cxnSp/p:spPr/a:ln">
				
				<!--Dash types-->
				<xsl:call-template name ="getDashType">
					<xsl:with-param name ="val" select ="a:prstDash/@val" />
					<xsl:with-param name ="cap" select ="@cap" />
				</xsl:call-template>
				
				<!-- Head End-->
				<xsl:if test ="a:headEnd">
					<xsl:call-template name ="getArrowType">
						<xsl:with-param name ="type" select ="a:headEnd/@type" />
						<xsl:with-param name="w" select ="a:headEnd/@w" />
						<xsl:with-param name ="len" select ="a:headEnd/@len" />
					</xsl:call-template>
				</xsl:if>

				<!-- Tail End-->
				<xsl:if test ="a:tailEnd">
					<xsl:call-template name ="getArrowType">
						<xsl:with-param name ="type" select ="a:tailEnd/@type" />
						<xsl:with-param name="w" select ="a:tailEnd/@w" />
						<xsl:with-param name ="len" select ="a:tailEnd/@len" />
					</xsl:call-template>
				</xsl:if>
			</xsl:for-each >

			<!--Dash types-->
			<xsl:for-each select ="document(concat('ppt/slides/',$SlideId))/p:sld/p:cSld/p:spTree/p:sp/p:spPr/a:ln">
				<xsl:call-template name ="getDashType">
					<xsl:with-param name ="val" select ="a:prstDash/@val" />
					<xsl:with-param name ="cap" select ="@cap" />
				</xsl:call-template>
			</xsl:for-each>

			 
		</xsl:for-each >
		
	</xsl:template>

	<xsl:template name ="getArrowType">
		<xsl:param name ="type" />
		<xsl:param name ="w" />
		<xsl:param name="len" />

		<xsl:choose>

			<!--Triangle-->
			<xsl:when test ="($type='triangle')">
				<xsl:choose>
					<xsl:when test ="($w='sm') and ($len='sm')">
						<draw:marker draw:name="msArrowEnd1" draw:display-name="msArrowEnd 1" svg:viewBox="0 0 140 140" svg:d="m70 0 70 140h-140z"/>
					</xsl:when>
					<xsl:when test ="($w='sm') and ($len='med')">
						<draw:marker draw:name="msArrowEnd2" draw:display-name="msArrowEnd 2" svg:viewBox="0 0 140 210" svg:d="m70 0 70 210h-140z"/>
					</xsl:when>
					<xsl:when test ="($w='sm') and ($len='lg')">
						<draw:marker draw:name="msArrowEnd3" draw:display-name="msArrowEnd 3" svg:viewBox="0 0 140 350" svg:d="m70 0 70 350h-140z"/>
					</xsl:when>
					<xsl:when test ="($w='med') and ($len='sm')">
						<draw:marker draw:name="msArrowEnd4" draw:display-name="msArrowEnd 4" svg:viewBox="0 0 210 140" svg:d="m105 0 105 140h-210z"/>
					</xsl:when>
					<xsl:when test ="($w='med') and ($len='lg')">
						<draw:marker draw:name="msArrowEnd6" draw:display-name="msArrowEnd 6" svg:viewBox="0 0 210 350" svg:d="m105 0 105 350h-210z"/>
					</xsl:when>
					<xsl:when test ="($w='lg') and ($len='sm')">
						<draw:marker draw:name="msArrowEnd7" draw:display-name="msArrowEnd 7" svg:viewBox="0 0 350 140" svg:d="m175 0 175 140h-350z"/>
					</xsl:when>
					<xsl:when test ="($w='lg') and ($len='med')">
						<draw:marker draw:name="msArrowEnd8" draw:display-name="msArrowEnd 8" svg:viewBox="0 0 350 210" svg:d="m175 0 175 210h-350z"/>
					</xsl:when>
					<xsl:when test ="($w='lg') and ($len='lg')">
						<draw:marker draw:name="msArrowEnd9" draw:display-name="msArrowEnd 9" svg:viewBox="0 0 350 350" svg:d="m175 0 175 350h-350z"/>
					</xsl:when>
					<xsl:otherwise>
						<draw:marker draw:name="msArrowEnd5" draw:display-name="msArrowEnd 5" svg:viewBox="0 0 210 210" svg:d="m105 0 105 210h-210z"/>
					</xsl:otherwise>
				</xsl:choose>
			</xsl:when>

			<!--Arrow-->
			<xsl:when test ="($type='arrow')">
				<xsl:choose>
					<xsl:when test ="($w='sm') and ($len='sm')">
						<draw:marker draw:name="msArrowOpenEnd1" draw:display-name="msArrowOpenEnd 1" svg:viewBox="0 0 245 245" svg:d="m122 0 123 222-37 23-86-157-86 157-36-23z"/>
					</xsl:when>
					<xsl:when test ="($w='sm') and ($len='med')">
						<draw:marker draw:name="msArrowOpenEnd2" draw:display-name="msArrowOpenEnd 2" svg:viewBox="0 0 245 315" svg:d="m122 0 123 286-37 29-86-202-86 202-36-29z"/>
					</xsl:when>
					<xsl:when test ="($w='sm') and ($len='lg')">
						<draw:marker draw:name="msArrowOpenEnd3" draw:display-name="msArrowOpenEnd 3" svg:viewBox="0 0 245 420" svg:d="m122 0 123 382-37 38-86-269-86 269-36-38z"/>
					</xsl:when>
					<xsl:when test ="($w='med') and ($len='sm')">
						<draw:marker draw:name="msArrowOpenEnd4" draw:display-name="msArrowOpenEnd 4" svg:viewBox="0 0 315 245" svg:d="m157 0 158 222-48 23-110-157-110 157-47-23z"/>
					</xsl:when>
					<xsl:when test ="($w='med') and ($len='lg')">
						<draw:marker draw:name="msArrowOpenEnd6" draw:display-name="msArrowOpenEnd 6" svg:viewBox="0 0 315 420" svg:d="m157 0 158 382-48 38-110-269-110 269-47-38z"/>
					</xsl:when>
					<xsl:when test ="($w='lg') and ($len='sm')">
						<draw:marker draw:name="msArrowOpenEnd7" draw:display-name="msArrowOpenEnd 7" svg:viewBox="0 0 420 245" svg:d="m210 0 210 222-63 23-147-157-147 157-63-23z"/>
					</xsl:when>
					<xsl:when test ="($w='lg') and ($len='med')">
						<draw:marker draw:name="msArrowOpenEnd8" draw:display-name="msArrowOpenEnd 8" svg:viewBox="0 0 420 315" svg:d="m210 0 210 286-63 29-147-202-147 202-63-29z"/>
					</xsl:when>
					<xsl:when test ="($w='lg') and ($len='lg')">
						<draw:marker draw:name="msArrowOpenEnd9" draw:display-name="msArrowOpenEnd 9" svg:viewBox="0 0 420 420" svg:d="m210 0 210 382-63 38-147-269-147 269-63-38z"/>
					</xsl:when>
					<xsl:otherwise>
						<draw:marker draw:name="msArrowOpenEnd5" draw:display-name="msArrowOpenEnd 5" svg:viewBox="0 0 315 315" svg:d="m157 0 158 286-48 29-110-202-110 202-47-29z"/>
					</xsl:otherwise>
				</xsl:choose>
			</xsl:when>

			<!--Stealth-->
			<xsl:when test ="($type='stealth')">
				<xsl:choose>
					<xsl:when test ="($w='sm') and ($len='sm')">
						<draw:marker draw:name="msArrowStealthEnd1" draw:display-name="msArrowStealthEnd 1" svg:viewBox="0 0 140 140" svg:d="m70 0 70 140-70-56-70 56z"/>
					</xsl:when>
					<xsl:when test ="($w='sm') and ($len='med')">
						<draw:marker draw:name="msArrowStealthEnd2" draw:display-name="msArrowStealthEnd 2" svg:viewBox="0 0 140 210" svg:d="m70 0 70 210-70-84-70 84z"/>
					</xsl:when>
					<xsl:when test ="($w='sm') and ($len='lg')">
						<draw:marker draw:name="msArrowStealthEnd3" draw:display-name="msArrowStealthEnd 3" svg:viewBox="0 0 140 350" svg:d="m70 0 70 350-70-140-70 140z"/>
					</xsl:when>
					<xsl:when test ="($w='med') and ($len='sm')">
						<draw:marker draw:name="msArrowStealthEnd4" draw:display-name="msArrowStealthEnd 4" svg:viewBox="0 0 210 140" svg:d="m105 0 105 140-105-56-105 56z"/>
					</xsl:when>
					<xsl:when test ="($w='med') and ($len='lg')">
						<draw:marker draw:name="msArrowStealthEnd6" draw:display-name="msArrowStealthEnd 6" svg:viewBox="0 0 210 350" svg:d="m105 0 105 350-105-140-105 140z"/>
					</xsl:when>
					<xsl:when test="($w='lg') and ($len='sm')">
						<draw:marker draw:name="msArrowStealthEnd7" draw:display-name="msArrowStealthEnd 7" svg:viewBox="0 0 350 140" svg:d="m175 0 175 140-175-56-175 56z"/>
					</xsl:when>
					<xsl:when test ="($w='lg') and ($len='med')">
						<draw:marker draw:name="msArrowStealthEnd8" draw:display-name="msArrowStealthEnd 8" svg:viewBox="0 0 350 210" svg:d="m175 0 175 210-175-84-175 84z"/>
					</xsl:when>
					<xsl:when test ="($w='lg') and ($len='lg')">
						<draw:marker draw:name="msArrowStealthEnd9" draw:display-name="msArrowStealthEnd 9" svg:viewBox="0 0 350 350" svg:d="m175 0 175 350-175-140-175 140z"/>
					</xsl:when>
					<xsl:otherwise>
						<draw:marker draw:name="msArrowStealthEnd5" draw:display-name="msArrowStealthEnd 5" svg:viewBox="0 0 210 210" svg:d="m105 0 105 210-105-84-105 84z"/>
					</xsl:otherwise> 
				</xsl:choose>
			</xsl:when>

			<!--Oval-->
			<xsl:when test ="($type='oval')">
				<xsl:choose>
					<xsl:when test ="($w='sm') and ($len='sm')">
						<draw:marker draw:name="msArrowOvalEnd1" draw:display-name="msArrowOvalEnd 1" svg:viewBox="0 0 140 140" svg:d="m140 0c0-38-32-70-70-70-38 0-70 32-70 70 0 38 32 70 70 70 38 0 70-32 70-70z"/>
					</xsl:when>
					<xsl:when test ="($w='sm') and ($len='med')">
						<draw:marker draw:name="msArrowOvalEnd2" draw:display-name="msArrowOvalEnd 2" svg:viewBox="0 0 140 210" svg:d="m140 0c0-57-32-105-70-105-38 0-70 48-70 105 0 57 32 105 70 105 38 0 70-48 70-105z"/>
					</xsl:when>
					<xsl:when test ="($w='sm') and ($len='lg')">
						<draw:marker draw:name="msArrowOvalEnd3" draw:display-name="msArrowOvalEnd 3" svg:viewBox="0 0 140 350" svg:d="m140 0c0-96-32-175-70-175-38 0-70 79-70 175 0 96 32 175 70 175 38 0 70-79 70-175z"/>
					</xsl:when>
					<xsl:when test ="($w='med') and ($len='sm')">
						<draw:marker draw:name="msArrowOvalEnd4" draw:display-name="msArrowOvalEnd 4" svg:viewBox="0 0 210 140" svg:d="m210 0c0-38-48-70-105-70-57 0-105 32-105 70 0 38 48 70 105 70 57 0 105-32 105-70z"/>
					</xsl:when>
					<xsl:when test ="($w='med') and ($len='lg')">
						<draw:marker draw:name="msArrowOvalEnd6" draw:display-name="msArrowOvalEnd 6" svg:viewBox="0 0 210 350" svg:d="m210 0c0-96-48-175-105-175-57 0-105 79-105 175 0 96 48 175 105 175 57 0 105-79 105-175z"/>
					</xsl:when>
					<xsl:when test ="($w='lg') and ($len='sm')">
						<draw:marker draw:name="msArrowOvalEnd7" draw:display-name="msArrowOvalEnd 7" svg:viewBox="0 0 350 140" svg:d="m350 0c0-38-79-70-175-70-96 0-175 32-175 70 0 38 79 70 175 70 96 0 175-32 175-70z"/>
					</xsl:when>
					<xsl:when test ="($w='lg') and ($len='med')">
						<draw:marker draw:name="msArrowOvalEnd8" draw:display-name="msArrowOvalEnd 8" svg:viewBox="0 0 350 210" svg:d="m350 0c0-57-79-105-175-105-96 0-175 48-175 105 0 57 79 105 175 105 96 0 175-48 175-105z"/>
					</xsl:when>
					<xsl:when test ="($w='lg') and ($len='lg')">
						<draw:marker draw:name="msArrowOvalEnd9" draw:display-name="msArrowOvalEnd 9" svg:viewBox="0 0 350 350" svg:d="m350 0c0-96-79-175-175-175-96 0-175 79-175 175 0 96 79 175 175 175 96 0 175-79 175-175z"/>
					</xsl:when>
					<xsl:otherwise>
						<draw:marker draw:name="msArrowOvalEnd5" draw:display-name="msArrowOvalEnd 5" svg:viewBox="0 0 210 210" svg:d="m210 0c0-57-48-105-105-105-57 0-105 48-105 105 0 57 48 105 105 105 57 0 105-48 105-105z"/>
					</xsl:otherwise>
				</xsl:choose>
			</xsl:when>

			<!--Diamond-->
			<xsl:when test ="($type='diamond')">
				<xsl:choose>
					<xsl:when test ="($w='sm') and ($len='sm')">
						<draw:marker draw:name="msArrowDiamondEnd1" draw:display-name="msArrowDiamondEnd 1" svg:viewBox="0 0 140 140" svg:d="m70 0 70 70-70 70-70-70z"/>
					</xsl:when>
					<xsl:when test ="($w='sm') and ($len='med')">
						<draw:marker draw:name="msArrowDiamondEnd2" draw:display-name="msArrowDiamondEnd 2" svg:viewBox="0 0 140 210" svg:d="m70 0 70 105-70 105-70-105z"/>
					</xsl:when>
					<xsl:when test ="($w='sm') and ($len='lg')">
						<draw:marker draw:name="msArrowDiamondEnd3" draw:display-name="msArrowDiamondEnd 3" svg:viewBox="0 0 140 350" svg:d="m70 0 70 175-70 175-70-175z"/>
					</xsl:when>
					<xsl:when test ="($w='med') and ($len='sm')">
						<draw:marker draw:name="msArrowDiamondEnd4" draw:display-name="msArrowDiamondEnd 4" svg:viewBox="0 0 210 140" svg:d="m105 0 105 70-105 70-105-70z"/>
					</xsl:when>
					<xsl:when test ="($w='med') and ($len='lg')">
						<draw:marker draw:name="msArrowDiamondEnd6" draw:display-name="msArrowDiamondEnd 6" svg:viewBox="0 0 210 350" svg:d="m105 0 105 175-105 175-105-175z"/>
					</xsl:when>
					<xsl:when test ="($w='lg') and ($len='sm')">
						<draw:marker draw:name="msArrowDiamondEnd7" draw:display-name="msArrowDiamondEnd 7" svg:viewBox="0 0 350 140" svg:d="m175 0 175 70-175 70-175-70z"/>
					</xsl:when>
					<xsl:when test ="($w='lg') and ($len='med')">
						<draw:marker draw:name="msArrowDiamondEnd8" draw:display-name="msArrowDiamondEnd 8" svg:viewBox="0 0 350 210" svg:d="m175 0 175 105-175 105-175-105z"/>
					</xsl:when>
					<xsl:when test ="($w='lg') and ($len='lg')">
						<draw:marker draw:name="msArrowDiamondEnd9" draw:display-name="msArrowDiamondEnd 9" svg:viewBox="0 0 350 350" svg:d="m175 0 175 175-175 175-175-175z"/>
					</xsl:when>
					<xsl:otherwise>
						<draw:marker draw:name="msArrowDiamondEnd5" draw:display-name="msArrowDiamondEnd 5" svg:viewBox="0 0 210 210" svg:d="m105 0 105 105-105 105-105-105z"/>
					</xsl:otherwise>
				</xsl:choose>
			</xsl:when >

		</xsl:choose>
	</xsl:template>

	<xsl:template name ="getDashType">
		<xsl:param name ="val" />
		<xsl:param name ="cap" />
		<xsl:choose>
			<xsl:when test ="($val='sysDot')">
				<xsl:if test ="($cap='rnd')">
					<draw:stroke-dash draw:name="sysDotRound" draw:display-name="Dash 2" draw:style="round" draw:dots1="1" draw:dots1-length="0.07cm" draw:distance="0.07cm"/>
				</xsl:if>
				<xsl:if test ="($cap!='rnd') or not($cap)">
					<draw:stroke-dash draw:name="sysDot" draw:display-name="Dash 2" draw:style="rect" draw:dots1="1" draw:dots1-length="0.07cm" draw:distance="0.07cm"/>
				</xsl:if>
			</xsl:when>
			<xsl:when test ="($val='sysDash')">
				<xsl:if test ="($cap='rnd')">
					<draw:stroke-dash draw:name="sysDashRound" draw:display-name="Dash 2" draw:style="round" draw:dots1="1" draw:dots1-length="0.07cm" draw:distance="0.07cm"/>
				</xsl:if>
				<xsl:if test ="($cap!='rnd') or not($cap)">
					<draw:stroke-dash draw:name="sysDash" draw:display-name="Dash 2" draw:style="rect" draw:dots1="1" draw:dots1-length="0.07cm" draw:distance="0.07cm"/>
				</xsl:if>
			</xsl:when>
			<xsl:when test ="($val='dash')">
				<xsl:if test ="($cap='rnd')">
					<draw:stroke-dash draw:name="dashRound" draw:display-name="Dash 3" draw:style="round" draw:dots2="1" draw:dots2-length="0.282cm" draw:distance="0.211cm"/>
				</xsl:if>
				<xsl:if test ="($cap!='rnd') or not($cap)">
					<draw:stroke-dash draw:name="dash" draw:display-name="Dash 3" draw:style="rect" draw:dots2="1" draw:dots2-length="0.282cm" draw:distance="0.211cm"/>
				</xsl:if>
			</xsl:when>
			<xsl:when test ="($val='dashDot')">
				<xsl:if test ="($cap='rnd')">
					<draw:stroke-dash draw:name="dashDotRound" draw:display-name="Dash 4" draw:style="round" draw:dots1="1" draw:dots1-length="0.07cm" draw:dots2="1" draw:dots2-length="0.282cm" draw:distance="0.211cm"/>
				</xsl:if>
				<xsl:if test ="($cap!='rnd') or not($cap)">
					<draw:stroke-dash draw:name="dashDot" draw:display-name="Dash 4" draw:style="rect" draw:dots1="1" draw:dots1-length="0.07cm" draw:dots2="1" draw:dots2-length="0.282cm" draw:distance="0.211cm"/>
				</xsl:if>
			</xsl:when>
			<xsl:when test ="($val='lgDash')">
				<xsl:if test ="($cap='rnd')">
					<draw:stroke-dash draw:name="lgDashRound" draw:display-name="Dash 5" draw:style="round" draw:dots2="1" draw:dots2-length="0.564cm" draw:distance="0.211cm"/>
				</xsl:if>
				<xsl:if test ="($cap!='rnd') or not($cap)">
					<draw:stroke-dash draw:name="lgDash" draw:display-name="Dash 5" draw:style="rect" draw:dots2="1" draw:dots2-length="0.564cm" draw:distance="0.211cm"/>
				</xsl:if>
			</xsl:when>
			<xsl:when test ="($val='lgDashDot')">
				<xsl:if test ="($cap='rnd')">
					<draw:stroke-dash draw:name="lgDashDotRound" draw:display-name="Dash 7" draw:style="round" draw:dots1="1" draw:dots1-length="0.07cm" draw:dots2="1" draw:dots2-length="0.564cm" draw:distance="0.211cm"/>
				</xsl:if>
				<xsl:if test ="($cap!='rnd') or not($cap)">
					<draw:stroke-dash draw:name="lgDashDot" draw:display-name="Dash 7" draw:style="rect" draw:dots1="1" draw:dots1-length="0.07cm" draw:dots2="1" draw:dots2-length="0.564cm" draw:distance="0.211cm"/>
				</xsl:if>
			</xsl:when>
			<xsl:when test ="($val='lgDashDotDot')">
				<xsl:if test ="($cap='rnd')">
					<draw:stroke-dash draw:name="lgDashDotDotRound" draw:display-name="Dash 8" draw:style="round" draw:dots1="2" draw:dots1-length="0.07cm" draw:dots2="1" draw:dots2-length="0.564cm" draw:distance="0.211cm"/>
				</xsl:if>
				<xsl:if test ="($cap!='rnd') or not($cap)">
					<draw:stroke-dash draw:name="lgDashDotDot" draw:display-name="Dash 8" draw:style="rect" draw:dots1="2" draw:dots1-length="0.07cm" draw:dots2="1" draw:dots2-length="0.564cm" draw:distance="0.211cm"/>
				</xsl:if>
			</xsl:when>
		</xsl:choose>
	</xsl:template>
	<!-- Changes made by Vijayeta-->
	<!-- Slide Size-->
	<xsl:template name="InsertSlideSize">
		<xsl:for-each select ="document('ppt/presentation.xml')//p:presentation/p:sldSz">
			<style:page-layout style:name="PM1">
				<style:page-layout-properties 
				 fo:margin-top="0cm" 
				 fo:margin-bottom="0cm" 
				 fo:margin-left="0cm" 
				 fo:margin-right="0cm">
					<xsl:attribute name ="fo:page-width">
						<xsl:call-template name="ConvertEmu">
							<xsl:with-param name="length" select="@cx"/>
							<xsl:with-param name="unit">cm</xsl:with-param>
						</xsl:call-template>
					</xsl:attribute>
					<xsl:attribute name ="fo:page-height">
						<xsl:call-template name="ConvertEmu">
							<xsl:with-param name="length" select="@cy"/>
							<xsl:with-param name="unit">cm</xsl:with-param>
						</xsl:call-template>
					</xsl:attribute>
					<xsl:attribute name ="style:print-orientation">
						<xsl:call-template name="CheckOrientation">
							<xsl:with-param name="cx" select="@cx"/>
							<xsl:with-param name="cy" select="@cy"/>
						</xsl:call-template>
						<!--<xsl:value-of select ="'portrait'"/>-->
					</xsl:attribute>
				</style:page-layout-properties>
			</style:page-layout>
			<!--<style:page-layout style:name="PM1">
        <style:page-layout-properties fo:margin-top="0cm" fo:margin-bottom="0cm" fo:margin-left="0cm" fo:margin-right="0cm" fo:page-width="28cm" fo:page-height="21cm" style:print-orientation="landscape"/>
      </style:page-layout>-->
		</xsl:for-each>
	</xsl:template>
	<!-- Notes Size-->
	<xsl:template name="InsertNotesSize">
		<!-- Check if notesSlide is present in the package-->
		<xsl:if test ="document('ppt/notesSlides/notesSlide1.xml')">
			<xsl:variable name ="Flag">
				<!--Check if size defined in notesSlide -->
				<xsl:for-each select ="document('ppt/notesSlides/notesSlide1.xml')//p:notes/p:cSld/p:spTree/p:sp">
					<xsl:if test ="p:nvSpPr/p:cNvPr/@name[contains(.,'Notes Placeholder 2')]">
						<xsl:if test ="p:spPr/a:xfrm/a:ext[@cx]">
							<xsl:value-of select ="'true'"/>
						</xsl:if>
						<xsl:if test ="not(p:spPr/a:xfrm/a:ext[@cx])">
							<xsl:value-of select ="'false'"/>
						</xsl:if>
					</xsl:if>
				</xsl:for-each>
			</xsl:variable>
			<xsl:choose >
				<xsl:when test ="$Flag='true'">
					<!-- notesSlide has Size Definition(user defined)-->
					<xsl:for-each select ="document('ppt/notesSlides/notesSlide1.xml')//p:notes/p:cSld/p:spTree/p:sp">
						<xsl:if test ="p:nvSpPr/p:cNvPr/@name[contains(.,'Notes Placeholder 2')]">
							<style:page-layout style:name="PM2">
								<style:page-layout-properties>
									<xsl:for-each select ="p:spPr/a:xfrm/a:ext[@cx]">
										<xsl:attribute name ="fo:page-width">
											<xsl:call-template name="ConvertEmu">
												<xsl:with-param name="length" select="@cx"/>
												<xsl:with-param name="unit">cm</xsl:with-param>
											</xsl:call-template>
										</xsl:attribute>
									</xsl:for-each>
									<xsl:for-each select ="p:spPr/a:xfrm/a:ext[@cy]">
										<xsl:attribute name ="fo:page-height">
											<xsl:call-template name="ConvertEmu">
												<xsl:with-param name="length" select="@cy"/>
												<xsl:with-param name="unit">cm</xsl:with-param>
											</xsl:call-template>
										</xsl:attribute>
									</xsl:for-each>
									<xsl:attribute name ="style:print-orientation">
										<xsl:call-template name="CheckOrientation">
											<xsl:with-param name="cx" select="@cx"/>
											<xsl:with-param name="cy" select="@cy"/>
										</xsl:call-template>
									</xsl:attribute>
								</style:page-layout-properties>
							</style:page-layout>
						</xsl:if>
					</xsl:for-each>
				</xsl:when>
				<xsl:otherwise>
					<!-- pre-defined size definition in notesMaster-->
					<xsl:for-each select ="document('ppt/notesMasters/notesMaster1.xml')//p:notesMaster/p:cSld/p:spTree/p:sp">
						<xsl:if test ="p:nvSpPr/p:cNvPr/@name[contains(.,'Notes Placeholder 4')]">
							<style:page-layout style:name="PM2">
								<style:page-layout-properties>
									<xsl:for-each select ="p:spPr/a:xfrm/a:ext[@cx]">
										<xsl:attribute name ="fo:page-width">
											<xsl:call-template name="ConvertEmu">
												<xsl:with-param name="length" select="@cx"/>
												<xsl:with-param name="unit">cm</xsl:with-param>
											</xsl:call-template>
										</xsl:attribute>
									</xsl:for-each>
									<xsl:for-each select ="p:spPr/a:xfrm/a:ext[@cy]">
										<xsl:attribute name ="fo:page-height">
											<xsl:call-template name="ConvertEmu">
												<xsl:with-param name="length" select="@cy"/>
												<xsl:with-param name="unit">cm</xsl:with-param>
											</xsl:call-template>
										</xsl:attribute>
									</xsl:for-each>
									<xsl:attribute name ="style:print-orientation">
										<xsl:call-template name="CheckOrientation">
											<xsl:with-param name="cx" select="@cx"/>
											<xsl:with-param name="cy" select="@cy"/>
										</xsl:call-template>
									</xsl:attribute>
								</style:page-layout-properties>
							</style:page-layout>
						</xsl:if>
					</xsl:for-each>
				</xsl:otherwise>
			</xsl:choose>
		</xsl:if>
		<!--Default Size in presentation.xml-->
		<xsl:if test ="not(document('ppt/notesSlides/notesSlide1.xml'))">
			<xsl:for-each select ="document('ppt/presentation.xml')//p:presentation/p:notesSz">
				<style:page-layout style:name="PM2">
					<style:page-layout-properties>
						<xsl:attribute name ="fo:page-width">
							<xsl:call-template name="ConvertEmu">
								<xsl:with-param name="length" select="@cx"/>
								<xsl:with-param name="unit">cm</xsl:with-param>
							</xsl:call-template>
						</xsl:attribute>
						<xsl:attribute name ="fo:page-height">
							<xsl:call-template name="ConvertEmu">
								<xsl:with-param name="length" select="@cy"/>
								<xsl:with-param name="unit">cm</xsl:with-param>
							</xsl:call-template>
						</xsl:attribute>
						<xsl:attribute name ="style:print-orientation">
							<xsl:call-template name="CheckOrientation">
								<xsl:with-param name="cx" select="@cx"/>
								<xsl:with-param name="cy" select="@cy"/>
							</xsl:call-template>
						</xsl:attribute>
					</style:page-layout-properties>
				</style:page-layout>
			</xsl:for-each>
		</xsl:if>
	</xsl:template>
	<!-- Slide Number,The equvivalent Not present in ODP-->
	<xsl:template name="InsertSlideNumber">
		<xsl:for-each select ="document('ppt/presentation.xml')//p:presentation">
			<draw:frame presentation:class="page-number">
				<draw:text-box>
					<text:p>
						<xsl:if test ="@firstSlideNum">
							<text:page-number>
								<xsl:value-of select ="@firstSlideNum"/>
							</text:page-number>
						</xsl:if>
					</text:p>
				</draw:text-box>
			</draw:frame>
		</xsl:for-each>
	</xsl:template>
	<!-- Add MasterStyles Definition-->

	<xsl:template name="InsertMasterStylesDefinition">

		<style:handout-master style:page-layout-name="PM0"/>
		<style:master-page style:name="Default" style:page-layout-name="PM1">
			<xsl:call-template name ="SetDateFooterPageNumberPosition">
				<xsl:with-param name ="PlaceHolder">
					<xsl:value-of select ="'dt'"/>
				</xsl:with-param>
				<xsl:with-param name ="PresentationClass">
					<xsl:value-of select ="'date-time'"/>
				</xsl:with-param>
			</xsl:call-template>
			<xsl:call-template name ="SetDateFooterPageNumberPosition">
				<xsl:with-param name ="PlaceHolder">
					<xsl:value-of select ="'ftr'"/>
				</xsl:with-param>
				<xsl:with-param name ="PresentationClass">
					<xsl:value-of select ="'footer'"/>
				</xsl:with-param>
			</xsl:call-template>
			<xsl:call-template name ="SetDateFooterPageNumberPosition">
				<xsl:with-param name ="PlaceHolder">
					<xsl:value-of select ="'sldNum'"/>
				</xsl:with-param>
				<xsl:with-param name ="PresentationClass">
					<xsl:value-of select ="'page-number'"/>
				</xsl:with-param>
			</xsl:call-template>
			<presentation:notes style:page-layout-name="PM2">
				<xsl:call-template name="InsertSlideNumber" />
			</presentation:notes>
		</style:master-page>
	</xsl:template>
	<!--checks cx/cy ratio,for orientation-->
	<xsl:template name ="CheckOrientation">
		<xsl:param name="cx"/>
		<xsl:param name="cy"/>
		<xsl:variable name="orientation"/>
		<xsl:choose>
			<xsl:when test ="$cx > $cy">
				<xsl:value-of select="'landscape'" />
			</xsl:when>
			<xsl:otherwise >
				<xsl:value-of select ="'portrait'"/>
			</xsl:otherwise>
		</xsl:choose>
		<xsl:value-of select="$orientation"/>
	</xsl:template>
	<!--  converts emu to given unit-->
	<xsl:template name="ConvertEmu">
		<xsl:param name="length"/>
		<xsl:param name="unit"/>
		<xsl:choose>
			<xsl:when
					test="$length = '' or not($length) or $length = 0 or format-number($length div 360000, '#.##') = ''">
				<xsl:value-of select="concat(0,'cm')"/>
			</xsl:when>
			<xsl:when test="$unit = 'cm'">
				<xsl:value-of select="concat(format-number($length div 360000, '#.##'), 'cm')"/>
			</xsl:when>
		</xsl:choose>
	</xsl:template>
	<!-- Changes made by Vijayeta-->
	<!--SetDateFooterPageNumberValues -->
	<xsl:template name ="SetDateFooterPageNumberPosition">
		<xsl:param name ="PlaceHolder"/>
		<xsl:param name ="PresentationClass"/>
		<xsl:for-each select ="document('ppt/presentation.xml')/p:presentation/p:sldIdLst/p:sldId">
			<!-- for each slide-->
			<xsl:variable name ="currentpos">
				<xsl:value-of select="position()"/>
			</xsl:variable>
			<xsl:variable name ="footerSlide">
				<xsl:value-of select ="concat(concat('ppt/slides/slide',position()),'.xml')"/>
			</xsl:variable>
			<xsl:variable name ="Flag">
				<xsl:for-each select ="document($footerSlide)/p:sld/p:cSld/p:spTree/p:sp">
					<!-- for each shape in current slide-->
					<xsl:if test ="p:nvSpPr/p:nvPr/p:ph/@type[contains(.,$PlaceHolder)]">
						<xsl:if test ="p:spPr/a:xfrm">
							<xsl:value-of select ="'slide'"/>
						</xsl:if>
					</xsl:if>
				</xsl:for-each>
				<!--END, for each shape in current slide-->
				<xsl:variable name ="SlideLayout">
					<xsl:variable name ="bool">
						<xsl:for-each select ="document($footerSlide)/p:sld/p:cSld/p:spTree/p:sp">
							<xsl:if test ="p:nvSpPr/p:nvPr/p:ph/@type[contains(.,$PlaceHolder)]">
								<xsl:if  test ="not(p:sp/p:spPr/a:xfrm)">
									<xsl:value-of select ="'true'"/>
								</xsl:if>
							</xsl:if>
						</xsl:for-each>
					</xsl:variable>
					<xsl:if test = "$bool = 'true'">
						<xsl:for-each select ="document(concat(concat(('ppt/slides/_rels/slide'),$currentpos),'.xml.rels'))//rels:Relationships/rels:Relationship[@Target]">
							<!-- for each find 'slideLayoutX.xml'-->
							<xsl:value-of select ="substring(@Target,17)"/>
						</xsl:for-each>
					</xsl:if>
				</xsl:variable>
				<xsl:for-each select ="document(concat(('ppt/slideLayouts/'),($SlideLayout)))//p:sldLayout/p:cSld/p:spTree/p:sp">
					<!-- for each Shape in layout-->
					<xsl:if test ="p:nvSpPr/p:nvPr/p:ph/@type[contains(.,$PlaceHolder)]">
						<xsl:if test ="p:spPr/a:xfrm">
							<xsl:value-of select ="'layout'"/>
						</xsl:if>
					</xsl:if>
				</xsl:for-each>
				<!--END, for each Shape in layout-->
			</xsl:variable>
			<xsl:choose >
				<xsl:when test ="$Flag = 'slide'">
					<xsl:for-each select ="document($footerSlide)/p:sld/p:cSld/p:spTree/p:sp">
						<xsl:if test ="p:nvSpPr/p:nvPr/p:ph/@type[contains(.,$PlaceHolder)]">
							<draw:frame draw:layer="backgroundobjects" presentation:style-name="pr1">
								<xsl:for-each select ="p:spPr/a:xfrm/a:off[@x]">
									<xsl:attribute name ="presentation:class">
										<xsl:value-of select ="$PresentationClass"/>
									</xsl:attribute>
									<xsl:attribute name ="svg:x">
										<xsl:call-template name="ConvertEmu">
											<xsl:with-param name="length" select="@x"/>
											<xsl:with-param name="unit">cm</xsl:with-param>
										</xsl:call-template>
									</xsl:attribute>
								</xsl:for-each>
								<xsl:for-each select ="p:spPr/a:xfrm/a:off[@y]">
									<xsl:attribute name ="svg:y">
										<xsl:call-template name="ConvertEmu">
											<xsl:with-param name="length" select="@y"/>
											<xsl:with-param name="unit">cm</xsl:with-param>
										</xsl:call-template>
									</xsl:attribute>
								</xsl:for-each>
								<xsl:for-each select ="p:spPr/a:xfrm/a:ext[@cx]">
									<xsl:attribute name ="svg:width">
										<xsl:call-template name="ConvertEmu">
											<xsl:with-param name="length" select="@cx"/>
											<xsl:with-param name="unit">cm</xsl:with-param>
										</xsl:call-template>
									</xsl:attribute>
								</xsl:for-each>
								<xsl:for-each select ="p:spPr/a:xfrm/a:ext[@cy]">
									<xsl:attribute name ="svg:height">
										<xsl:call-template name="ConvertEmu">
											<xsl:with-param name="length" select="@cy"/>
											<xsl:with-param name="unit">cm</xsl:with-param>
										</xsl:call-template>
										<xsl:value-of select ="SlideHeight"/>
									</xsl:attribute>
								</xsl:for-each>
								<xsl:if test ="$PresentationClass = 'date-time'">
									<xsl:call-template name ="date-time"/>
								</xsl:if>
								<xsl:if test ="$PresentationClass = 'footer'">
									<xsl:call-template name ="footer"/>
								</xsl:if>
								<xsl:if test ="$PresentationClass = 'page-number'">
									<xsl:call-template name ="page-number"/>
								</xsl:if>
							</draw:frame>
						</xsl:if>
					</xsl:for-each>
				</xsl:when>
				<xsl:when test ="$Flag = 'layout'">
					<xsl:variable name ="SlideLayout">
						<xsl:for-each select ="document(concat(concat(('ppt/slides/_rels/slide'),$currentpos),'.xml.rels'))//rels:Relationships/rels:Relationship[@Target]">
							<xsl:value-of select ="substring(@Target,17)"/>
						</xsl:for-each>
					</xsl:variable>
					<!--<xsl:for-each select ="document(concat(('ppt/slideLayouts/'),($SlideLayout)))//p:sldLayout/p:cSld/p:spTree/p:sp">-->
					<xsl:for-each select ="document(concat(('ppt/slideLayouts/'),($SlideLayout)))//p:sldLayout/p:cSld/p:spTree/p:sp">
						<xsl:if test ="p:nvSpPr/p:nvPr/p:ph/@type[contains(.,$PlaceHolder)]">
							<draw:frame draw:layer="backgroundobjects" presentation:style-name="pr1">
								<xsl:for-each select ="p:spPr/a:xfrm/a:off[@x]">
									<xsl:attribute name ="presentation:class">
										<xsl:value-of select ="$PresentationClass"/>
									</xsl:attribute>
									<xsl:attribute name ="svg:x">
										<xsl:call-template name="ConvertEmu">
											<xsl:with-param name="length" select="@x"/>
											<xsl:with-param name="unit">cm</xsl:with-param>
										</xsl:call-template>
									</xsl:attribute>
								</xsl:for-each>
								<xsl:for-each select ="p:spPr/a:xfrm/a:off[@y]">
									<xsl:attribute name ="svg:y">
										<xsl:call-template name="ConvertEmu">
											<xsl:with-param name="length" select="@y"/>
											<xsl:with-param name="unit">cm</xsl:with-param>
										</xsl:call-template>
									</xsl:attribute>
								</xsl:for-each>
								<xsl:for-each select ="p:spPr/a:xfrm/a:ext[@cx]">
									<xsl:attribute name ="svg:width">
										<xsl:call-template name="ConvertEmu">
											<xsl:with-param name="length" select="@cx"/>
											<xsl:with-param name="unit">cm</xsl:with-param>
										</xsl:call-template>
									</xsl:attribute>
								</xsl:for-each>
								<xsl:for-each select ="p:spPr/a:xfrm/a:ext[@cy]">
									<xsl:attribute name ="svg:height">
										<xsl:call-template name="ConvertEmu">
											<xsl:with-param name="length" select="@cy"/>
											<xsl:with-param name="unit">cm</xsl:with-param>
										</xsl:call-template>
										<xsl:value-of select ="SlideHeight"/>
									</xsl:attribute>
								</xsl:for-each>
								<xsl:if test ="$PresentationClass = 'date-time'">
									<xsl:call-template name ="date-time"/>
								</xsl:if>
								<xsl:if test ="$PresentationClass = 'footer'">
									<xsl:call-template name ="footer"/>
								</xsl:if>
								<xsl:if test ="$PresentationClass = 'page-number'">
									<xsl:call-template name ="page-number"/>
								</xsl:if>
							</draw:frame>
						</xsl:if>
					</xsl:for-each>
				</xsl:when>
				<xsl:otherwise>
					<!-- SlideMaster-->
					<xsl:for-each select ="document('ppt/slideMasters/slideMaster1.xml')//p:sldMaster/p:cSld/p:spTree/p:sp">
						<xsl:if test ="p:nvSpPr/p:nvPr/p:ph/@type[contains(.,$PlaceHolder)]">
							<draw:frame draw:layer="backgroundobjects" presentation:style-name="pr1">
								<xsl:attribute name ="presentation:class">
									<xsl:value-of select ="$PresentationClass"/>
								</xsl:attribute>
								<xsl:for-each select ="p:spPr/a:xfrm/a:off[@x]">
									<xsl:attribute name ="svg:x">
										<xsl:call-template name="ConvertEmu">
											<xsl:with-param name="length" select="@x"/>
											<xsl:with-param name="unit">cm</xsl:with-param>
										</xsl:call-template>
									</xsl:attribute>
								</xsl:for-each>
								<xsl:for-each select ="p:spPr/a:xfrm/a:off[@y]">
									<xsl:attribute name ="svg:y">
										<xsl:call-template name="ConvertEmu">
											<xsl:with-param name="length" select="@y"/>
											<xsl:with-param name="unit">cm</xsl:with-param>
										</xsl:call-template>
									</xsl:attribute>
								</xsl:for-each>
								<xsl:for-each select ="p:spPr/a:xfrm/a:ext[@cx]">
									<xsl:attribute name ="svg:width">
										<xsl:call-template name="ConvertEmu">
											<xsl:with-param name="length" select="@cx"/>
											<xsl:with-param name="unit">cm</xsl:with-param>
										</xsl:call-template>
									</xsl:attribute>
								</xsl:for-each>
								<xsl:for-each select ="p:spPr/a:xfrm/a:ext[@cy]">
									<xsl:attribute name ="svg:height">
										<xsl:call-template name="ConvertEmu">
											<xsl:with-param name="length" select="@cy"/>
											<xsl:with-param name="unit">cm</xsl:with-param>
										</xsl:call-template>
										<xsl:value-of select ="SlideHeight"/>
									</xsl:attribute>
								</xsl:for-each>
								<xsl:if test ="$PresentationClass = 'date-time'">
									<xsl:call-template name ="date-time"/>
								</xsl:if>
								<xsl:if test ="$PresentationClass = 'footer'">
									<xsl:call-template name ="footer"/>
								</xsl:if>
								<xsl:if test ="$PresentationClass = 'page-number'">
									<xsl:call-template name ="page-number"/>
								</xsl:if>
							</draw:frame>
						</xsl:if>
					</xsl:for-each>
				</xsl:otherwise>
				<!-- END,SlideMaster-->
			</xsl:choose>
		</xsl:for-each>
		<!-- END,for each slide-->

	</xsl:template>
	<xsl:template name ="date-time">
		<draw:text-box>
			<text:p>
				<presentation:date-time/>
			</text:p>
		</draw:text-box>
	</xsl:template>
	<xsl:template name ="footer">
		<draw:text-box>
			<text:p>
				<presentation:footer/>
			</text:p>
		</draw:text-box>
	</xsl:template>
	<xsl:template name ="page-number">
		<draw:text-box>
			<text:p>
				<text:page-number> &lt;number&gt;</text:page-number>
			</text:p>
		</draw:text-box>
	</xsl:template>
</xsl:stylesheet>

