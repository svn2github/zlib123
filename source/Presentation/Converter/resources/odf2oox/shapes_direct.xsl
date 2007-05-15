<?xml version="1.0" encoding="utf-8" ?>
<xsl:stylesheet version="1.0" 
  xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
  xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"   
  xmlns:odf="urn:odf"
  xmlns:pzip="urn:cleverage:xmlns:post-processings:zip"  
  xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" 
  xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"
  xmlns:number="urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0"
  xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0"
  xmlns:page="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0"
  xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0" 
  xmlns:presentation="urn:oasis:names:tc:opendocument:xmlns:presentation:1.0"
  xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0"
  xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" 
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" 
  xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
  exclude-result-prefixes="odf style text number draw page">
	<xsl:import href="common.xsl"/>

	<!-- Shapes in presentation-->
	<xsl:template name ="shapes">
		<!-- Rectangle -->
		<xsl:for-each select ="draw:rect">
			<xsl:call-template name ="CreateShape">
				<xsl:with-param name ="shapeName" select="'Rectangle '" />
			</xsl:call-template>
		</xsl:for-each>
		<!-- Rectangle - Custom shape -->
		<xsl:for-each select ="draw:custom-shape">
			<xsl:if test ="(draw:enhanced-geometry/@draw:type='rectangle') and (draw:enhanced-geometry/@draw:enhanced-path='M 0 0 L 21600 0 21600 21600 0 21600 0 0 Z N')">
				<xsl:call-template name ="CreateShape">
					<xsl:with-param name ="shapeName" select="'Rectangle Custom '" />
				</xsl:call-template>
			</xsl:if>
		</xsl:for-each>

		<!-- Text box-->
		<xsl:for-each select ="draw:frame">
			<xsl:if test ="(draw:text-box) and not(@presentation:style-name) and not(@presentation:class)">
				<xsl:call-template name ="CreateShape">
					<xsl:with-param name ="shapeName" select="'TextBox '" />
				</xsl:call-template>
			</xsl:if>
		</xsl:for-each>
		<!-- Text box - Custom shape-->
		<xsl:for-each select ="draw:custom-shape">
			<xsl:if test ="(draw:enhanced-geometry/@draw:type='mso-spt202') and (draw:enhanced-geometry/@draw:enhanced-path='M 0 0 L 21600 0 21600 21600 0 21600 0 0 Z N')">
				<xsl:call-template name ="CreateShape">
					<xsl:with-param name ="shapeName" select="'TextBox Custom '" />
				</xsl:call-template>
			</xsl:if>
		</xsl:for-each>

		<!-- Line -->
		<!--<xsl:for-each select ="draw:line">
			<xsl:call-template name ="drawLine" />
		</xsl:for-each>-->
	</xsl:template>
	<!-- Create p:sp node for shape -->
	<xsl:template name ="CreateShape">
		<xsl:param name ="shapeName" />
		<p:sp>
			<p:nvSpPr>
				<p:cNvPr>
					<xsl:attribute name ="id">
						<xsl:value-of select ="position()"/>
					</xsl:attribute>
					<xsl:attribute name ="name">
						<xsl:value-of select ="concat($shapeName, position())"/>
					</xsl:attribute>
				</p:cNvPr >
				<p:cNvSpPr />
				<p:nvPr />
			</p:nvSpPr>
			<p:spPr>

				<xsl:call-template name ="drawShape" />

				<xsl:call-template name ="getPresetGeom">
					<xsl:with-param name ="prstGeom" select ="$shapeName" />
				</xsl:call-template>

				<xsl:call-template name ="getGraphicProperties">
					<xsl:with-param name ="gr" select="@draw:style-name" />
				</xsl:call-template>
			</p:spPr>
			<p:style>
				<a:lnRef idx="2">
					<a:schemeClr val="dk1" />
				</a:lnRef>
				<a:fillRef idx="1">
					<a:srgbClr val="99ccff">
						<xsl:call-template name ="getShade">
							<xsl:with-param name ="gr" select ="@draw:style-name" />
						</xsl:call-template>
					</a:srgbClr>
				</a:fillRef>
				<a:effectRef idx="0">
					<a:schemeClr val="accent1"/>
				</a:effectRef>
				<a:fontRef idx="minor">
					<a:schemeClr val="dk1"/>
				</a:fontRef>
			</p:style>
			<p:txBody>
				<xsl:call-template name ="getParagraphProperties">
					<xsl:with-param name ="gr" select="@draw:style-name" />
				</xsl:call-template>

				<xsl:choose>
					<xsl:when test ="draw:text-box">
						<xsl:for-each select ="draw:text-box">
							<xsl:call-template name ="processShapeText">
								<xsl:with-param name="gr" select ="parent::node()/@draw:style-name" />
							</xsl:call-template>
						</xsl:for-each>
					</xsl:when>
					<xsl:otherwise>
						<xsl:call-template name ="processShapeText">
							<xsl:with-param name ="gr" select="@draw:style-name" />
						</xsl:call-template>
					</xsl:otherwise>
				</xsl:choose>
			</p:txBody>

		</p:sp>
	</xsl:template>
	<!-- Draw shape -->
	<xsl:template name ="drawShape">
		<a:xfrm>
			<a:off >
				<xsl:attribute name ="x">
					<xsl:call-template name ="convertToPoints">
						<xsl:with-param name ="unit" select ="'cm'"/>
						<xsl:with-param name ="length" select ="@svg:x"/>
					</xsl:call-template>
				</xsl:attribute>
				<xsl:attribute name ="y">
					<xsl:call-template name ="convertToPoints">
						<xsl:with-param name ="unit" select ="'cm'"/>
						<xsl:with-param name ="length" select ="@svg:y"/>
					</xsl:call-template>
				</xsl:attribute>
			</a:off>
			<a:ext>
				<xsl:attribute name ="cx">
					<xsl:call-template name ="convertToPoints">
						<xsl:with-param name ="unit" select ="'cm'"/>
						<xsl:with-param name ="length" select ="@svg:width"/>
					</xsl:call-template>
				</xsl:attribute>
				<xsl:attribute name ="cy">
					<xsl:call-template name ="convertToPoints">
						<xsl:with-param name ="unit" select ="'cm'"/>
						<xsl:with-param name ="length" select ="@svg:height"/>
					</xsl:call-template>
				</xsl:attribute>
			</a:ext>
		</a:xfrm>
	</xsl:template>
	<!-- Draw line-->
	<xsl:template name ="drawLine">
		<p:cxnSp>
			<p:nvCxnSpPr>
				<p:cNvPr>
					<xsl:attribute name ="id">
						<xsl:value-of select ="position()"/>
					</xsl:attribute>
					<xsl:attribute name ="name">
						<xsl:value-of select ="concat('Straight Connector ', position())" />
					</xsl:attribute>
				</p:cNvPr>
				<p:cNvCxnSpPr/>
				<p:nvPr/>
			</p:nvCxnSpPr>
			<p:spPr>

				<xsl:variable name ="x1">
					<xsl:call-template name ="convertToPoints">
						<xsl:with-param name ="unit" select ="'cm'"/>
						<xsl:with-param name ="length" select ="@svg:x1"/>
					</xsl:call-template>
				</xsl:variable>
				<xsl:variable name ="x2">
					<xsl:call-template name ="convertToPoints">
						<xsl:with-param name ="unit" select ="'cm'"/>
						<xsl:with-param name ="length" select ="@svg:x2"/>
					</xsl:call-template>
				</xsl:variable>
				<xsl:variable name ="y1">
					<xsl:call-template name ="convertToPoints">
						<xsl:with-param name ="unit" select ="'cm'"/>
						<xsl:with-param name ="length" select ="@svg:y1"/>
					</xsl:call-template>
				</xsl:variable>
				<xsl:variable name ="y2">
					<xsl:call-template name ="convertToPoints">
						<xsl:with-param name ="unit" select ="'cm'"/>
						<xsl:with-param name ="length" select ="@svg:y2"/>
					</xsl:call-template>
				</xsl:variable>

				<a:xfrm>
					<a:off >
						<xsl:attribute name ="x">
							<xsl:value-of select ="$x1"/>
						</xsl:attribute>
						<xsl:attribute name ="y">
							<xsl:value-of select ="$y1"/>
						</xsl:attribute>
					</a:off>
					<a:ext>
						<xsl:attribute name ="cx">
							<xsl:value-of select ="$x2 - $x1"/>
						</xsl:attribute>
						<xsl:attribute name ="cy">
							<xsl:value-of select ="$y2 - $y1"/>
						</xsl:attribute>
					</a:ext>
				</a:xfrm>

				<xsl:call-template name ="getPresetGeom">
					<xsl:with-param name ="prstGeom" select ="'Straight Connector'" />
				</xsl:call-template>

				<xsl:call-template name ="getGraphicProperties">
					<xsl:with-param name ="gr" select="@draw:style-name" />
				</xsl:call-template>
			</p:spPr>
			<p:style>
				<a:lnRef idx="2">
					<a:schemeClr val="dk1" />
				</a:lnRef>
				<a:fillRef idx="1">
					<a:srgbClr val="99ccff" />
				</a:fillRef>
				<a:effectRef idx="0">
					<a:schemeClr val="accent1"/>
				</a:effectRef>
				<a:fontRef idx="minor">
					<a:schemeClr val="dk1"/>
				</a:fontRef>
			</p:style>

		</p:cxnSp>
	</xsl:template>
	<!-- Get graphic properties for shape in context.xml-->
	<xsl:template name ="getGraphicProperties">
		<xsl:param name ="gr" />
		
		<xsl:for-each select ="document('content.xml')/office:document-content/office:automatic-styles/style:style[@style:name=$gr]/style:graphic-properties">
			<!-- Parent style name-->
			<xsl:variable name ="parentStyle">
				<xsl:value-of select ="parent::node()/@style:parent-style-name"/>
			</xsl:variable>
			
			<!--FILL-->
			<xsl:call-template name ="getFillDetails">
				<xsl:with-param name ="parentStyle" select="$parentStyle" />
			</xsl:call-template >
			
			<!--LINE COLOR AND STYLE-->
			<xsl:call-template name ="getLineStyle">
				<xsl:with-param name ="parentStyle" select="$parentStyle" />
			</xsl:call-template >
			
		</xsl:for-each>

		
	</xsl:template>
	<!-- Get preset geometry-->
	<xsl:template name ="getPresetGeom">
		<xsl:param name ="prstGeom" />
		<xsl:choose>
			<xsl:when test ="(contains($prstGeom, 'Rectangle')) or (contains($prstGeom, 'TextBox'))">
				<a:prstGeom prst="rect">
					<a:avLst/>
				</a:prstGeom>
			</xsl:when>
			<xsl:when test ="contains($prstGeom, 'Straight Connector')">
				<a:prstGeom prst="line">
					<a:avLst/>
				</a:prstGeom>
			</xsl:when>
		</xsl:choose>
	</xsl:template>
	<!-- Get fill details-->
	<xsl:template name ="getFillDetails">
		<xsl:param name ="parentStyle" />
		<xsl:choose>
			<!-- No fill-->
			<xsl:when test ="@draw:fill='none'">
				<a:noFill/>
			</xsl:when>
			<!-- Solid fill-->
			<xsl:when test="@draw:fill-color">
				<xsl:call-template name ="getFillColor">
					<xsl:with-param name ="fill-color" select ="@draw:fill-color" />
					<xsl:with-param name ="opacity" select ="@draw:opacity" />
				</xsl:call-template>
			</xsl:when>
			<!-- Default fill color from styles.xml-->
			<xsl:when test ="($parentStyle != '')">
				 <xsl:for-each select ="document('styles.xml')/office:document-styles/office:styles/style:style[@style:name = $parentStyle]/style:graphic-properties">
					<xsl:choose>
						<xsl:when test ="@draw:fill='none'">
							<a:noFill/>
						</xsl:when>
						<xsl:when test="@draw:fill-color">
							<xsl:call-template name ="getFillColor">
								<xsl:with-param name ="fill-color" select ="@draw:fill-color" />
								<xsl:with-param name ="opacity" select ="@draw:opacity" />
							</xsl:call-template>
						</xsl:when>
					</xsl:choose>
				</xsl:for-each>
			</xsl:when>
		</xsl:choose>
	</xsl:template>

	<!-- Get line color and style-->
	<xsl:template name ="getLineStyle">
		<xsl:param name ="parentStyle" />
		<a:ln>
			<!-- Border width -->
			<xsl:choose>
				<xsl:when test ="@svg:stroke-width">
					<xsl:attribute name ="w">
						<xsl:call-template name ="convertToPoints">
							<xsl:with-param name="length"  select ="@svg:stroke-width"/>
							<xsl:with-param name ="unit" select ="'cm'"/>
						</xsl:call-template>
					</xsl:attribute>
				</xsl:when>
				<!-- Default border width from styles.xml-->
				<xsl:when test ="($parentStyle != '')">
					<xsl:for-each select ="document('styles.xml')/office:document-styles/office:styles/style:style[@style:name = $parentStyle]/style:graphic-properties">
						<xsl:if test ="@svg:stroke-width">
							<xsl:attribute name ="w">
								<xsl:call-template name ="convertToPoints">
									<xsl:with-param name="length"  select ="@svg:stroke-width"/>
									<xsl:with-param name ="unit" select ="'cm'"/>
								</xsl:call-template>
							</xsl:attribute>
						</xsl:if >
					</xsl:for-each>
				</xsl:when>
			</xsl:choose>
			<!-- Cap type-->
			<xsl:if test ="@draw:stroke-dash">
				<xsl:variable name ="dash" select ="@draw:stroke-dash" />
				<xsl:if test ="document('styles.xml')/office:document-styles/office:styles/draw:stroke-dash[@draw:name=$dash]">
					<xsl:if test ="document('styles.xml')/office:document-styles/office:styles/draw:stroke-dash[@draw:name=$dash]/@draw:style='round'">
						<xsl:attribute name ="cap">
							<xsl:value-of select ="'rnd'"/>
						</xsl:attribute>
					</xsl:if>
				</xsl:if>
			</xsl:if>

			<!-- Line color -->
			<xsl:choose>
				<!-- Invisible line-->
				<xsl:when test ="@draw:stroke='none'">
					<a:noFill />
				</xsl:when>
				<!-- Solid color -->
				<xsl:when test ="@svg:stroke-color">
					<xsl:call-template name ="getFillColor">
						<xsl:with-param name ="fill-color" select ="@svg:stroke-color" />
						<xsl:with-param name ="opacity" select ="@svg:stroke-opacity" />
					</xsl:call-template>
				</xsl:when>
				<xsl:when test ="($parentStyle != '')">
					<xsl:for-each select ="document('styles.xml')/office:document-styles/office:styles/style:style[@style:name = $parentStyle]/style:graphic-properties">
						<xsl:choose>
							<!-- Invisible line-->
							<xsl:when test ="@draw:stroke='none'">
								<a:noFill />
							</xsl:when>
							<!-- Solid color -->
							<xsl:when test ="@svg:stroke-color">
								<xsl:call-template name ="getFillColor">
									<xsl:with-param name ="fill-color" select ="@svg:stroke-color" />
									<xsl:with-param name ="opacity" select ="@svg:stroke-opacity" />
								</xsl:call-template>
							</xsl:when>
						</xsl:choose>
					</xsl:for-each>
				</xsl:when>
			</xsl:choose>

			<!-- Dash type-->
			<xsl:if test ="(@draw:stroke='dash')">
				<a:prstDash>
					<xsl:attribute name ="val">
						<xsl:call-template name ="getDashType">
							<xsl:with-param name ="stroke-dash" select ="@draw:stroke-dash" />
						</xsl:call-template>
					</xsl:attribute>
				</a:prstDash>
			</xsl:if>

			<!-- Line join type-->
			<xsl:if test ="@draw:stroke-linejoin">
				<xsl:call-template name ="getJoinType">
					<xsl:with-param name ="stroke-linejoin" select ="@draw:stroke-linejoin" />
				</xsl:call-template>
			</xsl:if>

			<!--Arrow type-->
			<xsl:if test="(@draw:marker-start) and (@draw:marker-start != '')">
				<a:headEnd>
					<xsl:attribute name ="type">
						<xsl:call-template name ="getArrowType">
							<xsl:with-param name ="ArrowType" select ="@draw:marker-start" />
						</xsl:call-template>
					</xsl:attribute>
					<xsl:if test ="@draw:marker-start-width">
						<xsl:call-template name ="setArrowSize">
							<xsl:with-param name ="size" select ="substring-before(@draw:marker-start-width,'cm')" />
						</xsl:call-template >
					</xsl:if>
				</a:headEnd>
			</xsl:if>

			<xsl:if test="(@draw:marker-end) and (@draw:marker-end != '')">
				<a:tailEnd>
					<xsl:attribute name ="type">
						<xsl:call-template name ="getArrowType">
							<xsl:with-param name ="ArrowType" select ="@draw:marker-end" />
						</xsl:call-template>
					</xsl:attribute>
				
					<xsl:if test ="@draw:marker-end-width">
						<xsl:call-template name ="setArrowSize">
							<xsl:with-param name ="size" select ="substring-before(@draw:marker-end-width,'cm')" />
						</xsl:call-template >
					</xsl:if>
				</a:tailEnd>
			</xsl:if>
		</a:ln>
	</xsl:template>

	<xsl:template name ="getFillColor">
		<xsl:param name ="fill-color" />
		<xsl:param name ="opacity" />
		<a:solidFill>
			<a:srgbClr>
				<xsl:attribute name ="val">
					<xsl:value-of select ="substring-after($fill-color,'#')"/>
				</xsl:attribute>
				<xsl:if test ="$opacity != ''">
					<a:alpha>
						<xsl:variable name ="alpha" select ="substring-before($opacity,'%')" />
						<xsl:attribute name ="val">
							<xsl:if test ="$alpha = 0">
								<xsl:value-of select ="0000"/>
							</xsl:if>
							<xsl:if test ="$alpha != 0">
								<xsl:value-of select ="$alpha * 1000"/>
							</xsl:if>
						</xsl:attribute>
					</a:alpha>
				</xsl:if>
			</a:srgbClr>
		</a:solidFill>
	</xsl:template>
	<!-- Get shade for fill reference-->
	<xsl:template name ="getShade">
		<xsl:param name ="gr" />
		<xsl:if test ="document('content.xml')/office:document-content/office:automatic-styles/style:style[@style:name=$gr]/style:graphic-properties/@draw:shadow-opacity">
			<xsl:variable name ="shade">
				<xsl:value-of select ="document('content.xml')/office:document-content/office:automatic-styles/style:style[@style:name=$gr]/style:graphic-properties/@draw:shadow-opacity"/>
			</xsl:variable>
			<a:shade>
				<xsl:variable name ="alpha" select ="substring-before($shade,'%')" />
				<xsl:attribute name ="val">
					<xsl:if test ="$alpha = 0">
						<xsl:value-of select ="0000"/>
					</xsl:if>
					<xsl:if test ="$alpha != 0">
						<xsl:value-of select ="$alpha * 1000"/>
					</xsl:if>
				</xsl:attribute>
			</a:shade>
		</xsl:if>
	</xsl:template>

	<!-- Get arrow type-->
	<xsl:template name ="setArrowSize">
		<xsl:param name ="size" />

		<xsl:choose>
			<xsl:when test ="($size &lt; 0.15)">
				<xsl:attribute name ="w">
					<xsl:value-of select ="'sm'"/>
				</xsl:attribute>
				<xsl:attribute name ="len">
					<xsl:value-of select ="'sm'"/>
				</xsl:attribute>
			</xsl:when>
			<xsl:when test ="($size &gt; 0.15) and ($size &lt;= 0.18)">
				<xsl:attribute name ="w">
					<xsl:value-of select ="'sm'"/>
				</xsl:attribute>
				<xsl:attribute name ="len">
					<xsl:value-of select ="'med'"/>
				</xsl:attribute>
			</xsl:when>
			<xsl:when test ="($size &gt; 0.18) and ($size &lt;= 0.2)">
				<xsl:attribute name ="w">
					<xsl:value-of select ="'sm'"/>
				</xsl:attribute>
				<xsl:attribute name ="len">
					<xsl:value-of select ="'lg'"/>
				</xsl:attribute>
			</xsl:when>
			<xsl:when test ="($size &gt; 0.2) and ($size &lt;= 0.25)">
				<xsl:attribute name ="w">
					<xsl:value-of select ="'med'"/>
				</xsl:attribute>
				<xsl:attribute name ="len">
					<xsl:value-of select ="'sm'"/>
				</xsl:attribute>
			</xsl:when>
			<xsl:when test ="($size &gt; 0.25) and ($size &lt;= 0.3)">
				<xsl:attribute name ="w">
					<xsl:value-of select ="'med'"/>
				</xsl:attribute>
				<xsl:attribute name ="len">
					<xsl:value-of select ="'lg'"/>
				</xsl:attribute>
			</xsl:when>
			<xsl:when test ="($size &gt; 0.3) and ($size &lt;= 0.35)">
				<xsl:attribute name ="w">
					<xsl:value-of select ="'lg'"/>
				</xsl:attribute>
				<xsl:attribute name ="len">
					<xsl:value-of select ="'sm'"/>
				</xsl:attribute>
			</xsl:when>
			<xsl:when test ="($size &gt; 0.35) and ($size &lt;= 0.4)">
				<xsl:attribute name ="w">
					<xsl:value-of select ="'lg'"/>
				</xsl:attribute>
				<xsl:attribute name ="len">
					<xsl:value-of select ="'med'"/>
				</xsl:attribute>
			</xsl:when>
			<xsl:when test ="($size &gt; 0.4) and ($size &lt;= 0.45)">
				<xsl:attribute name ="w">
					<xsl:value-of select ="'lg'"/>
				</xsl:attribute>
				<xsl:attribute name ="len">
					<xsl:value-of select ="'lg'"/>
				</xsl:attribute>
			</xsl:when>
			<xsl:otherwise>
				<xsl:attribute name ="w">
					<xsl:value-of select ="'med'"/>
				</xsl:attribute>
				<xsl:attribute name ="len">
					<xsl:value-of select ="'med'"/>
				</xsl:attribute>
			</xsl:otherwise>
		</xsl:choose>

	</xsl:template>

	<!-- Get line join type-->
	<xsl:template name ="getJoinType">
		<xsl:param name ="stroke-linejoin" />
		<xsl:choose>
			<xsl:when test ="$stroke-linejoin='miter'">
				<a:miter lim="800000" />
			</xsl:when>
			<xsl:when test ="$stroke-linejoin='bevel'">
				<a:bevel/>
			</xsl:when>
			<xsl:when test ="$stroke-linejoin='round'">
				<a:round/>
			</xsl:when>
		</xsl:choose>
	</xsl:template>

	<!-- Get dash type-->
	<xsl:template name ="getDashType">
		<xsl:param name ="stroke-dash" />
		<xsl:for-each select ="document('styles.xml')/office:document-styles/office:styles/draw:stroke-dash[@draw:name=$stroke-dash]">
			<xsl:variable name ="dots1" select="@draw:dots1"/>
			<xsl:variable name ="dots1-length" select ="substring-before(@draw:dots1-length, 'cm')" />
			<xsl:variable name ="dots2" select="@draw:dots2"/>
			<xsl:variable name ="dots2-length" select ="substring-before(@draw:dots2-length, 'cm')"/>
			<xsl:variable name ="distance" select ="substring-before(@draw:distance, 'cm')" />
			<xsl:choose>
				<xsl:when test ="(($dots1='1') and ($dots1-length &lt;= 0.1) and ($dots2='1') and ($dots2-length &lt;= 0.1) and ($distance &lt;= 0.1)) or
								  (($dots1='1') and ($dots1-length &lt;= 0.1) and ($distance &lt;= 0.1)) or
								  (($dots2='1') and ($dots2-length &lt;= 0.1) and ($distance &lt;= 0.1))">
					<xsl:value-of select ="'sysDot'" />
				</xsl:when>
				<xsl:when test ="(($dots1='1') and ($dots1-length &lt;= 0.1) and ($dots2='1') and ($dots2-length &lt;= 0.3) and ($distance &lt;= 0.3)) ">
					<xsl:value-of select ="'dashDot'" />
				</xsl:when>
				<xsl:when test ="(($dots2='1') and ($dots2-length &gt;= 0.1) and ($dots2-length &lt;= 0.3) and ($distance &lt;= 0.3))">
					<xsl:value-of select ="'dash'" />
				</xsl:when>
				<xsl:when test ="(($dots1='1') and ($dots1-length &lt;= 0.1) and ($dots2='1') and ($dots2-length &gt;= 0.3) and ($dots2-length &lt;=0.6) and ($distance &lt;= 0.3)) ">
					<xsl:value-of select ="'lgDashDot'" />
				</xsl:when>
				<xsl:when test ="(($dots1='2') and ($dots1-length &lt;= 0.1) and ($dots2='1') and ($dots2-length &gt;= 0.3) and ($dots2-length &lt;= 0.6) and ($distance &lt;=0.3))">
					<xsl:value-of select ="'lgDashDotDot'" />
				</xsl:when>
				<xsl:when test ="(($dots2='1') and ($dots2-length &gt;= 0.3) and ($dots2-length &lt;= 0.6) and ($distance &lt;= 0.3))">
					<xsl:value-of select ="'lgDash'" />
				</xsl:when>
				<xsl:otherwise>
					<xsl:value-of select ="'sysDash'" />
				</xsl:otherwise>
			</xsl:choose>
		</xsl:for-each >
	</xsl:template>

	<!-- Get text layout for shape-->
	<xsl:template name ="getParagraphProperties">
		<xsl:param name ="gr" />

		<a:bodyPr>
			<!-- Text direction-->
			<xsl:for-each select ="document('content.xml')/office:document-content/office:automatic-styles/style:style[@style:name=$gr]/style:paragraph-properties">
				<xsl:if test ="@style:writing-mode">
					<xsl:attribute name ="vert">
						<xsl:choose>
							<xsl:when test ="@style:writing-mode='tb-rl'">
								<xsl:value-of select ="'vert'"/>
							</xsl:when>
							<xsl:otherwise>
								<xsl:value-of select ="'horz'"/>
							</xsl:otherwise> 
						</xsl:choose>
					</xsl:attribute>
				</xsl:if>
			</xsl:for-each>

			<xsl:for-each select ="document('content.xml')/office:document-content/office:automatic-styles/style:style[@style:name=$gr]/style:graphic-properties">
				<!-- Vertical alignment-->
				<xsl:attribute name ="anchor">
					<xsl:choose>
						<xsl:when test ="@draw:textarea-vertical-align='middle'">
							<xsl:value-of select ="'ctr'"/>
						</xsl:when>
						<xsl:when test ="@draw:textarea-vertical-align='bottom'">
							<xsl:value-of select ="'b'"/>
						</xsl:when>
						<xsl:otherwise>
							<xsl:value-of select ="'t'"/>
						</xsl:otherwise>
					</xsl:choose>
				</xsl:attribute>
				<xsl:attribute name ="anchorCtr">
					<xsl:choose>
						<xsl:when test ="@draw:textarea-horizontal-align='center'">
							<xsl:value-of select ="'1'"/>
						</xsl:when>
						<!--<xsl:when test ="@draw:textarea-horizontal-align='justify'">
						<xsl:value-of select ="0"/>
					</xsl:when>-->
						<xsl:otherwise>
							<xsl:value-of select ="'0'"/>
						</xsl:otherwise>
					</xsl:choose>
				</xsl:attribute>

				
				<!-- Default text style-->
				<xsl:variable name ="parentStyle">
					<xsl:value-of select ="parent::node()/@style:parent-style-name"/>
				</xsl:variable>
				
				<xsl:variable name ="default-padding-left">
					<xsl:call-template name ="getDefaultStyle">
						<xsl:with-param name ="parentStyle" select ="$parentStyle" />
						<xsl:with-param name ="attributeName" select ="'padding-left'" />
					</xsl:call-template> 
				</xsl:variable>
				<xsl:variable name ="default-padding-top">
					<xsl:call-template name ="getDefaultStyle">
						<xsl:with-param name ="parentStyle" select ="$parentStyle" />
						<xsl:with-param name ="attributeName" select ="'padding-top'" />
					</xsl:call-template>
				</xsl:variable>
				<xsl:variable name ="default-padding-right">
					<xsl:call-template name ="getDefaultStyle">
						<xsl:with-param name ="parentStyle" select ="$parentStyle" />
						<xsl:with-param name ="attributeName" select ="'padding-right'" />
					</xsl:call-template>
				</xsl:variable>
				<xsl:variable name ="default-padding-bottom">
					<xsl:call-template name ="getDefaultStyle">
						<xsl:with-param name ="parentStyle" select ="$parentStyle" />
						<xsl:with-param name ="attributeName" select ="'padding-bottom'" />
					</xsl:call-template>
				</xsl:variable>
				<xsl:variable name ="default-wrap-option">
					<xsl:call-template name ="getDefaultStyle">
						<xsl:with-param name ="parentStyle" select ="$parentStyle" />
						<xsl:with-param name ="attributeName" select ="'wrap-option'" />
					</xsl:call-template>
				</xsl:variable>
				<xsl:variable name ="default-autogrow-height">
					<xsl:call-template name ="getDefaultStyle">
						<xsl:with-param name ="parentStyle" select ="$parentStyle" />
						<xsl:with-param name ="attributeName" select ="'auto-grow-height'" />
					</xsl:call-template>
				</xsl:variable>
				<xsl:variable name ="default-autogrow-width">
					<xsl:call-template name ="getDefaultStyle">
						<xsl:with-param name ="parentStyle" select ="$parentStyle" />
						<xsl:with-param name ="attributeName" select ="'auto-grow-width'" />
					</xsl:call-template>
				</xsl:variable>
				
				<!--Text margins in shape-->
				<xsl:if test ="@fo:padding-left">
					<xsl:attribute name ="lIns">
						<xsl:call-template name ="convertToPoints">
							<xsl:with-param name ="unit" select ="'cm'"/>
							<xsl:with-param name ="length" select ="@fo:padding-left"/>
						</xsl:call-template>
					</xsl:attribute>
				</xsl:if>
				<xsl:if test ="not(@fo:padding-left)">
					<xsl:if test ="$default-padding-left != ''">
						<xsl:attribute name ="lIns">
							<xsl:call-template name ="convertToPoints">
								<xsl:with-param name ="unit" select ="'cm'"/>
								<xsl:with-param name ="length" select ="$default-padding-left"/>
							</xsl:call-template>
						</xsl:attribute>
					</xsl:if>
				</xsl:if>
				
				<xsl:if test ="@fo:padding-top">
					<xsl:attribute name ="tIns">
						<xsl:call-template name ="convertToPoints">
							<xsl:with-param name ="unit" select ="'cm'"/>
							<xsl:with-param name ="length" select ="@fo:padding-top"/>
						</xsl:call-template>
					</xsl:attribute>
				</xsl:if>
				<xsl:if test ="not(@fo:padding-top)">
					<xsl:if test ="$default-padding-top != ''">
						<xsl:attribute name ="tIns">
							<xsl:call-template name ="convertToPoints">
								<xsl:with-param name ="unit" select ="'cm'"/>
								<xsl:with-param name ="length" select ="$default-padding-top"/>
							</xsl:call-template>
						</xsl:attribute>
					</xsl:if>
				</xsl:if>
				
				<xsl:if test ="@fo:padding-right">
					<xsl:attribute name ="rIns">
						<xsl:call-template name ="convertToPoints">
							<xsl:with-param name ="unit" select ="'cm'"/>
							<xsl:with-param name ="length" select ="@fo:padding-right"/>
						</xsl:call-template>
					</xsl:attribute>
				</xsl:if>
				<xsl:if test ="not(@fo:padding-right)">
					<xsl:if test ="$default-padding-right != ''">
						<xsl:attribute name ="rIns">
							<xsl:call-template name ="convertToPoints">
								<xsl:with-param name ="unit" select ="'cm'"/>
								<xsl:with-param name ="length" select ="$default-padding-right"/>
							</xsl:call-template>
						</xsl:attribute>
					</xsl:if>
				</xsl:if>
				
				<xsl:if test ="@fo:padding-bottom">
					<xsl:attribute name ="bIns">
						<xsl:call-template name ="convertToPoints">
							<xsl:with-param name ="unit" select ="'cm'"/>
							<xsl:with-param name ="length" select ="@fo:padding-bottom"/>
						</xsl:call-template>
					</xsl:attribute>
				</xsl:if>
				<xsl:if test ="not(@fo:padding-bottom)">
					<xsl:if test ="$default-padding-bottom != ''">
						<xsl:attribute name ="bIns">
							<xsl:call-template name ="convertToPoints">
								<xsl:with-param name ="unit" select ="'cm'"/>
								<xsl:with-param name ="length" select ="$default-padding-bottom"/>
							</xsl:call-template>
						</xsl:attribute>
					</xsl:if>
				</xsl:if>
				
				<!-- Text wrapping in shape-->
				<xsl:choose>
					<xsl:when test ="@fo:wrap-option='no-wrap'">
						<xsl:attribute name ="wrap">
							<xsl:value-of select ="'square'"/>
						</xsl:attribute>
					</xsl:when>
					<xsl:when test ="(@fo:wrap-option='wrap')">
						<xsl:attribute name ="wrap">
							<xsl:value-of select ="'none'"/>
						</xsl:attribute>
					</xsl:when>
					<xsl:when test ="$default-wrap-option='no-wrap'">
						<xsl:attribute name ="wrap">
							<xsl:value-of select ="'square'"/>
						</xsl:attribute>
					</xsl:when>
					<xsl:when test ="$default-wrap-option='wrap'">
						<xsl:attribute name ="wrap">
							<xsl:value-of select ="'none'"/>
						</xsl:attribute>
					</xsl:when>
				</xsl:choose>

				<!-- AutoFit - Resize to accommodate the text in shape.-->
				<xsl:if test ="(@draw:auto-grow-height = 'true') or (@draw:auto-grow-width = 'true') or ($default-autogrow-height = 'true') or ($default-autogrow-width = 'true')">
					<a:spAutoFit/>
				</xsl:if>
			</xsl:for-each>
		</a:bodyPr>

	</xsl:template>
	<!-- Get arriw type-->
	<xsl:template name ="getArrowType">
		<xsl:param name ="ArrowType" />

		<xsl:if test ="document('styles.xml')/office:document-styles/office:styles/draw:marker[@draw:name=$ArrowType]">
			<xsl:variable name ="marker" select ="document('styles.xml')/office:document-styles/office:styles/draw:marker[@draw:name=$ArrowType]/@svg:d"/>

			<xsl:choose>
				<!-- Arrow-->
				<xsl:when test ="$marker = 'm0 2108v17 17l12 42 30 34 38 21 43 4 29-8 30-21 25-26 13-34 343-1532 339 1520 13 42 29 34 39 21 42 4 42-12 34-30 21-42v-39-12l-4 4-440-1998-9-42-25-39-38-25-43-8-42 8-38 25-26 39-8 42z'">
					<xsl:value-of select ="'arrow'" />
				</xsl:when>
				<!-- Stealth-->
				<xsl:when test ="$marker = 'm1013 1491 118 89-567-1580-564 1580 114-85 136-68 148-46 161-17 161 13 153 46z'">
					<xsl:value-of select ="'stealth'" />
				</xsl:when>
				<!-- Oval-->
				<xsl:when test ="$marker = 'm462 1118-102-29-102-51-93-72-72-93-51-102-29-102-13-105 13-102 29-106 51-102 72-89 93-72 102-50 102-34 106-9 101 9 106 34 98 50 93 72 72 89 51 102 29 106 13 102-13 105-29 102-51 102-72 93-93 72-98 51-106 29-101 13z'">
					<xsl:value-of select ="'oval'"/>
				</xsl:when>
				<!-- Diamond-->
				<xsl:when test ="$marker = 'm0 564 564 567 567-567-567-564z'">
					<xsl:value-of select ="'diamond'"/>
				</xsl:when>
				<xsl:otherwise>
					<xsl:value-of select ="'triangle'"/>
				</xsl:otherwise>
			</xsl:choose>
		</xsl:if>

	</xsl:template>
	<!-- Text formatting-->
	<xsl:template name ="processShapeText" >
		<xsl:param name ="gr" />

		<xsl:variable name ="parentStyle">
			<xsl:value-of select ="document('content.xml')/office:document-content/office:automatic-styles/style:style[@style:name=$gr]/@style:parent-style-name"/>
		</xsl:variable>
		<xsl:variable name ="defFontSize">
			<xsl:value-of  select ="office:document-styles/office:styles/style:style/style:text-properties/@fo:font-size"/>
		</xsl:variable>

		<xsl:for-each select ="node()[contains(name(), 'text:')]">
			<xsl:choose >
				<xsl:when test ="text:span">
					<a:p xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
						<xsl:call-template name ="paraProperties" >
							<xsl:with-param name ="paraId" >
								<xsl:value-of select ="@text:style-name"/>
							</xsl:with-param >
						</xsl:call-template >
						<xsl:for-each select ="text:span">
							<a:r>
								<a:rPr lang="en-US" smtClean="0">
									<!--Font Size -->
									<xsl:variable name ="textId">
										<xsl:value-of select ="@text:style-name"/>
									</xsl:variable>
									<xsl:if test ="not($textId ='')">
										<xsl:call-template name ="fontStyles">
											<xsl:with-param name ="Tid" select ="$textId" />
											<xsl:with-param name ="prClassName" select ="$parentStyle" />
										</xsl:call-template>
									</xsl:if>
								</a:rPr >
								<a:t>
									<xsl:if test =".=''">
										<xsl:value-of select ="' '"/>
									</xsl:if>
									<xsl:if test ="not(.='')">
										<xsl:value-of select ="."/>
									</xsl:if>
								</a:t>
							</a:r>
							<xsl:if test ="text:line-break">
								<xsl:call-template name ="processBR">
									<xsl:with-param name ="T" select ="@text:style-name" />
								</xsl:call-template>
							</xsl:if>
						</xsl:for-each >
					</a:p>

				</xsl:when>
				<xsl:when test ="text:list-item/text:p/text:span">
					<xsl:for-each select ="text:list-item/text:p">
						<a:p xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
							<xsl:call-template name ="paraProperties" >
								<xsl:with-param name ="paraId" >
									<xsl:value-of select ="@text:style-name"/>
								</xsl:with-param >
							</xsl:call-template >
							<xsl:for-each select ="child::node()[position()]">
								<xsl:choose >
									<xsl:when test ="name()='text:span'">
										<a:r>
											<a:rPr lang="en-US" smtClean="0">
												<!--Font Size -->
												<xsl:variable name ="textId">
													<xsl:value-of select ="@text:style-name"/>
												</xsl:variable>
												<xsl:if test ="not($textId ='')">
													<xsl:call-template name ="fontStyles">
														<xsl:with-param name ="Tid" select ="$textId" />
														<xsl:with-param name ="prClassName" select ="$parentStyle" />
													</xsl:call-template>
												</xsl:if>
											</a:rPr >
											<a:t>
												<xsl:if test =".=''">
													<xsl:value-of select ="' '"/>
												</xsl:if>
												<xsl:if test ="not(.='')">
													<xsl:value-of select ="."/>
												</xsl:if>
											</a:t>
										</a:r>
									</xsl:when >
									<xsl:when test ="not(name()='text:span')">
										<a:r>
											<a:rPr lang="en-US" smtClean="0">
												<!--Font Size -->
												<xsl:variable name ="textId">
													<xsl:value-of select ="@text:style-name"/>
												</xsl:variable>
												<xsl:if test ="not($textId ='')">
													<xsl:call-template name ="fontStyles">
														<xsl:with-param name ="Tid" select ="$textId" />
														<xsl:with-param name ="prClassName" select ="$parentStyle" />
													</xsl:call-template>
												</xsl:if>
											</a:rPr >
											<a:t>
												<xsl:if test =".=''">
													<xsl:value-of select ="' '"/>
												</xsl:if>
												<xsl:if test ="not(.='')">
													<xsl:value-of select ="."/>
												</xsl:if>
											</a:t>
										</a:r>
									</xsl:when >
								</xsl:choose>
							</xsl:for-each>
						</a:p >
					</xsl:for-each >
				</xsl:when>
				<xsl:when test ="text:list-item/text:p">
					<xsl:for-each select ="text:list-item/text:p">
						<a:p  xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
							<xsl:call-template name ="paraProperties" >
								<xsl:with-param name ="paraId" >
									<xsl:value-of select ="@text:style-name"/>
								</xsl:with-param >
							</xsl:call-template >
							<a:r >
								<a:rPr lang="en-US" smtClean="0">
								</a:rPr >
								<a:t>
									<xsl:if test =".=''">
										<xsl:value-of select ="' '"/>
									</xsl:if>
									<xsl:if test ="not(.='')">
										<xsl:value-of select ="."/>
									</xsl:if>
								</a:t>
							</a:r>
						</a:p>
					</xsl:for-each >
				</xsl:when>
				<xsl:otherwise >
					<a:p xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
						<xsl:if test ="@text:style-name">
							<xsl:call-template name ="paraProperties" >
								<xsl:with-param name ="paraId" >
									<xsl:value-of select ="@text:style-name"/>
								</xsl:with-param >
							</xsl:call-template >
						</xsl:if>
						<a:r >
							<a:rPr lang="en-US" smtClean="0">
								<!--Font Size -->
								<xsl:variable name ="textId">
									<xsl:value-of select ="@text:style-name"/>
								</xsl:variable>
								<xsl:if test ="not($textId ='')">
									<xsl:call-template name ="fontStyles">
										<xsl:with-param name ="Tid" select ="$textId" />
										<xsl:with-param name ="prClassName" select ="$parentStyle" />
									</xsl:call-template>
								</xsl:if>
							</a:rPr >
							<a:t>
								<xsl:if test =".=''">
									<xsl:value-of select ="' '"/>
								</xsl:if>
								<xsl:if test ="not(.='')">
									<xsl:value-of select ="."/>
								</xsl:if>
							</a:t>
						</a:r>
						<xsl:if test ="text:line-break">
							<xsl:call-template name ="processBR">
								<xsl:with-param name ="T" select ="@text:style-name" />
							</xsl:call-template>
						</xsl:if>
					</a:p >
				</xsl:otherwise>
			</xsl:choose>

		</xsl:for-each >

	</xsl:template>
	<xsl:template name ="processBR">
		<xsl:param name ="T" />
		<a:br>
			<a:rPr lang="en-US" smtClean="0">
				<!--Font Size -->
				<xsl:if test ="not($T ='')">
					<xsl:call-template name ="fontStyles">
						<xsl:with-param name ="Tid" select ="$T" />
					</xsl:call-template>
				</xsl:if>
			</a:rPr >

		</a:br>
	</xsl:template>

	<xsl:template name ="getDefaultStyle">
		<xsl:param name ="parentStyle" />
		<xsl:param name ="attributeName" />
		<xsl:for-each select ="document('styles.xml')/office:document-styles/office:styles/style:style[@style:name = $parentStyle]/style:graphic-properties">
			<xsl:choose>
				<xsl:when test ="$attributeName = 'fill'">
					<xsl:value-of select ="@draw:fill"/>
				</xsl:when>
				<xsl:when test ="$attributeName = 'padding-left'">
					<xsl:value-of select ="@fo:padding-left"/>
				</xsl:when>
				<xsl:when test ="$attributeName = 'padding-top'">
					<xsl:value-of select ="@fo:padding-top"/>
				</xsl:when>
				<xsl:when test ="$attributeName = 'padding-right'">
					<xsl:value-of select ="@fo:padding-right"/>
				</xsl:when>
				<xsl:when test ="$attributeName = 'padding-bottom'">
					<xsl:value-of select ="@fo:padding-bottom"/>
				</xsl:when>
				<xsl:when test ="$attributeName = 'wrap-option'">
					<xsl:value-of select ="@fo:wrap-option"/>
				</xsl:when>
				<xsl:when test ="$attributeName = 'auto-grow-height'">
					<xsl:value-of select ="@draw:auto-grow-height"/>
				</xsl:when>
				<xsl:when test ="$attributeName = 'auto-grow-width'">
					<xsl:value-of select ="@draw:auto-grow-width"/>
				</xsl:when>
			</xsl:choose> 
		</xsl:for-each>
	</xsl:template>
</xsl:stylesheet >