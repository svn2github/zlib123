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
					<xsl:with-param name ="shapeName" select="'Rectangle '" />
				</xsl:call-template>
			</xsl:if>
		</xsl:for-each>
		
		<!-- Text box-->
		<xsl:for-each select ="draw:frame">
			<xsl:if test ="not(@presentation:class)">
				<xsl:call-template name ="CreateShape">
					<xsl:with-param name ="shapeName" select="'TextBox '" />
				</xsl:call-template>
			</xsl:if>
		</xsl:for-each>
		<!-- Text box - Custom shape-->
		<xsl:for-each select ="draw:custom-shape">
			<xsl:if test ="(draw:enhanced-geometry/@draw:type='mso-spt202') and (draw:enhanced-geometry/@draw:enhanced-path='M 0 0 L 21600 0 21600 21600 0 21600 0 0 Z N')">
				<xsl:call-template name ="CreateShape">
					<xsl:with-param name ="shapeName" select="'TextBox '" />
				</xsl:call-template>
			</xsl:if>
		</xsl:for-each>
		
		<!-- Line -->
		<xsl:for-each select ="draw:line">
			<xsl:call-template name ="CreateShape">
				<xsl:with-param name ="shapeName" select="concat('Straight ','Connector ')" />
			</xsl:call-template>
		</xsl:for-each>
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

				<xsl:if test ="not(contains($shapeName, 'Straight Connector'))">
					<xsl:call-template name ="drawShape" />
				</xsl:if >
				<xsl:if test ="contains($shapeName, 'Straight Connector')">
					<xsl:call-template name ="drawLine" />
				</xsl:if >
				<xsl:call-template name ="getPresetGeom">
					<xsl:with-param name ="prstGeom" select ="$shapeName" />
				</xsl:call-template>
				
				<xsl:call-template name ="getGraphicProperties">
					<xsl:with-param name ="gr" select="@draw:style-name" />
				</xsl:call-template>
			</p:spPr>
			<p:style>
				<a:lnRef idx="2">
					<a:schemeClr val="accent1">
						<a:shade val="50000"/>
					</a:schemeClr>
				</a:lnRef>
				<a:fillRef idx="1">
					<a:schemeClr val="accent1">
						<xsl:call-template name ="getShade">
							<xsl:with-param name ="gr" select ="@draw:style-name" />
						</xsl:call-template>
					</a:schemeClr>
				</a:fillRef>
				<a:effectRef idx="0">
					<a:schemeClr val="accent1"/>
				</a:effectRef>
				<a:fontRef idx="minor">
					<a:schemeClr val="dk1"/>
				</a:fontRef>
			</p:style>
			<xsl:if test ="not(contains($shapeName, 'Straight Connector'))">
				<p:txBody>
					<xsl:call-template name ="getBodyProperties">
						<xsl:with-param name ="gr" select="@draw:style-name" />
					</xsl:call-template>

					<xsl:choose>
						<xsl:when test ="contains($shapeName, 'TextBox')">
							<xsl:for-each select ="draw:text-box">
								<xsl:call-template name ="processShapeText" />
							</xsl:for-each>
						</xsl:when>
						<xsl:otherwise>
							<xsl:call-template name ="processShapeText" />
						</xsl:otherwise> 
					</xsl:choose>
				</p:txBody>
			</xsl:if>
		</p:sp>
	</xsl:template>
	<!-- Get graphic properties for shape in context.xml-->
	<xsl:template name ="getGraphicProperties">
		<xsl:param name ="gr" />
		<xsl:for-each select ="document('content.xml')/office:document-content/office:automatic-styles/style:style[@style:name=$gr]/style:graphic-properties">
			 <!--FILL--> 
			<xsl:call-template name ="getFillDetails" />
			
			 <!--LINE COLOR AND STYLE-->
			<xsl:call-template name ="getLineStyle" />
				
		</xsl:for-each>
					 
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
		<xsl:choose>
			<xsl:when test ="@draw:fill='none'">
				<a:noFill/>
			</xsl:when>
			<xsl:when test="@draw:fill-color">
				<a:solidFill>
					<a:srgbClr>
						<xsl:attribute name ="val">
							<xsl:value-of select ="substring-after(@draw:fill-color,'#')"/>
						</xsl:attribute>
						<xsl:if test ="@draw:opacity">
							<a:alpha>
								<xsl:attribute name ="val">
									<xsl:value-of select ="concat(substring(@draw:opacity,1,2), '000')"/>
								</xsl:attribute>
							</a:alpha>
						</xsl:if>

					</a:srgbClr>
				</a:solidFill>
			</xsl:when>
		</xsl:choose>
	</xsl:template>
	<!-- Get line color and style-->
	<xsl:template name ="getLineStyle">
		<a:ln>
		<!-- Border width -->
		<xsl:if test ="@svg:stroke-width">
			<xsl:attribute name ="w">
				<xsl:call-template name ="convertToPoints">
					<xsl:with-param name="length"  select ="@svg:stroke-width"/>
					<xsl:with-param name ="unit" select ="'cm'"/>
				</xsl:call-template>
			</xsl:attribute>
		</xsl:if>

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
		
		<xsl:choose>
			<xsl:when test ="@draw:stroke='none'">
				<a:noFill />
			</xsl:when>
			<xsl:when test ="@svg:stroke-color">
				<!-- Line color -->
				<a:solidFill>
					<a:srgbClr>
						<xsl:attribute name ="val">
							<xsl:value-of select ="substring-after(@svg:stroke-color, '#')"/>
						</xsl:attribute>
						<xsl:if test ="@svg:stroke-opacity">
							<a:alpha>
								<xsl:attribute name ="val">
									<xsl:value-of select ="concat(substring(@svg:stroke-opacity,1,2), '000')"/>
								</xsl:attribute>
							</a:alpha>
						</xsl:if>
					</a:srgbClr>
				</a:solidFill>
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
		<xsl:if test="@draw:marker-start">
			<a:headEnd>
				<xsl:call-template name ="getArrowType">
					<xsl:with-param name ="ArrowType" select ="@draw:marker-start" />
				</xsl:call-template>
			</a:headEnd>
		</xsl:if>

		<xsl:if test="@draw:marker-end">
			<a:tailEnd>
				<xsl:call-template name ="getArrowType">
					<xsl:with-param name ="ArrowType" select ="@draw:marker-end" />
				</xsl:call-template>
			</a:tailEnd>
		</xsl:if>
	</a:ln>
	</xsl:template>
	<!-- Get shade for fill reference-->
	<xsl:template name ="getShade">
		<xsl:param name ="gr" />
		<xsl:if test ="document('content.xml')/office:document-content/office:automatic-styles/style:style[@style:name=$gr]/style:graphic-properties/@draw:shadow-opacity">
			<xsl:variable name ="shade">
				<xsl:value-of select ="document('content.xml')/office:document-content/office:automatic-styles/style:style[@style:name=$gr]/style:graphic-properties/@draw:shadow-opacity"/>
			</xsl:variable>
 			<a:shade>
				<xsl:attribute name ="val">
					<xsl:value-of select ="concat(substring($shade,1,2), '000')"/>
				</xsl:attribute>
			</a:shade>
		</xsl:if>
	</xsl:template>
	<!-- Get arrow type-->
	<xsl:template name ="drawArrowType">
		<xsl:param name ="type" />
		<xsl:param name ="w" />
		<xsl:param name ="len" />

		<xsl:attribute name ="type">
			<xsl:value-of select ="$type"/>
		</xsl:attribute>
		<xsl:attribute name ="w">
			<xsl:value-of select ="$w"/>
		</xsl:attribute>
		<xsl:attribute name ="len">
			<xsl:value-of select ="$len"/>
		</xsl:attribute>
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
			 <xsl:choose>
				 <xsl:when test ="((@draw:dots1='1') and (@draw:dots1-length='197%') and (@draw:dots2='1') and (@draw:dots2-length='34%') and (@draw:distance='169%')) or
								  ((@draw:dots1='1') and (@draw:dots1-length='0.07cm') and (@draw:dots2='1') and (@draw:dots2-length='0.07cm') and (@draw:distance='0.07cm')) or
								  ((@draw:dots1='1') and (@draw:dots1-length='0.282cm') and (@draw:distance='0.07cm'))">
					 <xsl:value-of select ="'sysDot'" />
				 </xsl:when>
				 <xsl:when test ="((@draw:dots1='1') and (@draw:dots1-length='0.07cm') and (@draw:dots2='1') and (@draw:dots2-length='0.282cm') and (@draw:distance='0.211cm')) ">
					 <xsl:value-of select ="'dashDot'" />
				 </xsl:when>
				 <xsl:when test ="((@draw:dots2='1') and (@draw:dots2-length='0.282cm') and (@draw:distance='0.211cm'))">
					 <xsl:value-of select ="'dash'" />
				 </xsl:when>
				 <xsl:when test ="((@draw:dots1='1') and (@draw:dots1-length='0.07cm') and (@draw:dots2='1') and (@draw:dots2-length='0.564cm') and (@draw:distance='0.211cm')) ">
					 <xsl:value-of select ="'lgDashDot'" />
				 </xsl:when>
				 <xsl:when test ="((@draw:dots1='2') and (@draw:dots1-length='0.07cm') and (@draw:dots2='1') and (@draw:dots2-length='0.564cm') and (@draw:distance='0.211cm'))">
					 <xsl:value-of select ="'lgDashDotDot'" />
				 </xsl:when>
				 <xsl:when test ="((@draw:dots2='1') and (@draw:dots2-length='0.564cm') and (@draw:distance='0.211cm'))">
					 <xsl:value-of select ="'lgDash'" />
				 </xsl:when>
				  <xsl:otherwise>
					 <xsl:value-of select ="'sysDash'" />
				 </xsl:otherwise> 
			</xsl:choose> 
		</xsl:for-each >	
	</xsl:template>
	<!-- Get text layout for shape-->
	<xsl:template name ="getBodyProperties">
		<xsl:param name ="gr" />
		
		<xsl:for-each select ="document('content.xml')/office:document-content/office:automatic-styles/style:style[@style:name=$gr]/style:graphic-properties">
			<a:bodyPr>
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

			<!--Text margins in shape-->
			<xsl:if test ="fo:padding-left">
				<xsl:attribute name ="lIns">
					<xsl:call-template name ="convertToPoints">
						<xsl:with-param name ="unit" select ="'cm'"/>
						<xsl:with-param name ="length" select ="@fo:padding-left"/>
					</xsl:call-template>
				</xsl:attribute>
			</xsl:if>
			<xsl:if test ="fo:padding-top">
				<xsl:attribute name ="tIns">
					<xsl:call-template name ="convertToPoints">
						<xsl:with-param name ="unit" select ="'cm'"/>
						<xsl:with-param name ="length" select ="@fo:padding-top"/>
					</xsl:call-template>
				</xsl:attribute>
			</xsl:if>
			<xsl:if test ="fo:padding-right">
				<xsl:attribute name ="rIns">
					<xsl:call-template name ="convertToPoints">
						<xsl:with-param name ="unit" select ="'cm'"/>
						<xsl:with-param name ="length" select ="@fo:padding-right"/>
					</xsl:call-template>
				</xsl:attribute>
			</xsl:if>
			<xsl:if test ="fo:padding-bottom">
				<xsl:attribute name ="bIns">
					<xsl:call-template name ="convertToPoints">
						<xsl:with-param name ="unit" select ="'cm'"/>
						<xsl:with-param name ="length" select ="@fo:padding-bottom"/>
					</xsl:call-template>
				</xsl:attribute>
			</xsl:if>
			
			<!-- Text wrapping in shape-->
			<xsl:choose>
				<xsl:when test ="@fo:wrap-option='no-wrap'">
					<xsl:attribute name ="wrap">
						<xsl:value-of select ="'none'"/>
					</xsl:attribute>
				</xsl:when>
				<xsl:when test ="@fo:wrap-option='wrap'">
					<xsl:attribute name ="wrap">
						<xsl:value-of select ="'square'"/>
					</xsl:attribute>
				</xsl:when>
			</xsl:choose> 
				
			</a:bodyPr>
			<a:lstStyle/>
			<a:p>
				<a:pPr>
					<!-- Horizontal alignment-->
					<xsl:attribute name ="algn">
						<xsl:choose>
							<xsl:when test ="@draw:textarea-horizontal-align='center'">
								<xsl:value-of select ="'ctr'"/>
							</xsl:when>
							<xsl:when test ="@draw:textarea-horizontal-align='right'">
								<xsl:value-of select ="'r'"/>
							</xsl:when>
							<xsl:when test ="@draw:textarea-horizontal-align='justify'">
								<xsl:value-of select ="'just'"/>
							</xsl:when>
							<xsl:otherwise>
								<xsl:value-of select ="'l'"/>
							</xsl:otherwise> 
						</xsl:choose>
					</xsl:attribute>
				</a:pPr>
				
			</a:p>
		</xsl:for-each>
	</xsl:template>
	<!-- Get arriw type-->
	<xsl:template name ="getArrowType">
		<xsl:param name ="ArrowType" />

		<xsl:if test ="document('styles.xml')/office:document-styles/office:styles/draw:marker[@draw:name=$ArrowType]">
			<xsl:variable name ="marker" select ="document('styles.xml')/office:document-styles/office:styles/draw:marker[@draw:name=$ArrowType]/@svg:d"/>

			<xsl:choose>
				<!-- Triangle-->
				<xsl:when test ="$marker = 'm70 0 70 140h-140z'">
					<xsl:call-template name ="drawArrowType">
						<xsl:with-param name ="type" select ="'triangle'" />
						<xsl:with-param name ="w" select ="'sm'" />
						<xsl:with-param name ="len" select ="'sm'" />
					</xsl:call-template>
				</xsl:when>
				<xsl:when test ="$marker = 'm70 0 70 210h-140z'">
					<xsl:call-template name ="drawArrowType">
						<xsl:with-param name ="type" select ="'triangle'" />
						<xsl:with-param name ="w" select ="'sm'" />
						<xsl:with-param name ="len" select ="'med'" />
					</xsl:call-template>
				</xsl:when>
				<xsl:when test ="$marker = 'm70 0 70 350h-140z'">
					<xsl:call-template name ="drawArrowType">
						<xsl:with-param name ="type" select ="'triangle'" />
						<xsl:with-param name ="w" select ="'sm'" />
						<xsl:with-param name ="len" select ="'lg'" />
					</xsl:call-template>
				</xsl:when>
				<xsl:when test ="$marker = 'm105 0 105 140h-210z'">
					<xsl:call-template name ="drawArrowType">
						<xsl:with-param name ="type" select ="'triangle'" />
						<xsl:with-param name ="w" select ="'med'" />
						<xsl:with-param name ="len" select ="'sm'" />
					</xsl:call-template>
				</xsl:when>
				<xsl:when test ="$marker = 'm105 0 105 210h-210z'">
					<xsl:call-template name ="drawArrowType">
						<xsl:with-param name ="type" select ="'triangle'" />
						<xsl:with-param name ="w" select ="'med'" />
						<xsl:with-param name ="len" select ="'med'" />
					</xsl:call-template>
				</xsl:when>
				<xsl:when test ="$marker = 'm105 0 105 350h-210z'">
					<xsl:call-template name ="drawArrowType">
						<xsl:with-param name ="type" select ="'triangle'" />
						<xsl:with-param name ="w" select ="'med'" />
						<xsl:with-param name ="len" select ="'lg'" />
					</xsl:call-template>
				</xsl:when>
				<xsl:when test ="$marker = 'm175 0 175 140h-350z'">
					<xsl:call-template name ="drawArrowType">
						<xsl:with-param name ="type" select ="'triangle'" />
						<xsl:with-param name ="w" select ="'lg'" />
						<xsl:with-param name ="len" select ="'sm'" />
					</xsl:call-template>
				</xsl:when>
				<xsl:when test ="$marker = 'm175 0 175 210h-350z'">
					<xsl:call-template name ="drawArrowType">
						<xsl:with-param name ="type" select ="'triangle'" />
						<xsl:with-param name ="w" select ="'lg'" />
						<xsl:with-param name ="len" select ="'med'" />
					</xsl:call-template>
				</xsl:when>
				<xsl:when test ="$marker = 'm175 0 175 350h-350z'">
					<xsl:call-template name ="drawArrowType">
						<xsl:with-param name ="type" select ="'triangle'" />
						<xsl:with-param name ="w" select ="'lg'" />
						<xsl:with-param name ="len" select ="'lg'" />
					</xsl:call-template>
				</xsl:when>

				<!-- Arrow-->
				<xsl:when test ="$marker = 'm122 0 123 222-37 23-86-157-86 157-36-23z'">
					<xsl:call-template name ="drawArrowType">
						<xsl:with-param name ="type" select ="'arrow'" />
						<xsl:with-param name ="w" select ="'sm'" />
						<xsl:with-param name ="len" select ="'sm'" />
					</xsl:call-template>
				</xsl:when>
				<xsl:when test ="$marker = 'm122 0 123 286-37 29-86-202-86 202-36-29z'">
					<xsl:call-template name ="drawArrowType">
						<xsl:with-param name ="type" select ="'arrow'" />
						<xsl:with-param name ="w" select ="'sm'" />
						<xsl:with-param name ="len" select ="'med'" />
					</xsl:call-template>
				</xsl:when>
				<xsl:when test ="$marker = 'm122 0 123 382-37 38-86-269-86 269-36-38z'">
					<xsl:call-template name ="drawArrowType">
						<xsl:with-param name ="type" select ="'arrow'" />
						<xsl:with-param name ="w" select ="'sm'" />
						<xsl:with-param name ="len" select ="'lg'" />
					</xsl:call-template>
				</xsl:when>
				<xsl:when test ="$marker = 'm157 0 158 222-48 23-110-157-110 157-47-23z'">
					<xsl:call-template name ="drawArrowType">
						<xsl:with-param name ="type" select ="'arrow'" />
						<xsl:with-param name ="w" select ="'med'" />
						<xsl:with-param name ="len" select ="'sm'" />
					</xsl:call-template>
				</xsl:when>
				<xsl:when test ="$marker = 'm157 0 158 286-48 29-110-202-110 202-47-29z'">
					<xsl:call-template name ="drawArrowType">
						<xsl:with-param name ="type" select ="'arrow'" />
						<xsl:with-param name ="w" select ="'med'" />
						<xsl:with-param name ="len" select ="'med'" />
					</xsl:call-template>
				</xsl:when>
				<xsl:when test ="$marker = 'm157 0 158 382-48 38-110-269-110 269-47-38z'">
					<xsl:call-template name ="drawArrowType">
						<xsl:with-param name ="type" select ="'arrow'" />
						<xsl:with-param name ="w" select ="'med'" />
						<xsl:with-param name ="len" select ="'lg'" />
					</xsl:call-template>
				</xsl:when>
				<xsl:when test ="$marker = 'm210 0 210 222-63 23-147-157-147 157-63-23z'">
					<xsl:call-template name ="drawArrowType">
						<xsl:with-param name ="type" select ="'arrow'" />
						<xsl:with-param name ="w" select ="'lg'" />
						<xsl:with-param name ="len" select ="'sm'" />
					</xsl:call-template>
				</xsl:when>
				<xsl:when test ="$marker = 'm210 0 210 286-63 29-147-202-147 202-63-29z'">
					<xsl:call-template name ="drawArrowType">
						<xsl:with-param name ="type" select ="'arrow'" />
						<xsl:with-param name ="w" select ="'lg'" />
						<xsl:with-param name ="len" select ="'med'" />
					</xsl:call-template>
				</xsl:when>
				<xsl:when test ="$marker = 'm210 0 210 382-63 38-147-269-147 269-63-38z'">
					<xsl:call-template name ="drawArrowType">
						<xsl:with-param name ="type" select ="'arrow'" />
						<xsl:with-param name ="w" select ="'lg'" />
						<xsl:with-param name ="len" select ="'lg'" />
					</xsl:call-template>
				</xsl:when>

				<!-- Stealth-->
				<xsl:when test ="$marker = 'm70 0 70 140-70-56-70 56z'">
					<xsl:call-template name ="drawArrowType">
						<xsl:with-param name ="type" select ="'stealth'" />
						<xsl:with-param name ="w" select ="'sm'" />
						<xsl:with-param name ="len" select ="'sm'" />
					</xsl:call-template>
				</xsl:when>
				<xsl:when test ="$marker = 'm70 0 70 210-70-84-70 84z'">
					<xsl:call-template name ="drawArrowType">
						<xsl:with-param name ="type" select ="'stealth'" />
						<xsl:with-param name ="w" select ="'sm'" />
						<xsl:with-param name ="len" select ="'med'" />
					</xsl:call-template>
				</xsl:when>
				<xsl:when test ="$marker = 'm70 0 70 350-70-140-70 140z'">
					<xsl:call-template name ="drawArrowType">
						<xsl:with-param name ="type" select ="'stealth'" />
						<xsl:with-param name ="w" select ="'sm'" />
						<xsl:with-param name ="len" select ="'lg'" />
					</xsl:call-template>
				</xsl:when>
				<xsl:when test ="$marker = 'm105 0 105 140-105-56-105 56z'">
					<xsl:call-template name ="drawArrowType">
						<xsl:with-param name ="type" select ="'stealth'" />
						<xsl:with-param name ="w" select ="'med'" />
						<xsl:with-param name ="len" select ="'sm'" />
					</xsl:call-template>
				</xsl:when>
				<xsl:when test ="$marker = 'm105 0 105 210-105-84-105 84z'">
					<xsl:call-template name ="drawArrowType">
						<xsl:with-param name ="type" select ="'stealth'" />
						<xsl:with-param name ="w" select ="'med'" />
						<xsl:with-param name ="len" select ="'med'" />
					</xsl:call-template>
				</xsl:when>
				<xsl:when test ="$marker = 'm105 0 105 350-105-140-105 140z'">
					<xsl:call-template name ="drawArrowType">
						<xsl:with-param name ="type" select ="'stealth'" />
						<xsl:with-param name ="w" select ="'med'" />
						<xsl:with-param name ="len" select ="'lg'" />
					</xsl:call-template>
				</xsl:when>
				<xsl:when test ="$marker = 'm175 0 175 140-175-56-175 56z'">
					<xsl:call-template name ="drawArrowType">
						<xsl:with-param name ="type" select ="'stealth'" />
						<xsl:with-param name ="w" select ="'lg'" />
						<xsl:with-param name ="len" select ="'sm'" />
					</xsl:call-template>
				</xsl:when>
				<xsl:when test ="$marker = 'm175 0 175 210-175-84-175 84z'">
					<xsl:call-template name ="drawArrowType">
						<xsl:with-param name ="type" select ="'stealth'" />
						<xsl:with-param name ="w" select ="'lg'" />
						<xsl:with-param name ="len" select ="'med'" />
					</xsl:call-template>
				</xsl:when>
				<xsl:when test ="$marker = 'm175 0 175 350-175-140-175 140z'">
					<xsl:call-template name ="drawArrowType">
						<xsl:with-param name ="type" select ="'stealth'" />
						<xsl:with-param name ="w" select ="'lg'" />
						<xsl:with-param name ="len" select ="'lg'" />
					</xsl:call-template>
				</xsl:when>

				<!-- Oval-->
				<xsl:when test ="$marker = 'm140 0c0-38-32-70-70-70-38 0-70 32-70 70 0 38 32 70 70 70 38 0 70-32 70-70z'">
					<xsl:call-template name ="drawArrowType">
						<xsl:with-param name ="type" select ="'oval'" />
						<xsl:with-param name ="w" select ="'sm'" />
						<xsl:with-param name ="len" select ="'sm'" />
					</xsl:call-template>
				</xsl:when>
				<xsl:when test ="$marker = 'm140 0c0-57-32-105-70-105-38 0-70 48-70 105 0 57 32 105 70 105 38 0 70-48 70-105z'">
					<xsl:call-template name ="drawArrowType">
						<xsl:with-param name ="type" select ="'oval'" />
						<xsl:with-param name ="w" select ="'sm'" />
						<xsl:with-param name ="len" select ="'med'" />
					</xsl:call-template>
				</xsl:when>
				<xsl:when test ="$marker = 'm140 0c0-96-32-175-70-175-38 0-70 79-70 175 0 96 32 175 70 175 38 0 70-79 70-175z'">
					<xsl:call-template name ="drawArrowType">
						<xsl:with-param name ="type" select ="'oval'" />
						<xsl:with-param name ="w" select ="'sm'" />
						<xsl:with-param name ="len" select ="'lg'" />
					</xsl:call-template>
				</xsl:when>
				<xsl:when test ="$marker = 'm210 0c0-38-48-70-105-70-57 0-105 32-105 70 0 38 48 70 105 70 57 0 105-32 105-70z'">
					<xsl:call-template name ="drawArrowType">
						<xsl:with-param name ="type" select ="'oval'" />
						<xsl:with-param name ="w" select ="'med'" />
						<xsl:with-param name ="len" select ="'sm'" />
					</xsl:call-template>
				</xsl:when>
				<xsl:when test ="$marker = 'm210 0c0-57-48-105-105-105-57 0-105 48-105 105 0 57 48 105 105 105 57 0 105-48 105-105z'">
					<xsl:call-template name ="drawArrowType">
						<xsl:with-param name ="type" select ="'oval'" />
						<xsl:with-param name ="w" select ="'med'" />
						<xsl:with-param name ="len" select ="'med'" />
					</xsl:call-template>
				</xsl:when>
				<xsl:when test ="$marker = 'm210 0c0-96-48-175-105-175-57 0-105 79-105 175 0 96 48 175 105 175 57 0 105-79 105-175z'">
					<xsl:call-template name ="drawArrowType">
						<xsl:with-param name ="type" select ="'oval'" />
						<xsl:with-param name ="w" select ="'med'" />
						<xsl:with-param name ="len" select ="'lg'" />
					</xsl:call-template>
				</xsl:when>
				<xsl:when test ="$marker = 'm350 0c0-38-79-70-175-70-96 0-175 32-175 70 0 38 79 70 175 70 96 0 175-32 175-70z'">
					<xsl:call-template name ="drawArrowType">
						<xsl:with-param name ="type" select ="'oval'" />
						<xsl:with-param name ="w" select ="'lg'" />
						<xsl:with-param name ="len" select ="'sm'" />
					</xsl:call-template>
				</xsl:when>
				<xsl:when test ="$marker = 'm350 0c0-57-79-105-175-105-96 0-175 48-175 105 0 57 79 105 175 105 96 0 175-48 175-105z'">
					<xsl:call-template name ="drawArrowType">
						<xsl:with-param name ="type" select ="'oval'" />
						<xsl:with-param name ="w" select ="'lg'" />
						<xsl:with-param name ="len" select ="'med'" />
					</xsl:call-template>
				</xsl:when>
				<xsl:when test ="$marker = 'm350 0c0-96-79-175-175-175-96 0-175 79-175 175 0 96 79 175 175 175 96 0 175-79 175-175z'">
					<xsl:call-template name ="drawArrowType">
						<xsl:with-param name ="type" select ="'oval'" />
						<xsl:with-param name ="w" select ="'lg'" />
						<xsl:with-param name ="len" select ="'lg'" />
					</xsl:call-template>
				</xsl:when>

				<!-- Diamond-->
				<xsl:when test ="$marker = 'm70 0 70 70-70 70-70-70z'">
					<xsl:call-template name ="drawArrowType">
						<xsl:with-param name ="type" select ="'diamond'" />
						<xsl:with-param name ="w" select ="'sm'" />
						<xsl:with-param name ="len" select ="'sm'" />
					</xsl:call-template>
				</xsl:when>
				<xsl:when test ="$marker = 'm70 0 70 105-70 105-70-105z'">
					<xsl:call-template name ="drawArrowType">
						<xsl:with-param name ="type" select ="'diamond'" />
						<xsl:with-param name ="w" select ="'sm'" />
						<xsl:with-param name ="len" select ="'med'" />
					</xsl:call-template>
				</xsl:when>
				<xsl:when test ="$marker = 'm70 0 70 175-70 175-70-175z'">
					<xsl:call-template name ="drawArrowType">
						<xsl:with-param name ="type" select ="'diamond'" />
						<xsl:with-param name ="w" select ="'sm'" />
						<xsl:with-param name ="len" select ="'lg'" />
					</xsl:call-template>
				</xsl:when>
				<xsl:when test ="$marker = 'm105 0 105 70-105 70-105-70z'">
					<xsl:call-template name ="drawArrowType">
						<xsl:with-param name ="type" select ="'diamond'" />
						<xsl:with-param name ="w" select ="'med'" />
						<xsl:with-param name ="len" select ="'sm'" />
					</xsl:call-template>
				</xsl:when>
				<xsl:when test ="$marker = 'm105 0 105 105-105 105-105-105z'">
					<xsl:call-template name ="drawArrowType">
						<xsl:with-param name ="type" select ="'diamond'" />
						<xsl:with-param name ="w" select ="'med'" />
						<xsl:with-param name ="len" select ="'med'" />
					</xsl:call-template>
				</xsl:when>
				<xsl:when test ="$marker = 'm105 0 105 175-105 175-105-175z'">
					<xsl:call-template name ="drawArrowType">
						<xsl:with-param name ="type" select ="'diamond'" />
						<xsl:with-param name ="w" select ="'med'" />
						<xsl:with-param name ="len" select ="'lg'" />
					</xsl:call-template>
				</xsl:when>
				<xsl:when test ="$marker = 'm175 0 175 70-175 70-175-70z'">
					<xsl:call-template name ="drawArrowType">
						<xsl:with-param name ="type" select ="'diamond'" />
						<xsl:with-param name ="w" select ="'lg'" />
						<xsl:with-param name ="len" select ="'sm'" />
					</xsl:call-template>
				</xsl:when>
				<xsl:when test ="$marker = 'm175 0 175 105-175 105-175-105z'">
					<xsl:call-template name ="drawArrowType">
						<xsl:with-param name ="type" select ="'diamond'" />
						<xsl:with-param name ="w" select ="'lg'" />
						<xsl:with-param name ="len" select ="'med'" />
					</xsl:call-template>
				</xsl:when>
				<xsl:when test ="$marker = 'm175 0 175 175-175 175-175-175z'">
					<xsl:call-template name ="drawArrowType">
						<xsl:with-param name ="type" select ="'diamond'" />
						<xsl:with-param name ="w" select ="'lg'" />
						<xsl:with-param name ="len" select ="'lg'" />
					</xsl:call-template>
				</xsl:when>
				<xsl:otherwise>
					<xsl:call-template name ="drawArrowType">
						<xsl:with-param name ="type" select ="'triangle'" />
						<xsl:with-param name ="w" select ="'med'" />
						<xsl:with-param name ="len" select ="'med'" />
					</xsl:call-template>
				</xsl:otherwise>
			</xsl:choose>
		</xsl:if>

	</xsl:template>
	<!-- Text formatting-->
	<xsl:template name ="processShapeText" >
		<xsl:variable name ="defFontSize">
			<xsl:value-of  select ="office:document-styles/office:styles/style:style/style:text-properties/@fo:font-size"/>
		</xsl:variable>
		 
			<xsl:choose >
				<xsl:when test ="text:p/text:span">
					<xsl:for-each select ="text:p/text:span">
						<a:p xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
							<xsl:call-template name ="paraProperties" >
								<xsl:with-param name ="paraId" >
									<xsl:value-of select ="parent::text:p/@text:style-name"/>
								</xsl:with-param >
							</xsl:call-template >
							<a:r>
								<a:rPr lang="en-US" smtClean="0">
									<!--Font Size -->
									<xsl:variable name ="textId">
										<xsl:value-of select ="@text:style-name"/>
									</xsl:variable>
									<xsl:if test ="not($textId ='')">
										<xsl:call-template name ="fontStyles">
											<xsl:with-param name ="Tid" select ="$textId" />
										</xsl:call-template>
									</xsl:if>
								</a:rPr >
								<a:t>
									<xsl:value-of select ="."/>
								</a:t>
							</a:r>
						</a:p>
					</xsl:for-each >
				</xsl:when>
				<xsl:when test ="text:list/text:list-item/text:p/text:span">
					<xsl:for-each select ="text:list/text:list-item/text:p">
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
													</xsl:call-template>
												</xsl:if>
											</a:rPr >
											<a:t>
												<xsl:value-of select ="."/>
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
													</xsl:call-template>
												</xsl:if>
											</a:rPr >
											<a:t>
												<xsl:value-of select ="."/>
											</a:t>
										</a:r>
									</xsl:when >
								</xsl:choose>
							</xsl:for-each>
						</a:p >
					</xsl:for-each >
				</xsl:when>
				<xsl:when test ="text:list/text:list-item/text:p">
					<xsl:for-each select ="text:list/text:list-item/text:p">
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
									<xsl:value-of select ="."/>
								</a:t>
							</a:r>
						</a:p>
					</xsl:for-each >
				</xsl:when>
				<xsl:when test ="text:p">
					<xsl:for-each select ="text:p">
						<a:p xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
							<xsl:call-template name ="paraProperties" >
								<xsl:with-param name ="paraId" >
									<xsl:value-of select ="@text:style-name"/>
								</xsl:with-param >
							</xsl:call-template >
							<a:r >
								<a:rPr lang="en-US" smtClean="0">
								</a:rPr >
								<a:t>
									<xsl:value-of select ="."/>
								</a:t>
							</a:r>
						</a:p >
					</xsl:for-each >
				</xsl:when>
			</xsl:choose>
			 
	</xsl:template>
	
	
</xsl:stylesheet >