<?xml version="1.0" encoding="UTF-8"?>
<!--
 * Copyright (c) 2006, Clever Age
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
 *     * Neither the name of Clever Age nor the names of its contributors 
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
 * SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.
  -->
<xsl:stylesheet xmlns:xsl="http://www.w3.org/1999/XSL/Transform" version="1.0"
  xmlns:xlink="http://www.w3.org/1999/xlink"
  xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
  xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office"
  xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
  xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0"
  xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0"
  xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0"
  xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
  xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0"
  xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
  xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"
  xmlns:w10="urn:schemas-microsoft-com:office:word"
  exclude-result-prefixes="xlink draw svg fo office style text">

	<xsl:key name="automatic-styles" match="office:automatic-styles/style:style" use="@style:name"/>
	<!--
  *************************************************************************
  SUMMARY
  *************************************************************************
  This stylesheet handles the conversion of shapes. For example shapes are 
  rectangles, ellipses, lines or custom-shapes.
  *************************************************************************
  -->

	<!-- 
  *************************************************************************
  MATCHING TEMPLATES
  *************************************************************************
  -->
	<!-- Sona Shape constants-->
	<xsl:variable name="sm-sm">
		<xsl:value-of select ="'0.15'"/>
	</xsl:variable>
	<xsl:variable name="sm-med">
		<xsl:value-of select ="'0.18'"/>
	</xsl:variable>
	<xsl:variable name="sm-lg">
		<xsl:value-of select ="'0.2'"/>
	</xsl:variable>
	<xsl:variable name="med-sm">
		<xsl:value-of select ="'0.21'" />
	</xsl:variable>
	<xsl:variable name="med-med">
		<xsl:value-of select ="'0.25'"/>
	</xsl:variable>
	<xsl:variable name="med-lg">
		<xsl:value-of select ="'0.3'" />
	</xsl:variable>
	<xsl:variable name="lg-sm">
		<xsl:value-of select ="'0.31'" />
	</xsl:variable>
	<xsl:variable name="lg-med">
		<xsl:value-of select ="'0.35'" />
	</xsl:variable>
	<xsl:variable name="lg-lg">
		<xsl:value-of select ="'0.4'" />
	</xsl:variable>
	<!-- 
  Summary:  Forward shapes in paragraph mode to shapes mode 
  Author:   Clever Age
  -->
	<xsl:template match="draw:custom-shape |draw:rect |draw:ellipse|draw:line|draw:connector" mode="paragraph">
		<!-- COMMENT : many other shapes to be handled by 1.1 -->
		<xsl:choose>
			<xsl:when test="ancestor::draw:text-box">
				<xsl:message terminate="no">translation.odf2oox.nestedFrames</xsl:message>
			</xsl:when>

			<xsl:otherwise>
				<xsl:apply-templates select="." mode="shapes"/>
			</xsl:otherwise>
		</xsl:choose>
	</xsl:template>


	<!-- 
  Summary:  Converts a cutom shape
  Author:   Clever Age
  -->
	<xsl:template match="draw:custom-shape" mode="shapes">
		<xsl:variable name="styleName" select=" @draw:style-name"/>
		<xsl:variable name="automaticStyle" select="key('automatic-styles', $styleName)"/>
		<xsl:variable name="officeStyle" select="document('styles.xml')/office:document-styles/office:styles/style:style[@style:name = $styleName]"/>
		<xsl:variable name="shapeStyle" select="$automaticStyle | $officeStyle" />
		<w:r>
			<w:pict>

				<!--
        makz:
        !!! Detecting a custom shape on their draw:type attribute is no clean solution !!!
        !!! It will only work for ODT files generated by OpenOffice !!!
        !!! This needs to be changed in future !!!
        -->
				<xsl:choose>
					<xsl:when test="draw:enhanced-geometry/@draw:type = 'rectangle' ">

						<xsl:call-template name="InsertRect">
							<xsl:with-param name="shapeStyle" select="$automaticStyle | $officeStyle"/>
							<xsl:with-param name="shape" select="."/>
						</xsl:call-template>

					</xsl:when>
					<xsl:when test="draw:enhanced-geometry/@draw:type = 'ellipse' ">

						<xsl:call-template name="InsertOval">
							<xsl:with-param name="shapeStyle" select="$automaticStyle | $officeStyle"/>
							<xsl:with-param name="shape" select="."/>
						</xsl:call-template>

					</xsl:when>
					<xsl:when test="draw:enhanced-geometry/@draw:enhanced-path = 'M ?f7 0 X 0 ?f8 L 0 ?f9 Y ?f7 21600 L ?f10 21600 X 21600 ?f9 L 21600 ?f8 Y ?f10 0 Z N' ">

						<xsl:call-template name="InsertRoundedRect">
							<xsl:with-param name="shapeStyle" select="$automaticStyle | $officeStyle"/>
							<xsl:with-param name="shape" select="."/>
						</xsl:call-template>

					</xsl:when>
					<!--#############Code changes done by Pradeep#######################-->
					<!-- Isosceles Triangle -->
					<xsl:when test ="draw:enhanced-geometry/@draw:enhanced-path='M ?f0 0 L 21600 21600 0 21600 Z N'">
						<v:shapetype id="_x0000_t5" coordsize="21600,21600" o:spt="5" adj="10800" path="m@0,l,21600r21600,xe">
							<v:stroke joinstyle="miter"/>
							<v:formulas>
								<v:f eqn="val #0"/>
								<v:f eqn="prod #0 1 2"/>
								<v:f eqn="sum @1 10800 0"/>
							</v:formulas>
							<v:path gradientshapeok="t" o:connecttype="custom"
							  o:connectlocs="@0,0;@1,10800;0,21600;10800,21600;21600,21600;@2,10800"
							  textboxrect="0,10800,10800,18000;5400,10800,16200,18000;10800,10800,21600,18000;0,7200,7200,21600;7200,7200,14400,21600;14400,7200,21600,21600"/>
							<v:handles>
								<v:h position="#0,topLeft" xrange="0,21600"/>
							</v:handles>
						</v:shapetype>
						<v:shape id="_x0000_s1027" type="#_x0000_t5" >
							<xsl:call-template name="ConvertShapeProperties">
								<xsl:with-param name="shapeStyle" select="$shapeStyle"/>
								<!-- Sona: Changed I/P parameter-->
								<xsl:with-param name="shape" select="."/>
							</xsl:call-template>
              <!-- Sona:Defect #2020254-->
              <xsl:if test="./text:p/node()">
							<!--insert text-box-->
							<xsl:call-template name="InsertTextBox">
								<!-- Sona: Changed I/P parameter-->
								<xsl:with-param name="frameStyle" select="$shapeStyle"/>
							</xsl:call-template>
              </xsl:if>
						</v:shape>
					</xsl:when>
					<!-- Right Triangle -->
					<xsl:when test ="draw:enhanced-geometry/@draw:enhanced-path='M 0 0 L 21600 21600 0 21600 0 0 Z N'">
						<v:shapetype id="_x0000_t6" coordsize="21600,21600" o:spt="6" path="m,l,21600r21600,xe">
							<v:stroke joinstyle="miter"/>
							<v:path gradientshapeok="t" o:connecttype="custom" o:connectlocs="0,0;0,10800;0,21600;10800,21600;21600,21600;10800,10800" textboxrect="1800,12600,12600,19800"/>
						</v:shapetype>
						<v:shape id="_x0000_s1035" type="#_x0000_t6">
							<xsl:call-template name="ConvertShapeProperties">
								<xsl:with-param name="shapeStyle" select="$shapeStyle"/>
								<xsl:with-param name="shape" select="."/>
							</xsl:call-template>
              <!-- Sona:Defect #2020254-->
              <xsl:if test="./text:p/node()">
							<!--insert text-box-->
							<xsl:call-template name="InsertTextBox">
								<xsl:with-param name="frameStyle" select="$shapeStyle"/>
							</xsl:call-template>
              </xsl:if>
						</v:shape>
					</xsl:when>
					<!--Text-->
					<!--<xsl:when test ="./child::node()[1]=draw:text-box">
            <v:shapetype id="_x0000_t202" coordsize="21600,21600" o:spt="202" path="m,l,21600r21600,l21600,xe">
              <v:stroke joinstyle="miter"/>
              <v:path gradientshapeok="t" o:connecttype="rect"/>
            </v:shapetype>
            <v:shape id="_x0000_s1027" type="#_x0000_t202" style="position:absolute;margin-left:76.5pt;margin-top:152.25pt;width:232.5pt;height:69pt;z-index:251659264">
              -->
					<!--<xsl:call-template name="ConvertShapeProperties">
                <xsl:with-param name="shapeStyle" select="$shapeStyle"/>
                <xsl:with-param name="shape" select="."/>
              </xsl:call-template>-->
					<!--
              -->
					<!--insert text-box-->
					<!--
              <xsl:call-template name="InsertTextBox">
                <xsl:with-param name="frameStyle" select="$shapeStyle"/>
              </xsl:call-template>
            </v:shape>-->

					<!--</xsl:when>-->
					<!-- Flowchart Process-->
					<xsl:when test ="draw:enhanced-geometry/@draw:enhanced-path ='M 0 0 L 21600 0 21600 21600 0 21600 0 0 Z N'">
						<v:shapetype id="_x0000_t109" coordsize="21600,21600" o:spt="109" path="m,l,21600r21600,l21600,xe">
							<v:stroke joinstyle="miter"/>
							<v:path gradientshapeok="t" o:connecttype="rect"/>
						</v:shapetype>
						<v:shape id="_x0000_s1037" type="#_x0000_t109" >
							<xsl:call-template name="ConvertShapeProperties">
								<xsl:with-param name="shapeStyle" select="$shapeStyle"/>
								<xsl:with-param name="shape" select="."/>
							</xsl:call-template>
              <!-- Sona:Defect #2020254-->
              <xsl:if test="./text:p/node()">
							<!-- insert text-box -->
							<xsl:call-template name="InsertTextBox">
								<xsl:with-param name="frameStyle" select="$shapeStyle"/>
							</xsl:call-template>
              </xsl:if>
						</v:shape>
					</xsl:when>
					<!-- Flowchart Alternate Process -->
					<xsl:when test ="draw:enhanced-geometry/@draw:enhanced-path ='M 0 ?f2 Y ?f0 0 L ?f1 0 X 21600 ?f2 L 21600 ?f3 Y ?f1 21600 L ?f0 21600 X 0 ?f3 Z N'">
						<v:shapetype id="_x0000_t176" coordsize="21600,21600" o:spt="176" adj="2700" path="m@0,qx0@0l0@2qy@0,21600l@1,21600qx21600@2l21600@0qy@1,xe">
							<v:stroke joinstyle="miter"/>
							<v:formulas>
								<v:f eqn="val #0"/>
								<v:f eqn="sum width 0 #0"/>
								<v:f eqn="sum height 0 #0"/>
								<v:f eqn="prod @0 2929 10000"/>
								<v:f eqn="sum width 0 @3"/>
								<v:f eqn="sum height 0 @3"/>
								<v:f eqn="val width"/>
								<v:f eqn="val height"/>
								<v:f eqn="prod width 1 2"/>
								<v:f eqn="prod height 1 2"/>
							</v:formulas>
							<v:path gradientshapeok="t" limo="10800,10800" o:connecttype="custom" o:connectlocs="@8,0;0,@9;@8,@7;@6,@9" textboxrect="@3,@3,@4,@5"/>
						</v:shapetype>
						<v:shape id="_x0000_s1038" type="#_x0000_t176" style="position:absolute;margin-left:77.25pt;margin-top:117.75pt;width:153pt;height:63.75pt;z-index:251659264">
							<xsl:call-template name="ConvertShapeProperties">
								<xsl:with-param name="shapeStyle" select="$shapeStyle"/>
								<xsl:with-param name="shape" select="."/>
							</xsl:call-template>
              <!-- Sona:Defect #2020254-->
              <xsl:if test="./text:p/node()">
							<!-- insert text-box -->
							<xsl:call-template name="InsertTextBox">
								<xsl:with-param name="frameStyle" select="$shapeStyle"/>
							</xsl:call-template>
              </xsl:if>
						</v:shape>
					</xsl:when>
					<!--Flowchart Decision -->
					<xsl:when test ="draw:enhanced-geometry/@draw:enhanced-path ='M 0 10800 L 10800 0 21600 10800 10800 21600 0 10800 Z N'">
						<v:shapetype id="_x0000_t110" coordsize="21600,21600" o:spt="110" path="m10800,l,10800,10800,21600,21600,10800xe">
							<v:stroke joinstyle="miter"/>
							<v:path gradientshapeok="t" o:connecttype="rect" textboxrect="5400,5400,16200,16200"/>
						</v:shapetype>
						<v:shape id="_x0000_s1039" type="#_x0000_t110" >
							<xsl:call-template name="ConvertShapeProperties">
								<xsl:with-param name="shapeStyle" select="$shapeStyle"/>
								<xsl:with-param name="shape" select="."/>
							</xsl:call-template>
              <!-- Sona:Defect #2020254-->
              <xsl:if test="./text:p/node()">
							<!--insert text-box-->
							<xsl:call-template name="InsertTextBox">
								<xsl:with-param name="frameStyle" select="$shapeStyle"/>
							</xsl:call-template>
              </xsl:if>
						</v:shape>
					</xsl:when>
					<!-- Flowchart Data -->
					<xsl:when test ="draw:enhanced-geometry/@draw:enhanced-path ='M 4230 0 L 21600 0 17370 21600 0 21600 4230 0 Z N'">
						<v:shapetype id="_x0000_t111" coordsize="21600,21600" o:spt="111" path="m4321,l21600,,17204,21600,,21600xe">
							<v:stroke joinstyle="miter"/>
							<v:path gradientshapeok="t" o:connecttype="custom" o:connectlocs="12961,0;10800,0;2161,10800;8602,21600;10800,21600;19402,10800" textboxrect="4321,0,17204,21600"/>
						</v:shapetype>
						<v:shape id="_x0000_s1040" type="#_x0000_t111" style="position:absolute;margin-left:70.5pt;margin-top:348.75pt;width:172.5pt;height:63.75pt;z-index:251661312">
							<xsl:call-template name="ConvertShapeProperties">
								<xsl:with-param name="shapeStyle" select="$shapeStyle"/>
								<xsl:with-param name="shape" select="."/>
							</xsl:call-template>
              <!-- Sona:Defect #2020254-->
              <xsl:if test="./text:p/node()">
							<!--insert text-box-->
							<xsl:call-template name="InsertTextBox">
								<xsl:with-param name="frameStyle" select="$shapeStyle"/>
							</xsl:call-template>
              </xsl:if>
						</v:shape>
					</xsl:when>
					<!-- Flowchart Predefined Process-->
					<xsl:when test ="draw:enhanced-geometry/@draw:enhanced-path ='M 0 0 L 21600 0 21600 21600 0 21600 Z N M 2540 0 L 2540 21600 N M 19060 0 L 19060 21600 N'">
						<v:shapetype id="_x0000_t112" coordsize="21600,21600" o:spt="112" path="m,l,21600r21600,l21600,xem2610,nfl2610,21600em18990,nfl18990,21600e">
							<v:stroke joinstyle="miter"/>
							<v:path o:extrusionok="f" gradientshapeok="t" o:connecttype="rect" textboxrect="2610,0,18990,21600"/>
						</v:shapetype>
						<v:shape id="_x0000_s1026" type="#_x0000_t112" style="position:absolute;margin-left:79.5pt;margin-top:35.25pt;width:211.5pt;height:101.25pt;z-index:251658240">
							<xsl:call-template name="ConvertShapeProperties">
								<xsl:with-param name="shapeStyle" select="$shapeStyle"/>
								<xsl:with-param name="shape" select="."/>
							</xsl:call-template>
              <!-- Sona:Defect #2020254-->
              <xsl:if test="./text:p/node()">
							<!--insert text-box-->
							<xsl:call-template name="InsertTextBox">
								<xsl:with-param name="frameStyle" select="$shapeStyle"/>
							</xsl:call-template>
              </xsl:if>
						</v:shape>
					</xsl:when>
					<!-- Flowchart Internal Storage-->
					<xsl:when test ="draw:enhanced-geometry/@draw:enhanced-path ='M 0 0 L 21600 0 21600 21600 0 21600 Z N M 4230 0 L 4230 21600 N M 0 4230 L 21600 4230 N'">
						<v:shapetype id="_x0000_t113" coordsize="21600,21600" o:spt="113" path="m,l,21600r21600,l21600,xem4236,nfl4236,21600em,4236nfl21600,4236e">
							<v:stroke joinstyle="miter"/>
							<v:path o:extrusionok="f" gradientshapeok="t" o:connecttype="rect" textboxrect="4236,4236,21600,21600"/>
						</v:shapetype>
						<v:shape id="_x0000_s1039" type="#_x0000_t113" >
							<xsl:call-template name="ConvertShapeProperties">
								<xsl:with-param name="shapeStyle" select="$shapeStyle"/>
								<xsl:with-param name="shape" select="."/>
							</xsl:call-template>
              <!-- Sona:Defect #2020254-->
              <xsl:if test="./text:p/node()">
							<!--insert text-box-->
							<xsl:call-template name="InsertTextBox">
								<xsl:with-param name="frameStyle" select="$shapeStyle"/>
							</xsl:call-template>
              </xsl:if>
						</v:shape>
					</xsl:when>
					<!-- Flowchart Document-->
					<xsl:when test ="draw:enhanced-geometry/@draw:enhanced-path ='M 0 0 L 21600 0 21600 17360 C 13050 17220 13340 20770 5620 21600 2860 21100 1850 20700 0 20120 Z N'">
						<v:shapetype id="_x0000_t114" coordsize="21600,21600" o:spt="114" path="m,20172v945,400,1887,628,2795,913c3587,21312,4342,21370,5060,21597v2037,,2567,-227,3095,-285c8722,21197,9325,20970,9855,20800v490,-228,945,-400,1472,-740c11817,19887,12347,19660,12875,19375v567,-228,1095,-513,1700,-740c15177,18462,15782,18122,16537,17950v718,-113,1398,-398,2228,-513c19635,17437,20577,17322,21597,17322l21597,,,xe">
							<v:stroke joinstyle="miter"/>
							<v:path o:connecttype="custom" o:connectlocs="10800,0;0,10800;10800,20400;21600,10800" textboxrect="0,0,21600,17322"/>
						</v:shapetype>
						<v:shape id="_x0000_s1040" type="#_x0000_t114" >
							<xsl:call-template name="ConvertShapeProperties">
								<xsl:with-param name="shapeStyle" select="$shapeStyle"/>
								<xsl:with-param name="shape" select="."/>
							</xsl:call-template>
              <!-- Sona:Defect #2020254-->
              <xsl:if test="./text:p/node()">
							<!--insert text-box-->
							<xsl:call-template name="InsertTextBox">
								<xsl:with-param name="frameStyle" select="$shapeStyle"/>
							</xsl:call-template>
              </xsl:if>
						</v:shape>
					</xsl:when>
					<!-- Flowchart Multi Document-->
					<xsl:when test ="draw:enhanced-geometry/@draw:enhanced-path ='M 0 3600 L 1500 3600 1500 1800 3000 1800 3000 0 21600 0 21600 14409 20100 14409 20100 16209 18600 16209 18600 18009 C 11610 17893 11472 20839 4833 21528 2450 21113 1591 20781 0 20300 Z N M 1500 3600 F L 18600 3600 18600 16209 N M 3000 1800 F L 20100 1800 20100 14409 N'">
						<v:shapetype id="_x0000_t115" coordsize="21600,21600" o:spt="115" path="m,20465v810,317,1620,452,2397,725c3077,21325,3790,21417,4405,21597v1620,,2202,-180,2657,-272c7580,21280,8002,21010,8455,20917v422,-135,810,-405,1327,-542c10205,20150,10657,19967,11080,19742v517,-182,970,-407,1425,-590c13087,19017,13605,18745,14255,18610v615,-180,1262,-318,1942,-408c16975,18202,17785,18022,18595,18022r,-1670l19192,16252r808,l20000,14467r722,-75l21597,14392,21597,,2972,r,1815l1532,1815r,1860l,3675,,20465xem1532,3675nfl18595,3675r,12677em2972,1815nfl20000,1815r,12652e">
							<v:stroke joinstyle="miter"/>
							<v:path o:extrusionok="f" o:connecttype="custom" o:connectlocs="10800,0;0,10800;10800,19890;21600,10800" textboxrect="0,3675,18595,18022"/>
						</v:shapetype>
						<v:shape id="_x0000_s1041" type="#_x0000_t115" style="position:absolute;margin-left:66.75pt;margin-top:401.25pt;width:208.5pt;height:116.25pt;z-index:251661312">
							<xsl:call-template name="ConvertShapeProperties">
								<xsl:with-param name="shapeStyle" select="$shapeStyle"/>
								<xsl:with-param name="shape" select="."/>
							</xsl:call-template>
              <!-- Sona:Defect #2020254-->
              <xsl:if test="./text:p/node()">
							<!--insert text-box-->
							<xsl:call-template name="InsertTextBox">
								<xsl:with-param name="frameStyle" select="$shapeStyle"/>
							</xsl:call-template>
              </xsl:if>
						</v:shape>
					</xsl:when>
					<!-- Flowchart Terminator-->
					<xsl:when test ="draw:enhanced-geometry/@draw:enhanced-path ='M 3470 21600 X 0 10800 3470 0 L 18130 0 X 21600 10800 18130 21600 Z N'">
						<v:shapetype id="_x0000_t116" coordsize="21600,21600" o:spt="116" path="m3475,qx,10800,3475,21600l18125,21600qx21600,10800,18125,xe">
							<v:stroke joinstyle="miter"/>
							<v:path gradientshapeok="t" o:connecttype="rect" textboxrect="1018,3163,20582,18437"/>
						</v:shapetype>
						<v:shape id="_x0000_s1042" type="#_x0000_t116">
							<xsl:call-template name="ConvertShapeProperties">
								<xsl:with-param name="shapeStyle" select="$shapeStyle"/>
								<xsl:with-param name="shape" select="."/>
							</xsl:call-template>
              <!-- Sona:Defect #2020254-->
              <xsl:if test="./text:p/node()">
							<!--insert text-box-->
							<xsl:call-template name="InsertTextBox">
								<xsl:with-param name="frameStyle" select="$shapeStyle"/>
							</xsl:call-template>
              </xsl:if>
						</v:shape>
					</xsl:when>
					<!-- Flowchart Collate-->
					<xsl:when test ="draw:enhanced-geometry/@draw:enhanced-path ='M 0 0 L 21600 21600 0 21600 21600 0 0 0 Z N'">
						<v:shapetype id="_x0000_t125" coordsize="21600,21600" o:spt="125" path="m21600,21600l,21600,21600,,,xe">
							<v:stroke joinstyle="miter"/>
							<v:path o:extrusionok="f" gradientshapeok="t" o:connecttype="custom" o:connectlocs="10800,0;10800,10800;10800,21600" textboxrect="5400,5400,16200,16200"/>
						</v:shapetype>
						<v:shape id="_x0000_s1026" type="#_x0000_t125">
							<xsl:call-template name="ConvertShapeProperties">
								<xsl:with-param name="shapeStyle" select="$shapeStyle"/>
								<xsl:with-param name="shape" select="."/>
							</xsl:call-template>
              <!-- Sona:Defect #2020254-->
              <xsl:if test="./text:p/node()">
							<!--insert text-box-->
							<xsl:call-template name="InsertTextBox">
								<xsl:with-param name="frameStyle" select="$shapeStyle"/>
							</xsl:call-template>
              </xsl:if>
						</v:shape>

					</xsl:when>
					<!-- Flowchart Sort-->
					<xsl:when test ="draw:enhanced-geometry/@draw:enhanced-path ='M 0 10800 L 10800 0 21600 10800 10800 21600 Z N M 0 10800 L 21600 10800 N'">
						<v:shapetype id="_x0000_t126" coordsize="21600,21600" o:spt="126" path="m10800,l,10800,10800,21600,21600,10800xem,10800nfl21600,10800e">
							<v:stroke joinstyle="miter"/>
							<v:path o:extrusionok="f" gradientshapeok="t" o:connecttype="rect" textboxrect="5400,5400,16200,16200"/>
						</v:shapetype>
						<v:shape id="_x0000_s1027" type="#_x0000_t126">
							<xsl:call-template name="ConvertShapeProperties">
								<xsl:with-param name="shapeStyle" select="$shapeStyle"/>
								<xsl:with-param name="shape" select="."/>
							</xsl:call-template>
              <!-- Sona:Defect #2020254-->
              <xsl:if test="./text:p/node()">
							<!--insert text-box-->
							<xsl:call-template name="InsertTextBox">
								<xsl:with-param name="frameStyle" select="$shapeStyle"/>
							</xsl:call-template>
              </xsl:if>
						</v:shape>
					</xsl:when>
					<!-- Flowchart Extract-->
					<xsl:when test ="draw:enhanced-geometry/@draw:enhanced-path ='M 10800 0 L 21600 21600 0 21600 10800 0 Z N'">
						<v:shapetype id="_x0000_t127" coordsize="21600,21600" o:spt="127" path="m10800,l21600,21600,,21600xe">
							<v:stroke joinstyle="miter"/>
							<v:path gradientshapeok="t" o:connecttype="custom" o:connectlocs="10800,0;5400,10800;10800,21600;16200,10800" textboxrect="5400,10800,16200,21600"/>
						</v:shapetype>
						<v:shape id="_x0000_s1028" type="#_x0000_t127" >
							<xsl:call-template name="ConvertShapeProperties">
								<xsl:with-param name="shapeStyle" select="$shapeStyle"/>
								<xsl:with-param name="shape" select="."/>
							</xsl:call-template>
              <!-- Sona:Defect #2020254-->
              <xsl:if test="./text:p/node()">
							<!--insert text-box-->
							<xsl:call-template name="InsertTextBox">
								<xsl:with-param name="frameStyle" select="$shapeStyle"/>
							</xsl:call-template>
              </xsl:if>
						</v:shape>
					</xsl:when>
					<!--Flowchart Merge-->
					<xsl:when test ="draw:enhanced-geometry/@draw:enhanced-path ='M 0 0 L 21600 0 10800 21600 0 0 Z N'">
						<v:shapetype id="_x0000_t128" coordsize="21600,21600" o:spt="128" path="m,l21600,,10800,21600xe">
							<v:stroke joinstyle="miter"/>
							<v:path gradientshapeok="t" o:connecttype="custom" o:connectlocs="10800,0;5400,10800;10800,21600;16200,10800" textboxrect="5400,0,16200,10800"/>
						</v:shapetype>
						<v:shape id="_x0000_s1029" type="#_x0000_t128" style="position:absolute;margin-left:123pt;margin-top:468.8pt;width:87pt;height:96pt;z-index:251661312">
							<xsl:call-template name="ConvertShapeProperties">
								<xsl:with-param name="shapeStyle" select="$shapeStyle"/>
								<xsl:with-param name="shape" select="."/>
							</xsl:call-template>
              <!-- Sona:Defect #2020254-->
              <xsl:if test="./text:p/node()">
							<!--insert text-box-->
							<xsl:call-template name="InsertTextBox">
								<xsl:with-param name="frameStyle" select="$shapeStyle"/>
							</xsl:call-template>
              </xsl:if>
						</v:shape>
					</xsl:when>
					<!-- Flowchart Stored Data-->
					<xsl:when test ="draw:enhanced-geometry/@draw:enhanced-path ='M 3600 21600 X 0 10800 3600 0 L 21600 0 X 18000 10800 21600 21600 Z N'">
						<v:shapetype id="_x0000_t130" coordsize="21600,21600" o:spt="130" path="m3600,21597c2662,21202,1837,20075,1087,18440,487,16240,75,13590,,10770,75,8007,487,5412,1087,3045,1837,1465,2662,337,3600,l21597,v-937,337,-1687,1465,-2512,3045c18485,5412,18072,8007,17997,10770v75,2820,488,5470,1088,7670c19910,20075,20660,21202,21597,21597xe">
							<v:stroke joinstyle="miter"/>
							<v:path gradientshapeok="t" o:connecttype="custom" o:connectlocs="10800,0;0,10800;10800,21600;17997,10800" textboxrect="3600,0,17997,21600"/>
						</v:shapetype>
						<v:shape id="_x0000_s1030" type="#_x0000_t130" >
							<xsl:call-template name="ConvertShapeProperties">
								<xsl:with-param name="shapeStyle" select="$shapeStyle"/>
								<xsl:with-param name="shape" select="."/>
							</xsl:call-template>
              <!-- Sona:Defect #2020254-->
              <xsl:if test="./text:p/node()">
							<!--insert text-box-->
							<xsl:call-template name="InsertTextBox">
								<xsl:with-param name="frameStyle" select="$shapeStyle"/>
							</xsl:call-template>
              </xsl:if>
						</v:shape>
					</xsl:when>
					<!-- Flowchart Delay-->
					<xsl:when test ="draw:enhanced-geometry/@draw:enhanced-path ='M 10800 0 X 21600 10800 10800 21600 L 0 21600 0 0 Z N'">
						<v:shapetype id="_x0000_t135" coordsize="21600,21600" o:spt="135" path="m10800,qx21600,10800,10800,21600l,21600,,xe">
							<v:stroke joinstyle="miter"/>
							<v:path gradientshapeok="t" o:connecttype="rect" textboxrect="0,3163,18437,18437"/>
						</v:shapetype>
						<v:shape id="_x0000_s1026" type="#_x0000_t135">
							<xsl:call-template name="ConvertShapeProperties">
								<xsl:with-param name="shapeStyle" select="$shapeStyle"/>
								<xsl:with-param name="shape" select="."/>
							</xsl:call-template>
              <!-- Sona:Defect #2020254-->
              <xsl:if test="./text:p/node()">
							<!--insert text-box-->
							<xsl:call-template name="InsertTextBox">
								<xsl:with-param name="frameStyle" select="$shapeStyle"/>
							</xsl:call-template>
              </xsl:if>
						</v:shape>
					</xsl:when>
					<!-- Flowchart Sequential Access-->
					<xsl:when test ="draw:enhanced-geometry/@draw:enhanced-path ='M 20980 18150 L 20980 21600 10670 21600 C 4770 21540 0 16720 0 10800 0 4840 4840 0 10800 0 16740 0 21600 4840 21600 10800 21600 13520 20550 16160 18670 18170 Z N'">
						<v:shapetype id="_x0000_t131" coordsize="21600,21600" o:spt="131" path="ar,,21600,21600,18685,18165,10677,21597l20990,21597r,-3432xe">
							<v:stroke joinstyle="miter"/>
							<v:path o:connecttype="rect" textboxrect="3163,3163,18437,18437"/>
						</v:shapetype>
						<v:shape id="_x0000_s1027" type="#_x0000_t131" style="position:absolute;margin-left:140.25pt;margin-top:132pt;width:118.5pt;height:1in;z-index:251659264">
							<xsl:call-template name="ConvertShapeProperties">
								<xsl:with-param name="shapeStyle" select="$shapeStyle"/>
								<xsl:with-param name="shape" select="."/>
							</xsl:call-template>
              <!-- Sona:Defect #2020254-->
              <xsl:if test="./text:p/node()">
							<!--insert text-box-->
							<xsl:call-template name="InsertTextBox">
								<xsl:with-param name="frameStyle" select="$shapeStyle"/>
							</xsl:call-template>
              </xsl:if>
						</v:shape>
					</xsl:when>
					<!-- Flowchart Megnetic Disk-->
					<xsl:when test ="draw:enhanced-geometry/@draw:enhanced-path ='M 0 3400 Y 10800 0 21600 3400 L 21600 18200 Y 10800 21600 0 18200 Z N M 0 3400 Y 10800 6800 21600 3400 N'">
						<v:shapetype id="_x0000_t132" coordsize="21600,21600" o:spt="132" path="m10800,qx,3391l,18209qy10800,21600,21600,18209l21600,3391qy10800,xem,3391nfqy10800,6782,21600,3391e">
							<v:path o:extrusionok="f" gradientshapeok="t" o:connecttype="custom" o:connectlocs="10800,6782;10800,0;0,10800;10800,21600;21600,10800" o:connectangles="270,270,180,90,0" textboxrect="0,6782,21600,18209"/>
						</v:shapetype>
						<v:shape id="_x0000_s1028" type="#_x0000_t132" >
							<xsl:call-template name="ConvertShapeProperties">
								<xsl:with-param name="shapeStyle" select="$shapeStyle"/>
								<xsl:with-param name="shape" select="."/>
							</xsl:call-template>
              <!-- Sona:Defect #2020254-->
              <xsl:if test="./text:p/node()">
							<!--insert text-box-->
							<xsl:call-template name="InsertTextBox">
								<xsl:with-param name="frameStyle" select="$shapeStyle"/>
							</xsl:call-template>
              </xsl:if>
						</v:shape>
					</xsl:when>
					<!-- Flowchart Direct Access Storage-->
					<xsl:when test ="draw:enhanced-geometry/@draw:enhanced-path ='M 18200 0 X 21600 10800 18200 21600 L 3400 21600 X 0 10800 3400 0 Z N M 18200 0 X 14800 10800 18200 21600 N'">
						<v:shapetype id="_x0000_t133" coordsize="21600,21600" o:spt="133" path="m21600,10800qy18019,21600l3581,21600qx,10800,3581,l18019,qx21600,10800xem18019,21600nfqx14438,10800,18019,e">
							<v:path o:extrusionok="f" gradientshapeok="t" o:connecttype="custom" o:connectlocs="10800,0;0,10800;10800,21600;14438,10800;21600,10800" o:connectangles="270,180,90,0,0" textboxrect="3581,0,14438,21600"/>
						</v:shapetype>
						<v:shape id="_x0000_s1029" type="#_x0000_t133" >
							<xsl:call-template name="ConvertShapeProperties">
								<xsl:with-param name="shapeStyle" select="$shapeStyle"/>
								<xsl:with-param name="shape" select="."/>
							</xsl:call-template>
              <!-- Sona:Defect #2020254-->
              <xsl:if test="./text:p/node()">
							<!--insert text-box-->
							<xsl:call-template name="InsertTextBox">
								<xsl:with-param name="frameStyle" select="$shapeStyle"/>
							</xsl:call-template>
              </xsl:if>
						</v:shape>
					</xsl:when>
					<!--Flowchart Display -->
					<xsl:when test ="draw:enhanced-geometry/@draw:enhanced-path ='M 3600 0 L 17800 0 X 21600 10800 17800 21600 L 3600 21600 0 10800 Z N'">

						<v:shapetype id="_x0000_t134" coordsize="21600,21600" o:spt="134" path="m17955,v862,282,1877,1410,2477,3045c21035,5357,21372,7895,21597,10827v-225,2763,-562,5300,-1165,7613c19832,20132,18817,21260,17955,21597r-14388,l,10827,3567,xe">
							<v:stroke joinstyle="miter"/>
							<v:path o:connecttype="rect" textboxrect="3567,0,17955,21600"/>
						</v:shapetype>
						<v:shape id="_x0000_s1030" type="#_x0000_t134" >
							<xsl:call-template name="ConvertShapeProperties">
								<xsl:with-param name="shapeStyle" select="$shapeStyle"/>
								<xsl:with-param name="shape" select="."/>
							</xsl:call-template>
              <!-- Sona:Defect #2020254-->
              <xsl:if test="./text:p/node()">
							<!--insert text-box-->
							<xsl:call-template name="InsertTextBox">
								<xsl:with-param name="frameStyle" select="$shapeStyle"/>
							</xsl:call-template>
              </xsl:if>
						</v:shape>
					</xsl:when>
					<!-- Rectangular Callout-->
					<xsl:when test ="draw:enhanced-geometry/@draw:enhanced-path ='M 0 0 L 0 3590 ?f2 ?f3 0 8970 0 12630 ?f4 ?f5 0 18010 0 21600 3590 21600 ?f6 ?f7 8970 21600 12630 21600 ?f8 ?f9 18010 21600 21600 21600 21600 18010 ?f10 ?f11 21600 12630 21600 8970 ?f12 ?f13 21600 3590 21600 0 18010 0 ?f14 ?f15 12630 0 8970 0 ?f16 ?f17 3590 0 0 0 Z N'">
						<v:shapetype id="_x0000_t61" coordsize="21600,21600" o:spt="61" adj="1350,25920" path="m,l0@8@12@24,0@9,,21600@6,21600@15@27@7,21600,21600,21600,21600@9@18@30,21600@8,21600,0@7,0@21@33@6,xe">
							<v:stroke joinstyle="miter"/>
							<v:formulas>
								<v:f eqn="sum 10800 0 #0"/>
								<v:f eqn="sum 10800 0 #1"/>
								<v:f eqn="sum #0 0 #1"/>
								<v:f eqn="sum @0 @1 0"/>
								<v:f eqn="sum 21600 0 #0"/>
								<v:f eqn="sum 21600 0 #1"/>
								<v:f eqn="if @0 3600 12600"/>
								<v:f eqn="if @0 9000 18000"/>
								<v:f eqn="if @1 3600 12600"/>
								<v:f eqn="if @1 9000 18000"/>
								<v:f eqn="if @2 0 #0"/>
								<v:f eqn="if @3 @10 0"/>
								<v:f eqn="if #0 0 @11"/>
								<v:f eqn="if @2 @6 #0"/>
								<v:f eqn="if @3 @6 @13"/>
								<v:f eqn="if @5 @6 @14"/>
								<v:f eqn="if @2 #0 21600"/>
								<v:f eqn="if @3 21600 @16"/>
								<v:f eqn="if @4 21600 @17"/>
								<v:f eqn="if @2 #0 @6"/>
								<v:f eqn="if @3 @19 @6"/>
								<v:f eqn="if #1 @6 @20"/>
								<v:f eqn="if @2 @8 #1"/>
								<v:f eqn="if @3 @22 @8"/>
								<v:f eqn="if #0 @8 @23"/>
								<v:f eqn="if @2 21600 #1"/>
								<v:f eqn="if @3 21600 @25"/>
								<v:f eqn="if @5 21600 @26"/>
								<v:f eqn="if @2 #1 @8"/>
								<v:f eqn="if @3 @8 @28"/>
								<v:f eqn="if @4 @8 @29"/>
								<v:f eqn="if @2 #1 0"/>
								<v:f eqn="if @3 @31 0"/>
								<v:f eqn="if #1 0 @32"/>
								<v:f eqn="val #0"/>
								<v:f eqn="val #1"/>
							</v:formulas>
							<v:path o:connecttype="custom" o:connectlocs="10800,0;0,10800;10800,21600;21600,10800;@34,@35"/>
							<v:handles>
								<v:h position="#0,#1"/>
							</v:handles>
						</v:shapetype>
						<v:shape id="_x0000_s1033" type="#_x0000_t61">
							<xsl:call-template name="ConvertShapeProperties">
								<xsl:with-param name="shapeStyle" select="$shapeStyle"/>
								<xsl:with-param name="shape" select="."/>
							</xsl:call-template>
              <!-- Sona:Defect #2020254-->
              <xsl:if test="./text:p/node()">
							<!--insert text-box-->
							<xsl:call-template name="InsertTextBox">
								<xsl:with-param name="frameStyle" select="$shapeStyle"/>
							</xsl:call-template>
              </xsl:if>
						</v:shape>
					</xsl:when>
					<!-- Round Rectangular Callout-->
					<xsl:when test ="draw:enhanced-geometry/@draw:enhanced-path ='M 3590 0 X 0 3590 L ?f2 ?f3 0 8970 0 12630 ?f4 ?f5 0 18010 Y 3590 21600 L ?f6 ?f7 8970 21600 12630 21600 ?f8 ?f9 18010 21600 X 21600 18010 L ?f10 ?f11 21600 12630 21600 8970 ?f12 ?f13 21600 3590 Y 18010 0 L ?f14 ?f15 12630 0 8970 0 ?f16 ?f17 Z N'">
						<v:shapetype id="_x0000_t62" coordsize="21600,21600" o:spt="62" adj="1350,25920" path="m3600,qx,3600l0@8@12@24,0@9,,18000qy3600,21600l@6,21600@15@27@7,21600,18000,21600qx21600,18000l21600@9@18@30,21600@8,21600,3600qy18000,l@7,0@21@33@6,xe">
							<v:stroke joinstyle="miter"/>
							<v:formulas>
								<v:f eqn="sum 10800 0 #0"/>
								<v:f eqn="sum 10800 0 #1"/>
								<v:f eqn="sum #0 0 #1"/>
								<v:f eqn="sum @0 @1 0"/>
								<v:f eqn="sum 21600 0 #0"/>
								<v:f eqn="sum 21600 0 #1"/>
								<v:f eqn="if @0 3600 12600"/>
								<v:f eqn="if @0 9000 18000"/>
								<v:f eqn="if @1 3600 12600"/>
								<v:f eqn="if @1 9000 18000"/>
								<v:f eqn="if @2 0 #0"/>
								<v:f eqn="if @3 @10 0"/>
								<v:f eqn="if #0 0 @11"/>
								<v:f eqn="if @2 @6 #0"/>
								<v:f eqn="if @3 @6 @13"/>
								<v:f eqn="if @5 @6 @14"/>
								<v:f eqn="if @2 #0 21600"/>
								<v:f eqn="if @3 21600 @16"/>
								<v:f eqn="if @4 21600 @17"/>
								<v:f eqn="if @2 #0 @6"/>
								<v:f eqn="if @3 @19 @6"/>
								<v:f eqn="if #1 @6 @20"/>
								<v:f eqn="if @2 @8 #1"/>
								<v:f eqn="if @3 @22 @8"/>
								<v:f eqn="if #0 @8 @23"/>
								<v:f eqn="if @2 21600 #1"/>
								<v:f eqn="if @3 21600 @25"/>
								<v:f eqn="if @5 21600 @26"/>
								<v:f eqn="if @2 #1 @8"/>
								<v:f eqn="if @3 @8 @28"/>
								<v:f eqn="if @4 @8 @29"/>
								<v:f eqn="if @2 #1 0"/>
								<v:f eqn="if @3 @31 0"/>
								<v:f eqn="if #1 0 @32"/>
								<v:f eqn="val #0"/>
								<v:f eqn="val #1"/>
							</v:formulas>
							<v:path o:connecttype="custom" o:connectlocs="10800,0;0,10800;10800,21600;21600,10800;@34,@35" textboxrect="791,791,20809,20809"/>
							<v:handles>
								<v:h position="#0,#1"/>
							</v:handles>
						</v:shapetype>
						<v:shape id="_x0000_s1034" type="#_x0000_t62">
							<xsl:call-template name="ConvertShapeProperties">
								<xsl:with-param name="shapeStyle" select="$shapeStyle"/>
								<xsl:with-param name="shape" select="."/>
							</xsl:call-template>
              <!-- Sona:Defect #2020254-->
              <xsl:if test="./text:p/node()">
							<!--insert text-box-->
							<xsl:call-template name="InsertTextBox">
								<xsl:with-param name="frameStyle" select="$shapeStyle"/>
							</xsl:call-template>
              </xsl:if>
						</v:shape>
					</xsl:when>
					<!-- Oval Callout-->
					<xsl:when test ="draw:enhanced-geometry/@draw:enhanced-path ='W 0 0 21600 21600 ?f22 ?f23 ?f18 ?f19 L ?f14 ?f15 Z N'">
						<v:shapetype id="_x0000_t63" coordsize="21600,21600" o:spt="63" adj="1350,25920" path="wr,,21600,21600@15@16@17@18l@21@22xe">
							<v:stroke joinstyle="miter"/>
							<v:formulas>
								<v:f eqn="val #0"/>
								<v:f eqn="val #1"/>
								<v:f eqn="sum 10800 0 #0"/>
								<v:f eqn="sum 10800 0 #1"/>
								<v:f eqn="atan2 @2 @3"/>
								<v:f eqn="sumangle @4 11 0"/>
								<v:f eqn="sumangle @4 0 11"/>
								<v:f eqn="cos 10800 @4"/>
								<v:f eqn="sin 10800 @4"/>
								<v:f eqn="cos 10800 @5"/>
								<v:f eqn="sin 10800 @5"/>
								<v:f eqn="cos 10800 @6"/>
								<v:f eqn="sin 10800 @6"/>
								<v:f eqn="sum 10800 0 @7"/>
								<v:f eqn="sum 10800 0 @8"/>
								<v:f eqn="sum 10800 0 @9"/>
								<v:f eqn="sum 10800 0 @10"/>
								<v:f eqn="sum 10800 0 @11"/>
								<v:f eqn="sum 10800 0 @12"/>
								<v:f eqn="mod @2 @3 0"/>
								<v:f eqn="sum @19 0 10800"/>
								<v:f eqn="if @20 #0 @13"/>
								<v:f eqn="if @20 #1 @14"/>
							</v:formulas>
							<v:path o:connecttype="custom" o:connectlocs="10800,0;3163,3163;0,10800;3163,18437;10800,21600;18437,18437;21600,10800;18437,3163;@21,@22" textboxrect="3163,3163,18437,18437"/>
							<v:handles>
								<v:h position="#0,#1"/>
							</v:handles>
						</v:shapetype>
						<v:shape id="_x0000_s1035" type="#_x0000_t63">
							<xsl:call-template name="ConvertShapeProperties">
								<xsl:with-param name="shapeStyle" select="$shapeStyle"/>
								<xsl:with-param name="shape" select="."/>
							</xsl:call-template>
              <!-- Sona:Defect #2020254-->
              <xsl:if test="./text:p/node()">
							<!--insert text-box-->
							<xsl:call-template name="InsertTextBox">
								<xsl:with-param name="frameStyle" select="$shapeStyle"/>
							</xsl:call-template>
              </xsl:if>
						</v:shape>
					</xsl:when>
					<!--Right Arrow-->
					<xsl:when test ="draw:enhanced-geometry/@draw:enhanced-path ='M 0 ?f0 L ?f1 ?f0 ?f1 0 21600 10800 ?f1 21600 ?f1 ?f2 0 ?f2 Z N'">
						<v:shapetype id="_x0000_t13" coordsize="21600,21600" o:spt="13" adj="16200,5400" path="m@0,l@0@1,0@1,0@2@0@2@0,21600,21600,10800xe">
							<v:stroke joinstyle="miter"/>
							<v:formulas>
								<v:f eqn="val #0"/>
								<v:f eqn="val #1"/>
								<v:f eqn="sum height 0 #1"/>
								<v:f eqn="sum 10800 0 #1"/>
								<v:f eqn="sum width 0 #0"/>
								<v:f eqn="prod @4 @3 10800"/>
								<v:f eqn="sum width 0 @5"/>
							</v:formulas>
							<v:path o:connecttype="custom" o:connectlocs="@0,0;0,10800;@0,21600;21600,10800" o:connectangles="270,180,90,0" textboxrect="0,@1,@6,@2"/>
							<v:handles>
								<v:h position="#0,#1" xrange="0,21600" yrange="0,10800"/>
							</v:handles>
						</v:shapetype>
						<v:shape id="_x0000_s1032" type="#_x0000_t13" >
							<xsl:call-template name="ConvertShapeProperties">
								<xsl:with-param name="shapeStyle" select="$shapeStyle"/>
								<xsl:with-param name="shape" select="."/>
							</xsl:call-template>
              <!-- Sona:Defect #2020254-->
              <xsl:if test="./text:p/node()">
							<!--insert text-box-->
							<xsl:call-template name="InsertTextBox">
								<xsl:with-param name="frameStyle" select="$shapeStyle"/>
							</xsl:call-template>
              </xsl:if>
						</v:shape>
					</xsl:when>
					<!-- Left Arrow-->
					<xsl:when test ="draw:enhanced-geometry/@draw:enhanced-path ='M 21600 ?f0 L ?f1 ?f0 ?f1 0 0 10800 ?f1 21600 ?f1 ?f2 21600 ?f2 Z N'">
						<v:shapetype id="_x0000_t66" coordsize="21600,21600" o:spt="66" adj="5400,5400" path="m@0,l@0@1,21600@1,21600@2@0@2@0,21600,,10800xe">
							<v:stroke joinstyle="miter"/>
							<v:formulas>
								<v:f eqn="val #0"/>
								<v:f eqn="val #1"/>
								<v:f eqn="sum 21600 0 #1"/>
								<v:f eqn="prod #0 #1 10800"/>
								<v:f eqn="sum #0 0 @3"/>
							</v:formulas>
							<v:path o:connecttype="custom" o:connectlocs="@0,0;0,10800;@0,21600;21600,10800" o:connectangles="270,180,90,0" textboxrect="@4,@1,21600,@2"/>
							<v:handles>
								<v:h position="#0,#1" xrange="0,21600" yrange="0,10800"/>
							</v:handles>
						</v:shapetype>
						<v:shape id="_x0000_s1033" type="#_x0000_t66">
							<xsl:call-template name="ConvertShapeProperties">
								<xsl:with-param name="shapeStyle" select="$shapeStyle"/>
								<xsl:with-param name="shape" select="."/>
							</xsl:call-template>
              <!-- Sona:Defect #2020254-->
              <xsl:if test="./text:p/node()">
							<!--insert text-box-->
							<xsl:call-template name="InsertTextBox">
								<xsl:with-param name="frameStyle" select="$shapeStyle"/>
							</xsl:call-template>
              </xsl:if>
						</v:shape>
					</xsl:when>
					<!-- Up Arrow -->
					<xsl:when test ="draw:enhanced-geometry/@draw:enhanced-path ='M ?f0 21600 L ?f0 ?f1 0 ?f1 10800 0 21600 ?f1 ?f2 ?f1 ?f2 21600 Z N'">
						<v:shapetype id="_x0000_t68" coordsize="21600,21600" o:spt="68" adj="5400,5400" path="m0@0l@1@0@1,21600@2,21600@2@0,21600@0,10800,xe">
							<v:stroke joinstyle="miter"/>
							<v:formulas>
								<v:f eqn="val #0"/>
								<v:f eqn="val #1"/>
								<v:f eqn="sum 21600 0 #1"/>
								<v:f eqn="prod #0 #1 10800"/>
								<v:f eqn="sum #0 0 @3"/>
							</v:formulas>
							<v:path o:connecttype="custom" o:connectlocs="10800,0;0,@0;10800,21600;21600,@0" o:connectangles="270,180,90,0" textboxrect="@1,@4,@2,21600"/>
							<v:handles>
								<v:h position="#1,#0" xrange="0,10800" yrange="0,21600"/>
							</v:handles>
						</v:shapetype>
						<v:shape id="_x0000_s1034" type="#_x0000_t68">
							<xsl:call-template name="ConvertShapeProperties">
								<xsl:with-param name="shapeStyle" select="$shapeStyle"/>
								<xsl:with-param name="shape" select="."/>
							</xsl:call-template>
              <!-- Sona:Defect #2020254-->
              <xsl:if test="./text:p/node()">
							<!--insert text-box-->
							<xsl:call-template name="InsertTextBox">
								<xsl:with-param name="frameStyle" select="$shapeStyle"/>
							</xsl:call-template>
              </xsl:if>
						</v:shape>
					</xsl:when>
					<!-- Down Arrow-->
					<xsl:when test ="draw:enhanced-geometry/@draw:enhanced-path ='M ?f0 0 L ?f0 ?f1 0 ?f1 10800 21600 21600 ?f1 ?f2 ?f1 ?f2 0 Z N'">
						<v:shapetype id="_x0000_t67" coordsize="21600,21600" o:spt="67" adj="16200,5400" path="m0@0l@1@0@1,0@2,0@2@0,21600@0,10800,21600xe">
							<v:stroke joinstyle="miter"/>
							<v:formulas>
								<v:f eqn="val #0"/>
								<v:f eqn="val #1"/>
								<v:f eqn="sum height 0 #1"/>
								<v:f eqn="sum 10800 0 #1"/>
								<v:f eqn="sum width 0 #0"/>
								<v:f eqn="prod @4 @3 10800"/>
								<v:f eqn="sum width 0 @5"/>
							</v:formulas>
							<v:path o:connecttype="custom" o:connectlocs="10800,0;0,@0;10800,21600;21600,@0" o:connectangles="270,180,90,0" textboxrect="@1,0,@2,@6"/>
							<v:handles>
								<v:h position="#1,#0" xrange="0,10800" yrange="0,21600"/>
							</v:handles>
						</v:shapetype>
						<v:shape id="_x0000_s1035" type="#_x0000_t67" >
							<xsl:call-template name="ConvertShapeProperties">
								<xsl:with-param name="shapeStyle" select="$shapeStyle"/>
								<xsl:with-param name="shape" select="."/>
							</xsl:call-template>
              <!-- Sona:Defect #2020254-->
              <xsl:if test="./text:p/node()">
							<!--insert text-box-->
							<xsl:call-template name="InsertTextBox">
								<xsl:with-param name="frameStyle" select="$shapeStyle"/>
							</xsl:call-template>
              </xsl:if>
						</v:shape>


					</xsl:when>
					<!-- Trapezoid-->
					<xsl:when test ="draw:enhanced-geometry/@draw:enhanced-path ='M 0 0 L 21600 0 ?f0 21600 ?f1 21600 Z N'">
						<v:shapetype id="_x0000_t8" coordsize="21600,21600" o:spt="8" adj="5400" path="m,l@0,21600@1,21600,21600,xe">
							<v:stroke joinstyle="miter"/>
							<v:formulas>
								<v:f eqn="val #0"/>
								<v:f eqn="sum width 0 #0"/>
								<v:f eqn="prod #0 1 2"/>
								<v:f eqn="sum width 0 @2"/>
								<v:f eqn="mid #0 width"/>
								<v:f eqn="mid @1 0"/>
								<v:f eqn="prod height width #0"/>
								<v:f eqn="prod @6 1 2"/>
								<v:f eqn="sum height 0 @7"/>
								<v:f eqn="prod width 1 2"/>
								<v:f eqn="sum #0 0 @9"/>
								<v:f eqn="if @10 @8 0"/>
								<v:f eqn="if @10 @7 height"/>
							</v:formulas>
							<v:path gradientshapeok="t" o:connecttype="custom" o:connectlocs="@3,10800;10800,21600;@2,10800;10800,0" textboxrect="1800,1800,19800,19800;4500,4500,17100,17100;7200,7200,14400,14400"/>
							<v:handles>
								<v:h position="#0,bottomRight" xrange="0,10800"/>
							</v:handles>
						</v:shapetype>
						<v:shape id="_x0000_s1037" type="#_x0000_t8" >
							<xsl:call-template name="ConvertShapeProperties">
								<xsl:with-param name="shapeStyle" select="$shapeStyle"/>
								<xsl:with-param name="shape" select="."/>
							</xsl:call-template>
              <!-- Sona:Defect #2020254-->
              <xsl:if test="./text:p/node()">
							<!--insert text-box-->
							<xsl:call-template name="InsertTextBox">
								<xsl:with-param name="frameStyle" select="$shapeStyle"/>
							</xsl:call-template>
              </xsl:if>
						</v:shape>

						<!-- Can-->
					</xsl:when>
					<xsl:when test ="draw:enhanced-geometry/@draw:enhanced-path ='M 44 0 C 20 0 0 ?f2 0 ?f0 L 0 ?f3 C 0 ?f4 20 21600 44 21600 68 21600 88 ?f4 88 ?f3 L 88 ?f0 C 88 ?f2 68 0 44 0 Z N M 44 0 C 20 0 0 ?f2 0 ?f0 0 ?f5 20 ?f6 44 ?f6 68 ?f6 88 ?f5 88 ?f0 88 ?f2 68 0 44 0 Z N'">
						<v:shapetype id="_x0000_t22" coordsize="21600,21600" o:spt="22" adj="5400" path="m10800,qx0@1l0@2qy10800,21600,21600@2l21600@1qy10800,xem0@1qy10800@0,21600@1nfe">
							<v:formulas>
								<v:f eqn="val #0"/>
								<v:f eqn="prod #0 1 2"/>
								<v:f eqn="sum height 0 @1"/>
							</v:formulas>
							<v:path o:extrusionok="f" gradientshapeok="t" o:connecttype="custom" o:connectlocs="10800,@0;10800,0;0,10800;10800,21600;21600,10800" o:connectangles="270,270,180,90,0" textboxrect="0,@0,21600,@2"/>
							<v:handles>
								<v:h position="center,#0" yrange="0,10800"/>
							</v:handles>
							<o:complex v:ext="view"/>
						</v:shapetype>
						<v:shape id="_x0000_s1038" type="#_x0000_t22" >
							<xsl:call-template name="ConvertShapeProperties">
								<xsl:with-param name="shapeStyle" select="$shapeStyle"/>
								<xsl:with-param name="shape" select="."/>
							</xsl:call-template>
              <!-- Sona:Defect #2020254-->
              <xsl:if test="./text:p/node()">
							<!--insert text-box-->
							<xsl:call-template name="InsertTextBox">
								<xsl:with-param name="frameStyle" select="$shapeStyle"/>
							</xsl:call-template>
              </xsl:if>
						</v:shape>
					</xsl:when>
					<!-- Cube-->
					<xsl:when test ="draw:enhanced-geometry/@draw:enhanced-path ='M 0 ?f12 L 0 ?f1 ?f2 0 ?f11 0 ?f11 ?f3 ?f4 ?f12 Z N M 0 ?f1 L ?f2 0 ?f11 0 ?f4 ?f1 Z N M ?f4 ?f12 L ?f4 ?f1 ?f11 0 ?f11 ?f3 Z N'">
						<v:shapetype id="_x0000_t16" coordsize="21600,21600" o:spt="16" adj="5400" path="m@0,l0@0,,21600@1,21600,21600@2,21600,xem0@0nfl@1@0,21600,em@1@0nfl@1,21600e">
							<v:stroke joinstyle="miter"/>
							<v:formulas>
								<v:f eqn="val #0"/>
								<v:f eqn="sum width 0 #0"/>
								<v:f eqn="sum height 0 #0"/>
								<v:f eqn="mid height #0"/>
								<v:f eqn="prod @1 1 2"/>
								<v:f eqn="prod @2 1 2"/>
								<v:f eqn="mid width #0"/>
							</v:formulas>
							<v:path o:extrusionok="f" gradientshapeok="t" limo="10800,10800" o:connecttype="custom" o:connectlocs="@6,0;@4,@0;0,@3;@4,21600;@1,@3;21600,@5" o:connectangles="270,270,180,90,0,0" textboxrect="0,@0,@1,21600"/>
							<v:handles>
								<v:h position="topLeft,#0" switch="" yrange="0,21600"/>
							</v:handles>
							<o:complex v:ext="view"/>
						</v:shapetype>
						<v:shape id="_x0000_s1039" type="#_x0000_t16" style="position:absolute;margin-left:2in;margin-top:276.75pt;width:81pt;height:79.5pt;z-index:251660288">
							<xsl:call-template name="ConvertShapeProperties">
								<xsl:with-param name="shapeStyle" select="$shapeStyle"/>
								<xsl:with-param name="shape" select="."/>
							</xsl:call-template>
              <!-- Sona:Defect #2020254-->
              <xsl:if test="./text:p/node()">
							<!--insert text-box-->
							<xsl:call-template name="InsertTextBox">
								<xsl:with-param name="frameStyle" select="$shapeStyle"/>
							</xsl:call-template>
              </xsl:if>
						</v:shape>
					</xsl:when>
					<!-- Octagon -->
					<xsl:when test ="draw:enhanced-geometry/@draw:enhanced-path ='M ?f0 0 L ?f2 0 21600 ?f1 21600 ?f3 ?f2 21600 ?f0 21600 0 ?f3 0 ?f1 Z N'">
						<v:shapetype id="_x0000_t10" coordsize="21600,21600" o:spt="10" adj="6326" path="m@0,l0@0,0@2@0,21600@1,21600,21600@2,21600@0@1,xe">
							<v:stroke joinstyle="miter"/>
							<v:formulas>
								<v:f eqn="val #0"/>
								<v:f eqn="sum width 0 #0"/>
								<v:f eqn="sum height 0 #0"/>
								<v:f eqn="prod @0 2929 10000"/>
								<v:f eqn="sum width 0 @3"/>
								<v:f eqn="sum height 0 @3"/>
								<v:f eqn="val width"/>
								<v:f eqn="val height"/>
								<v:f eqn="prod width 1 2"/>
								<v:f eqn="prod height 1 2"/>
							</v:formulas>
							<v:path gradientshapeok="t" limo="10800,10800" o:connecttype="custom" o:connectlocs="@8,0;0,@9;@8,@7;@6,@9" textboxrect="0,0,21600,21600;2700,2700,18900,18900;5400,5400,16200,16200"/>
							<v:handles>
								<v:h position="#0,topLeft" switch="" xrange="0,10800"/>
							</v:handles>
						</v:shapetype>
						<v:shape id="_x0000_s1040" type="#_x0000_t10" style="position:absolute;margin-left:134.25pt;margin-top:420.75pt;width:99pt;height:100.5pt;z-index:251661312">
							<xsl:call-template name="ConvertShapeProperties">
								<xsl:with-param name="shapeStyle" select="$shapeStyle"/>
								<xsl:with-param name="shape" select="."/>
							</xsl:call-template>
              <!-- Sona:Defect #2020254-->
              <xsl:if test="./text:p/node()">
							<!--insert text-box-->
							<xsl:call-template name="InsertTextBox">
								<xsl:with-param name="frameStyle" select="$shapeStyle"/>
							</xsl:call-template>
              </xsl:if>
						</v:shape>
					</xsl:when>
					<!-- Parellelogram-->
					<xsl:when test ="draw:enhanced-geometry/@draw:enhanced-path ='M ?f0 0 L 21600 0 ?f1 21600 0 21600 Z N'">
						<v:shapetype id="_x0000_t7" coordsize="21600,21600" o:spt="7" adj="5400" path="m@0,l,21600@1,21600,21600,xe">
							<v:stroke joinstyle="miter"/>
							<v:formulas>
								<v:f eqn="val #0"/>
								<v:f eqn="sum width 0 #0"/>
								<v:f eqn="prod #0 1 2"/>
								<v:f eqn="sum width 0 @2"/>
								<v:f eqn="mid #0 width"/>
								<v:f eqn="mid @1 0"/>
								<v:f eqn="prod height width #0"/>
								<v:f eqn="prod @6 1 2"/>
								<v:f eqn="sum height 0 @7"/>
								<v:f eqn="prod width 1 2"/>
								<v:f eqn="sum #0 0 @9"/>
								<v:f eqn="if @10 @8 0"/>
								<v:f eqn="if @10 @7 height"/>
							</v:formulas>
							<v:path gradientshapeok="t" o:connecttype="custom" o:connectlocs="@4,0;10800,@11;@3,10800;@5,21600;10800,@12;@2,10800" textboxrect="1800,1800,19800,19800;8100,8100,13500,13500;10800,10800,10800,10800"/>
							<v:handles>
								<v:h position="#0,topLeft" xrange="0,21600"/>
							</v:handles>
						</v:shapetype>
						<v:shape id="_x0000_s1029" type="#_x0000_t7" style="position:absolute;margin-left:80.25pt;margin-top:202.55pt;width:264.75pt;height:69.75pt;z-index:251659264">
							<xsl:call-template name="ConvertShapeProperties">
								<xsl:with-param name="shapeStyle" select="$shapeStyle"/>
								<xsl:with-param name="shape" select="."/>
							</xsl:call-template>
              <!-- Sona:Defect #2020254-->
              <xsl:if test="./text:p/node()">
							<!--insert text-box-->
							<xsl:call-template name="InsertTextBox">
								<xsl:with-param name="frameStyle" select="$shapeStyle"/>
							</xsl:call-template>
              </xsl:if>
						</v:shape>

					</xsl:when>
					<xsl:when test ="draw:enhanced-geometry/@draw:type =''">


					</xsl:when>
					<xsl:when test ="draw:enhanced-geometry/@draw:type =''">


					</xsl:when>
					<xsl:when test ="draw:enhanced-geometry/@draw:type =''">


					</xsl:when>
					<xsl:when test ="draw:enhanced-geometry/@draw:type =''">


					</xsl:when>
					<xsl:when test ="draw:enhanced-geometry/@draw:type =''">


					</xsl:when>
					<xsl:when test ="draw:enhanced-geometry/@draw:type =''">


					</xsl:when>
					<xsl:when test ="draw:enhanced-geometry/@draw:type =''">


					</xsl:when>
					<xsl:when test ="draw:enhanced-geometry/@draw:type =''">


					</xsl:when>
					<xsl:when test ="draw:enhanced-geometry/@draw:type =''">


					</xsl:when>
					<xsl:when test ="draw:enhanced-geometry/@draw:type =''">


					</xsl:when>
					<xsl:when test ="draw:enhanced-geometry/@draw:type =''">


					</xsl:when>
					<xsl:when test ="draw:enhanced-geometry/@draw:type =''">


					</xsl:when>
					<!--#############Code changes done by Pradeep#######################-->
					<!-- TODO - other shapes -->
					<xsl:otherwise/>
				</xsl:choose>

			</w:pict>
		</w:r>
	</xsl:template>


	<!-- 
  Summary:  Converts draw:rect to VML rectangle
  Modified: makz (DIaLOGIKa)
  -->
	<xsl:template match="draw:rect" mode="shapes">
		<xsl:variable name="styleName" select=" @draw:style-name"/>
		<xsl:variable name="automaticStyle" select="key('automatic-styles', $styleName)"/>
		<xsl:variable name="officeStyle" select="document('styles.xml')/office:document-styles/office:styles/style:style[@style:name = $styleName]"/>

		<w:r>
			<w:pict>
				<xsl:call-template name="InsertRect">
					<xsl:with-param name="shapeStyle" select="$automaticStyle | $officeStyle"/>
					<xsl:with-param name="shape" select="."/>
				</xsl:call-template>
			</w:pict>
		</w:r>
	</xsl:template>


	<!--
  Summary:  Converts draw:ellipse to VML oval
  Author:   makz (DIaLOGIKa)
  -->
	<xsl:template match="draw:ellipse" mode="shapes">
		<xsl:variable name="styleName" select=" @draw:style-name"/>
		<xsl:variable name="automaticStyle" select="key('automatic-styles', $styleName)"/>
		<xsl:variable name="officeStyle" select="document('styles.xml')/office:document-styles/office:styles/style:style[@style:name = $styleName]"/>

		<w:r>
			<w:pict>
				<xsl:call-template name="InsertOval">
					<xsl:with-param name="shapeStyle" select="$automaticStyle | $officeStyle"/>
					<xsl:with-param name="shape" select="."/>
				</xsl:call-template>
			</w:pict>
		</w:r>
	</xsl:template>
  <!-- Sona: Defect #2019374-->
  <xsl:template match="draw:text-box" mode="paragraph">
    <w:r>
      <w:rPr>
        <xsl:variable name="prefixedStyleName">
          <xsl:call-template name="GetPrefixedStyleName">
            <xsl:with-param name="styleName" select="parent::draw:frame/@draw:style-name"/>
          </xsl:call-template>
        </xsl:variable>
        <xsl:if test="$prefixedStyleName!=''">
          <w:rStyle w:val="{$prefixedStyleName}"/>
        </xsl:if>
      </w:rPr>
      <w:pict>

        <!-- this properties are needed to make z-index work properly -->
        <v:shapetype  id="_x0000_t202" coordsize="21600,21600" path="m,l,21600r21600,l21600,xe"
          xmlns:o="urn:schemas-microsoft-com:office:office">
          <v:stroke joinstyle="miter"/>
          <v:path gradientshapeok="t" o:connecttype="rect"/>
        </v:shapetype>
        <!-- Sona: Also fixed defect #2025700-->
        <v:shape id="_x0000_s1026" type="#_x0000_t202">
          <xsl:variable name="styleName" select="parent::draw:frame/@draw:style-name"/>
          <xsl:variable name="automaticStyle" select="key('automatic-styles', $styleName)"/>
          <xsl:variable name="officeStyle" select="document('styles.xml')/office:document-styles/office:styles/style:style[@style:name = $styleName]"/>
          <xsl:variable name="shapeStyle" select="$automaticStyle | $officeStyle"/>

          <xsl:call-template name="ConvertShapeProperties">
            <xsl:with-param name="shapeStyle" select="$shapeStyle"/>
            <xsl:with-param name="shape" select="parent::draw:frame"/>
          </xsl:call-template>

          <xsl:call-template name="InsertTextBox">
            <xsl:with-param name="frameStyle" select="$shapeStyle"/>
            <xsl:with-param name="frame" select="parent::draw:frame"/>
          </xsl:call-template>
        </v:shape>
      </w:pict>
    </w:r>
  </xsl:template>

	<!-- Sona: Line Shapes-->
	<xsl:template match="draw:line" mode="shapes">
		<xsl:variable name="styleName" select=" @draw:style-name"/>
		<xsl:variable name="automaticStyle" select="key('automatic-styles', $styleName)"/>
		<xsl:variable name="officeStyle" select="document('styles.xml')/office:document-styles/office:styles/style:style[@style:name = $styleName]"/>

		<w:r>
			<w:pict>
				<v:shapetype id="_x0000_t32" coordsize="21600,21600" o:spt="32" o:oned="t" path="m,l21600,21600e" filled="f">
					<v:path arrowok="t" fillok="f" o:connecttype="none"/>
					<o:lock v:ext="edit" shapetype="t"/>
				</v:shapetype>
				<v:shape id="_x0000_s1026" type="#_x0000_t32">
          <!--added by chhavi for alttext-->
          <xsl:if test ="child::node()= svg:desc">
            <xsl:attribute name="alt">
              <xsl:value-of select="svg:desc/child::node()"/>
            </xsl:attribute>
          </xsl:if>
          <!--end here-->
					<xsl:attribute name="style">
						<xsl:call-template name ="GetLineCoordinatesODF">
							<xsl:with-param name="shape" select="."/>
						</xsl:call-template>
            <!-- Sona : Defect #2026780-->
            <!-- z-Index-->
            <xsl:call-template name="InsertShapeZIndex">
              <xsl:with-param name="shapeStyle" select="$automaticStyle | $officeStyle"/>
              <xsl:with-param name="shape" select="."/>
            </xsl:call-template>
            
            <!--Sona Added margin for completing Shape Wrap feature-->
            <xsl:call-template name="FrameToShapeMargin">
              <xsl:with-param name="frameStyle" select="$automaticStyle | $officeStyle"/>
              <xsl:with-param name="frame" select="."/>
            </xsl:call-template>
					</xsl:attribute>
					<xsl:attribute name="o:connectortype">
						<xsl:value-of select="'straight'"/>
					</xsl:attribute>
					<xsl:call-template name="InsertShapeStroke">
						<xsl:with-param name="shapeStyle" select="$automaticStyle|$officeStyle"/>
					</xsl:call-template>
					<!-- Sona Added Dashed Lines-->
					<xsl:call-template name="GetLineStroke">
						<xsl:with-param name="shapeStyle" select="$automaticStyle|$officeStyle"/>
					</xsl:call-template>
					<!--Sona Added Shape Wrap-->
					<xsl:call-template name="FrameToShapeWrap">
						<xsl:with-param name="frameStyle" select="$automaticStyle|$officeStyle"/>
					</xsl:call-template>
          <!--Sona Added Shape shadow-->
          <xsl:call-template name="FrameToShapeShadow">
            <xsl:with-param name="frameStyle" select="$automaticStyle|$officeStyle"/>
          </xsl:call-template>
				</v:shape>
			</w:pict>
		</w:r>
	</xsl:template>
	<xsl:template match="draw:connector" mode="shapes">
		<xsl:variable name="styleName" select=" @draw:style-name"/>
		<xsl:variable name="automaticStyle" select="key('automatic-styles', $styleName)"/>
		<xsl:variable name="officeStyle" select="document('styles.xml')/office:document-styles/office:styles/style:style[@style:name = $styleName]"/>
		<w:r>
			<w:pict>
				<xsl:choose>
					<xsl:when test ="@draw:type='curve'">
						<v:shapetype id="_x0000_t38" coordsize="21600,21600" o:spt="38" o:oned="t" path="m,c@0,0@1,5400@1,10800@1,16200@2,21600,21600,21600e" filled="f">
							<v:formulas>
								<v:f eqn="mid #0 0"/>
								<v:f eqn="val #0"/>
								<v:f eqn="mid #0 21600"/>
							</v:formulas>
							<v:path arrowok="t" fillok="f" o:connecttype="none"/>
							<v:handles>
								<v:h position="#0,center"/>
							</v:handles>
							<o:lock v:ext="edit" shapetype="t"/>
						</v:shapetype>
            <!-- Sona: Defect 2019239-->
						<v:shape id="_x0000_s1036" type="#_x0000_t38" adj="10800,22173,-43579">
							<xsl:attribute name="style">
								<xsl:call-template name ="GetLineCoordinatesODF">
									<xsl:with-param name="shape" select="."/>
								</xsl:call-template>
                <!-- Sona : Defect #2026780-->
                <!-- z-Index-->
                <xsl:call-template name="InsertShapeZIndex">
                  <xsl:with-param name="shapeStyle" select="$automaticStyle | $officeStyle"/>
                  <xsl:with-param name="shape" select="."/>
                </xsl:call-template>

                <!--Sona Added margin for completing Shape Wrap feature-->
                <xsl:call-template name="FrameToShapeMargin">
                  <xsl:with-param name="frameStyle" select="$automaticStyle | $officeStyle"/>
                  <xsl:with-param name="frame" select="."/>
                </xsl:call-template>
							</xsl:attribute>
							<xsl:attribute name="o:connectortype">
								<xsl:value-of select="'curved'"/>
							</xsl:attribute>
							<xsl:call-template name="InsertShapeStroke">
								<xsl:with-param name="shapeStyle" select="$automaticStyle|$officeStyle"/>
							</xsl:call-template>
							<!-- Sona Added Dashed Lines-->
							<xsl:call-template name="GetLineStroke">
								<xsl:with-param name="shapeStyle" select="$automaticStyle|$officeStyle"/>
							</xsl:call-template>
							<!--Sona Added Shape Wrap-->
							<xsl:call-template name="FrameToShapeWrap">
								<xsl:with-param name="frameStyle" select="$automaticStyle|$officeStyle"/>
							</xsl:call-template>
              <!--Sona Added Shape shadow-->
              <xsl:call-template name="FrameToShapeShadow">
                <xsl:with-param name="frameStyle" select="$automaticStyle|$officeStyle"/>
              </xsl:call-template>
						</v:shape>
					</xsl:when>
          <!-- Sona: Defect #2019464-->
          <xsl:when test="@draw:type='line'">
            <v:shapetype id="_x0000_t32" coordsize="21600,21600" o:spt="32" o:oned="t" path="m,l21600,21600e" filled="f">
              <v:path arrowok="t" fillok="f" o:connecttype="none"/>
              <o:lock v:ext="edit" shapetype="t"/>
            </v:shapetype>
            <v:shape id="_x0000_s1026" type="#_x0000_t32">
              <xsl:attribute name="style">
                <xsl:call-template name ="GetLineCoordinatesODF">
                  <xsl:with-param name="shape" select="."/>
                </xsl:call-template>
                <!-- Sona : Defect #2026780-->
                <!-- z-Index-->
                <xsl:call-template name="InsertShapeZIndex">
                  <xsl:with-param name="shapeStyle" select="$automaticStyle | $officeStyle"/>
                  <xsl:with-param name="shape" select="."/>
                </xsl:call-template>

                <!--Sona Added margin for completing Shape Wrap feature-->
                <xsl:call-template name="FrameToShapeMargin">
                  <xsl:with-param name="frameStyle" select="$automaticStyle | $officeStyle"/>
                  <xsl:with-param name="frame" select="."/>
                </xsl:call-template>
              </xsl:attribute>
              <xsl:attribute name="o:connectortype">
                <xsl:value-of select="'straight'"/>
              </xsl:attribute>
              <xsl:call-template name="InsertShapeStroke">
                <xsl:with-param name="shapeStyle" select="$automaticStyle|$officeStyle"/>
              </xsl:call-template>
              <!-- Sona Added Dashed Lines-->
              <xsl:call-template name="GetLineStroke">
                <xsl:with-param name="shapeStyle" select="$automaticStyle|$officeStyle"/>
              </xsl:call-template>
              <!--Sona Added Shape Wrap-->
              <xsl:call-template name="FrameToShapeWrap">
                <xsl:with-param name="frameStyle" select="$automaticStyle|$officeStyle"/>
              </xsl:call-template>
              <!--Sona Added Shape shadow-->
              <xsl:call-template name="FrameToShapeShadow">
                <xsl:with-param name="frameStyle" select="$automaticStyle|$officeStyle"/>
              </xsl:call-template>
            </v:shape>
          </xsl:when>
          <!-- Sona: Defect #2019464-->
					<xsl:otherwise>
						<v:shapetype id="_x0000_t34" coordsize="21600,21600" o:spt="34" o:oned="t" adj="10800" path="m,l@0,0@0,21600,21600,21600e" filled="f">
							<v:stroke joinstyle="miter"/>
							<v:formulas>
								<v:f eqn="val #0"/>
							</v:formulas>
							<v:path arrowok="t" fillok="f" o:connecttype="none"/>
							<v:handles>
								<v:h position="#0,center"/>
							</v:handles>
							<o:lock v:ext="edit" shapetype="t"/>
						</v:shapetype>
						<v:shape id="_x0000_s1033" type="#_x0000_t34" adj=",-249300,-14316">
							<xsl:attribute name="style">
								<xsl:call-template name ="GetLineCoordinatesODF">
									<xsl:with-param name="shape" select="."/>
								</xsl:call-template>
                <!-- Sona : Defect #2026780-->
                <!-- z-Index-->
                <xsl:call-template name="InsertShapeZIndex">
                  <xsl:with-param name="shapeStyle" select="$automaticStyle | $officeStyle"/>
                  <xsl:with-param name="shape" select="."/>
                </xsl:call-template>

                <!--Sona Added margin for completing Shape Wrap feature-->
                <xsl:call-template name="FrameToShapeMargin">
                  <xsl:with-param name="frameStyle" select="$automaticStyle | $officeStyle"/>
                  <xsl:with-param name="frame" select="."/>
                </xsl:call-template>
							</xsl:attribute>
							<xsl:attribute name="o:connectortype">
								<xsl:value-of select="'elbow'"/>
							</xsl:attribute>
							<xsl:call-template name="InsertShapeStroke">
								<xsl:with-param name="shapeStyle" select="$automaticStyle|$officeStyle"/>
							</xsl:call-template>
							<!-- Sona Added Dashed Lines-->
							<xsl:call-template name="GetLineStroke">
								<xsl:with-param name="shapeStyle" select="$automaticStyle|$officeStyle"/>
							</xsl:call-template>
							<!--Sona Added Shape Wrap-->
							<xsl:call-template name="FrameToShapeWrap">
								<xsl:with-param name="frameStyle" select="$automaticStyle|$officeStyle"/>
							</xsl:call-template>
              <!--Sona Added Shape shadow-->
              <xsl:call-template name="FrameToShapeShadow">
                <xsl:with-param name="frameStyle" select="$automaticStyle|$officeStyle"/>
              </xsl:call-template>
						</v:shape>
					</xsl:otherwise>
				</xsl:choose>
			</w:pict>
		</w:r>
	</xsl:template>
	<!-- 
  *************************************************************************
  CALLED TEMPLATES
  *************************************************************************
  -->


	<!--
  Summary:  Inserts a VML rectangle
  Author:   makz (DIaLOGIKa)
  Params:   shapeStyle: The automatic style of the draw:shape
            shape: The draw:shape itself
  -->
	<xsl:template name="InsertRect">
		<xsl:param name="shapeStyle"/>
		<xsl:param name="shape" />

		<v:rect id="_x0000_s1026">

			<xsl:call-template name="ConvertShapeProperties">
				<xsl:with-param name="shapeStyle" select="$shapeStyle"/>
				<xsl:with-param name="shape" select="$shape"/>
			</xsl:call-template>
      <!-- Sona:Defect #2020254-->
      <xsl:if test="$shape/text:p/node()">
			<!-- reuse the frame template, attributes are the same -->
			<xsl:call-template name="InsertTextBox">
				<xsl:with-param name="frameStyle" select="$shapeStyle"/>
			</xsl:call-template>
      </xsl:if>
		</v:rect>
	</xsl:template>


	<!--
  Summary:  Inserts a VML rounded rectangle
  Author:   makz (DIaLOGIKa)
  Params:   shapeStyle: The automatic style of the draw:shape
            shape: The draw:shape itself
  -->
	<xsl:template name="InsertRoundedRect">
		<xsl:param name="shapeStyle"/>
		<xsl:param name="shape" />

    <v:roundrect id="_x0000_s1032">

			<xsl:call-template name="ConvertShapeProperties">
				<xsl:with-param name="shapeStyle" select="$shapeStyle"/>
				<xsl:with-param name="shape" select="$shape"/>
			</xsl:call-template>

      <!-- Sona:Defect #2020254-->
      <xsl:if test="$shape/text:p/node()">
			<!-- reuse the frame template, attributes are the same -->
			<xsl:call-template name="InsertTextBox">
				<xsl:with-param name="frameStyle" select="$shapeStyle"/>
			</xsl:call-template>
      </xsl:if>
		</v:roundrect>
	</xsl:template>


	<!--
  Summary:  Inserts a VML oval
  Author:   makz (DIaLOGIKa)
  Params:   shapeStyle: The automatic style of the draw:shape
            shape: The draw:shape itself
  -->
	<xsl:template name="InsertOval">
		<xsl:param name="shapeStyle"/>
		<xsl:param name="shape" />

    <v:oval id="_x0000_s1037">

			<xsl:call-template name="ConvertShapeProperties">
				<xsl:with-param name="shapeStyle" select="$shapeStyle"/>
				<xsl:with-param name="shape" select="$shape"/>
			</xsl:call-template>
      <!-- Sona:Defect #2020254-->
      <xsl:if test="$shape/text:p/node()">
			<!-- reuse the frame template, attributes are the same -->
			<xsl:call-template name="InsertTextBox">
				<xsl:with-param name="frameStyle" select="$shapeStyle"/>
			</xsl:call-template>
      </xsl:if>
		</v:oval>
	</xsl:template>


	<!-- 
  Summary:  Converts the properties of a draw:rect/draw:oval (e.g.) to VML properties
  Author:   CleverAge
  Params:   shapeStyle: The automatic style of the draw:shape
            shape: The draw:shape itself
  -->
	<xsl:template name="ConvertShapeProperties">
		<xsl:param name="shapeStyle"/>
		<xsl:param name="shape"/>
    <!--added by chhavi for alttext-->
    <xsl:if test ="child::node()= svg:desc">
      <xsl:attribute name="alt">
        <xsl:value-of select="svg:desc/child::node()"/>
      </xsl:attribute>
    </xsl:if>
    <!--end here-->
		<xsl:if test="$shapeStyle != 0 or count($shapeStyle) &gt; 1">

			<xsl:call-template name="InsertShapeStyleAttribute">
				<xsl:with-param name="shapeStyle" select="$shapeStyle"/>
				<xsl:with-param name="shape" select="$shape"/>
			</xsl:call-template>

			<xsl:call-template name="InsertShapeStroke">
				<xsl:with-param name="shapeStyle" select="$shapeStyle"/>
			</xsl:call-template>

			<xsl:call-template name="InsertShapeFill">
				<xsl:with-param name="shapeStyle" select="$shapeStyle"/>
		  <xsl:with-param name="shape" select="$shape"/>
			</xsl:call-template>

			<!-- Sona Added Dashed Lines-->
			<xsl:call-template name="GetLineStroke">
				<xsl:with-param name="shapeStyle" select="$shapeStyle"/>
			</xsl:call-template>

			<!--Sona Added Shape Wrap-->
			<xsl:call-template name="FrameToShapeWrap">
				<xsl:with-param name="frameStyle" select="$shapeStyle"/>
			</xsl:call-template>
      <!--Sona Added Shape shadow-->
      <xsl:call-template name="FrameToShapeShadow">
        <xsl:with-param name="frameStyle" select="$shapeStyle"/>
      </xsl:call-template>
		</xsl:if>
	</xsl:template>


	<!-- 
  Summary:  Inserts the style attribute of a VML shape
  Author:   CleverAge
  Params:   shapeStyle: The automatic style of the draw:shape
            shape: The draw:shape itself
  -->
	<xsl:template name="InsertShapeStyleAttribute">
		<xsl:param name="shapeStyle"/>
		<xsl:param name="shape"/>

		<xsl:attribute name="style">

			<!-- width: -->
			<xsl:variable name="frameW">
				
          <!-- Sona: Defect #2019374-->
        <xsl:choose>
          <xsl:when test="$shape/@svg:width">
            <!-- Sona: Defect #2166160-->
            <xsl:call-template name="point-measure">
              <xsl:with-param name="length" select="$shape/@svg:width"/>
				</xsl:call-template>
          </xsl:when>
          <xsl:otherwise>
            <!-- Sona: Defect #2166160-->
            <xsl:call-template name="point-measure">
              <xsl:with-param name="length" select="$shape/@fo:min-width|$shapeStyle/style:graphic-properties/@fo:min-width"/>
            </xsl:call-template>
          </xsl:otherwise>
        </xsl:choose>
			</xsl:variable>
			<xsl:value-of select="concat('width:',$frameW,'pt;')"/>

			<!-- height: -->
			<xsl:variable name="frameH">
        <xsl:choose>
          <xsl:when test ="$shape/@svg:height">
            <!-- Sona: Defect #2166160-->
				<xsl:call-template name="point-measure">
          <!-- Sona: Defect #2019374-->
              <xsl:with-param name="length" select="$shape/@svg:height"/>
				</xsl:call-template>
          </xsl:when>
          <xsl:otherwise>
            <xsl:call-template name="point-measure">
            <!-- Sona: Defect #2166160-->
            <xsl:with-param name="length" select="$shape/child::node()/@fo:min-height|$shapeStyle/style:graphic-properties/@fo:min-height"/>
            </xsl:call-template>
          </xsl:otherwise>
        </xsl:choose>
			</xsl:variable>
			<xsl:value-of select="concat('height:',$frameH,'pt;')"/>

      <!--code added by Chhavi for relative size in Frame - Defect #2166036-->
      <xsl:if test="(parent::node()[name()='draw:frame'] or self::node()[name()='draw:frame']) and $shape/@style:rel-width">
        <xsl:value-of select="concat('mso-width-percent:',substring-before(./parent::node()/@style:rel-width,'%') * 10,';')"/>        
      </xsl:if>
      <xsl:if test="(parent::node()[name()='draw:frame'] or self::node()[name()='draw:frame']) and $shape/@style:rel-height">
        <xsl:value-of select="concat('mso-height-percent:',substring-before(./parent::node()/@style:rel-height,'%') * 10,';')"/>
      </xsl:if>
      
      <!-- z-Index-->
      <xsl:call-template name="InsertShapeZIndex">
        <xsl:with-param name="shapeStyle" select="$shapeStyle"/>
        <xsl:with-param name="shape" select="$shape"/>
      </xsl:call-template>

			<!-- reuse the frame template, attributes are the same -->
			<xsl:call-template name="FrameToRelativeShapePosition">
				<xsl:with-param name="frameStyle" select="$shapeStyle"/>
				<xsl:with-param name="frame" select="$shape"/>
			</xsl:call-template>

			<!-- reuse the frame template, attributes are the same -->
			<xsl:call-template name="FrameToShapePosition">
				<xsl:with-param name="frameStyle" select="$shapeStyle"/>
				<xsl:with-param name="frame" select="$shape"/>
			</xsl:call-template>

			<!-- reuse the frame template, attributes are the same -->
			<xsl:call-template name="FrameToTextAnchor">
				<xsl:with-param name="frameStyle" select="$shapeStyle"/>
			</xsl:call-template>
      
      <!--Sona Added margin for completing Shape Wrap feature-->
      <xsl:call-template name="FrameToShapeMargin">
        <xsl:with-param name="frameStyle" select="$shapeStyle"/>
        <xsl:with-param name="frame" select="$shape"/>
      </xsl:call-template>
		</xsl:attribute>
	</xsl:template>


	<!-- 
  Summary:  Inserts the VML shape z-index
  Author:   Sona 
  -->
  <xsl:template name ="InsertShapeZIndex">
    <xsl:param name ="shapeStyle"></xsl:param>
    <xsl:param name ="shape"></xsl:param>
    <!-- z-index: -->
    <!-- Sona: Defect #2026780-->
    <xsl:variable name="frameWrap">
      <xsl:call-template name="GetGraphicProperties">
        <xsl:with-param name="shapeStyle" select="$shapeStyle"/>
        <xsl:with-param name="attribName">style:wrap</xsl:with-param>
      </xsl:call-template>
    </xsl:variable>

    <xsl:variable name="runThrought">
      <xsl:call-template name="GetGraphicProperties">
        <xsl:with-param name="shapeStyle" select="$shapeStyle"/>
        <xsl:with-param name="attribName">style:run-through</xsl:with-param>
      </xsl:call-template>
    </xsl:variable>
    <xsl:choose>
      <xsl:when test="$frameWrap='run-through' and $runThrought='background'">
        <xsl:value-of select="concat('z-index:',-251658240,';')"/>
      </xsl:when>
      <xsl:when test="$frameWrap='run-through' and not($runThrought)">
        <xsl:value-of select="concat('z-index:',251658240,';')"/>
      </xsl:when>
      <xsl:otherwise>
        <xsl:choose>
          <!-- Sona: Defect #2019374-->
          <xsl:when test="$shape/@draw:z-index=0">
            <xsl:value-of select="concat('z-index:',2516572155-$shape/@draw:z-index,';')"/>
          </xsl:when>
          <xsl:otherwise>
            <xsl:value-of select="concat('z-index:',251659264+$shape/@draw:z-index,';')"/>
          </xsl:otherwise>
        </xsl:choose>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>

  <!-- 
  Summary:  Inserts the VML shape fill
  Author:   CleverAge
  Params:   shapeStyle: The automatic style of the draw:shape
  -->
	<xsl:template name="InsertShapeFill">
		<xsl:param name="shapeStyle"/>
	  <xsl:param name="shape"/>

		<xsl:variable name="fillColor">
      <xsl:choose>
        <xsl:when test="not($shapeStyle/style:graphic-properties/@draw:fill-color)">
          <xsl:call-template name="GetDrawnGraphicProperties">
            <xsl:with-param name="attrib">fo:background-color</xsl:with-param>
            <xsl:with-param name="shapeStyle" select="$shapeStyle"/>
          </xsl:call-template>
        </xsl:when>
        <xsl:otherwise>
			<xsl:call-template name="GetDrawnGraphicProperties">
				<xsl:with-param name="attrib">draw:fill-color</xsl:with-param>
				<xsl:with-param name="shapeStyle" select="$shapeStyle"/>
			</xsl:call-template>
        </xsl:otherwise>
      </xsl:choose>
		</xsl:variable>
		<xsl:variable name="fillProperty">
			<xsl:call-template name="GetGraphicProperties">
				<xsl:with-param name="shapeStyle" select="$shapeStyle"/>
				<xsl:with-param name="attribName">draw:fill</xsl:with-param>
			</xsl:call-template>
		</xsl:variable>
		<xsl:variable name="bgTransparency">
			<xsl:call-template name="GetGraphicProperties">
				<xsl:with-param name="shapeStyle" select="$shapeStyle"/>
				<xsl:with-param name="attribName">style:background-transparency</xsl:with-param>
			</xsl:call-template>
		</xsl:variable>
		<xsl:variable name="opacity">
			<xsl:variable name="draw-opacity">
				<xsl:call-template name="GetGraphicProperties">
					<xsl:with-param name="shapeStyle" select="$shapeStyle"/>
					<xsl:with-param name="attribName">draw:opacity</xsl:with-param>
				</xsl:call-template>
			</xsl:variable>
			<xsl:choose>
				<xsl:when test="$draw-opacity != '' ">
					<xsl:value-of select="$draw-opacity"/>
				</xsl:when>
				<xsl:otherwise>
					<xsl:if test="$bgTransparency != '' ">
						<xsl:value-of
						  select="concat((100 - number(substring-before($bgTransparency,'%'))), '%')"/>
					</xsl:if>
				</xsl:otherwise>
			</xsl:choose>
		</xsl:variable>

		<xsl:if test="$fillColor != '' ">
			<xsl:attribute name="fillcolor">
				<xsl:value-of select="$fillColor"/>
			</xsl:attribute>
		</xsl:if>
		<xsl:if test="$fillProperty != 'none' ">
			<v:fill>
				<xsl:if test="$opacity != '' ">
					<xsl:attribute name="opacity">
    <!-- Sonata: shape transperancy -->
	<xsl:value-of select="number(substring-before($opacity,'%')) div 100"/>
					</xsl:attribute>
				</xsl:if>
				<!-- other fill properties -->
				<xsl:choose>
					<xsl:when test="$fillProperty = 'solid' ">
						<xsl:attribute name="color">
							<xsl:call-template name="GetGraphicProperties">
								<xsl:with-param name="shapeStyle" select="$shapeStyle"/>
								<xsl:with-param name="attribName">draw:fill-color</xsl:with-param>
							</xsl:call-template>
						</xsl:attribute>
					</xsl:when>
  <!-- Sonata: shape Picture Fill -->
          <xsl:when test="$fillProperty = 'bitmap' ">
            <xsl:variable name="BitmapName">
              <xsl:call-template name="GetGraphicProperties">
                <xsl:with-param name="shapeStyle" select="$shapeStyle"/>
                <xsl:with-param name="attribName">draw:fill-image-name</xsl:with-param>
              </xsl:call-template>
            </xsl:variable>
            <xsl:variable name="stretch">
              <xsl:value-of select="$shapeStyle/style:graphic-properties/@style:repeat"/>
            </xsl:variable>

            <xsl:for-each select="document('styles.xml')//draw:fill-image[@draw:name = $BitmapName]">
              <!-- radial gradients not handled yet -->
              <xsl:choose>
                <xsl:when test="$stretch='stretch'">
                  <xsl:attribute name="type">frame</xsl:attribute>
                </xsl:when>
                <xsl:otherwise>
                   <xsl:attribute name="type">tile</xsl:attribute>
                </xsl:otherwise>
              </xsl:choose>
              <xsl:attribute name="recolor">t</xsl:attribute>
              <xsl:attribute name="color2">black</xsl:attribute>
              <xsl:attribute name="r:id">
                <xsl:value-of select="concat('Bitmap_',generate-id($shape))"/>
              </xsl:attribute>
            </xsl:for-each>
          </xsl:when>
					<xsl:when test="$fillProperty = 'gradient' ">
						<!-- simple linear gradient -->
						<xsl:variable name="gradientName">
							<xsl:call-template name="GetGraphicProperties">
								<xsl:with-param name="shapeStyle" select="$shapeStyle"/>
								<xsl:with-param name="attribName">draw:fill-gradient-name</xsl:with-param>
							</xsl:call-template>
						</xsl:variable>
						<xsl:for-each select="document('styles.xml')/office:document-styles/office:styles/draw:gradient[@draw:name = $gradientName]">
							<!-- radial gradients not handled yet -->
							<xsl:attribute name="type">gradient</xsl:attribute>
							<xsl:if test="@draw:angle">
								<xsl:attribute name="angle">
									<xsl:value-of select="round(number(@draw:angle) div 10)"/>
								</xsl:attribute>
							</xsl:if>
							<xsl:if test="@draw:end-color">
								<xsl:attribute name="color">
									<xsl:value-of select="@draw:end-color"/>
								</xsl:attribute>
							</xsl:if>
							<xsl:if test="@draw:start-color">
								<xsl:attribute name="color2">
									<xsl:value-of select="@draw:start-color"/>
								</xsl:attribute>
							</xsl:if>
						</xsl:for-each>
					</xsl:when>
          <!-- Sona: Picture fill for frame-->
          <xsl:when test ="$fillProperty = '' and (parent::node()[name()='draw:frame'] or self::node()[name()='draw:frame'])">
            <xsl:if test ="$shapeStyle/style:graphic-properties/style:background-image">
            <xsl:variable name="stretch">
              <xsl:value-of select="$shapeStyle/style:graphic-properties/style:background-image/@style:repeat"/>
            </xsl:variable>
            <xsl:choose>
              <xsl:when test="$stretch='stretch'">
                <xsl:attribute name="type">frame</xsl:attribute>
              </xsl:when>
              <xsl:otherwise>
                <xsl:attribute name="type">tile</xsl:attribute>
              </xsl:otherwise>
            </xsl:choose>          
        <!--<xsl:attribute name="recolor">t</xsl:attribute>-->
        <xsl:attribute name="color2">black</xsl:attribute>
        <xsl:attribute name="r:id">
          <xsl:value-of select="concat('Bitmap_',generate-id($shape))"/>
        </xsl:attribute>
            </xsl:if>
          </xsl:when>
				</xsl:choose>
			</v:fill>
		</xsl:if>
    <!--chhavi filled color-->
    <xsl:if test="$fillProperty = 'none' ">
      <xsl:attribute name="filled">
        <xsl:value-of select="'f'"/>
      </xsl:attribute>
    </xsl:if>
    <!--end here-->
	</xsl:template>


	<!-- 
  Summary:  Inserts the VML stroke of a shape
  Author:   CleverAge
  Params:   shapeStyle: The automatic style of the draw:shape
  -->
	<xsl:template name="InsertShapeStroke">
		<xsl:param name="shapeStyle"/>

		<xsl:variable name="strokeColor">
			<xsl:call-template name="GetDrawnGraphicProperties">
				<xsl:with-param name="attrib">
					<xsl:choose>
						<xsl:when test="name(parent::node()) = 'draw:frame' and not($shapeStyle/style:graphic-properties/@svg:stroke-color)">
							<xsl:value-of select="'fo:border'"/>
						</xsl:when>
						<xsl:otherwise>
							<xsl:value-of select="'svg:stroke-color'"/>
						</xsl:otherwise>
					</xsl:choose>
				</xsl:with-param>
				<xsl:with-param name="shapeStyle" select="$shapeStyle"/>
			</xsl:call-template>
		</xsl:variable>
		<xsl:variable name="strokeWeight">
			<xsl:call-template name="GetDrawnGraphicProperties">
				<!--code changed by yeswanth.s : to get width of rect & frame-->
				<!--<xsl:with-param name="attrib">svg:stroke-width</xsl:with-param>-->				
				<xsl:with-param name="attrib">
					<xsl:choose>
						<xsl:when test="name(parent::node()) = 'draw:frame' and not($shapeStyle/style:graphic-properties/@svg:stroke-width)">
							<xsl:value-of select="'fo:border'"/>
						</xsl:when>
						<xsl:otherwise>
							<xsl:value-of select="'svg:stroke-width'"/>
						</xsl:otherwise>
					</xsl:choose>
				</xsl:with-param>				
				<xsl:with-param name="shapeStyle" select="$shapeStyle"/>
			</xsl:call-template>
		</xsl:variable>

		<!-- Sona Bug fix 1835088 -->
		<xsl:variable name="shapeBorder">
			<xsl:call-template name="GetDrawnGraphicProperties">
                                <!--changes made by yeswanth.s-->
				<xsl:with-param name="attrib">
					<xsl:choose>
						<xsl:when test="name(parent::node()) = 'draw:frame' and not($shapeStyle/style:graphic-properties/@draw:stroke)">
							<xsl:value-of select="'fo:border'"/>
						</xsl:when>
						<xsl:otherwise>
							<xsl:value-of select="'draw:stroke'"/>
						</xsl:otherwise>
					</xsl:choose>
				</xsl:with-param>
                                <!--end-->
				<xsl:with-param name="shapeStyle" select="$shapeStyle"/>
			</xsl:call-template>
		</xsl:variable>

		<!--code changed by yeswanth.s : For Stroke Weight-->
		<xsl:choose>
			<xsl:when test ="$shapeBorder='none' or $shapeBorder = ''">
				<xsl:attribute name ="stroked">
					<xsl:value-of select ="'f'"/>
				</xsl:attribute>
			</xsl:when>		
			<xsl:otherwise>
		<xsl:if test="$strokeColor != '' and substring-after($strokeColor,'#')!=''">
			<xsl:attribute name="strokecolor">
				<xsl:value-of select="concat('#',substring-after($strokeColor,'#'))"/>
			</xsl:attribute>
		</xsl:if>
		<xsl:if test="$strokeWeight != '' ">
			<xsl:attribute name="strokeweight">
				<xsl:call-template name="point-measure">
					<xsl:with-param name="length" select="$strokeWeight"/>
							<!--yeswanth.s added parameter : 14-Oct-08-->
							<xsl:with-param name="round" select="'false'"/>
				</xsl:call-template>
				<xsl:text>pt</xsl:text>
			</xsl:attribute>
		</xsl:if>
			</xsl:otherwise>
		</xsl:choose>
		

	</xsl:template>


	<!-- 
  Summary:  Gets graphic properties for shapes 
  Author:   CleverAge
  -->
	<xsl:template name="GetDrawnGraphicProperties">
		<xsl:param name="attrib"/>
		<xsl:param name="shapeStyle"/>
		<xsl:choose>
			<xsl:when test="attribute::node()[name()=$attrib]">
				<xsl:value-of select="attribute::node()[name()=$attrib]"/>
			</xsl:when>
			<xsl:otherwise>
				<xsl:call-template name="GetGraphicProperties">
					<xsl:with-param name="attribName" select="$attrib"/>
					<xsl:with-param name="shapeStyle" select="$shapeStyle"/>
				</xsl:call-template>
			</xsl:otherwise>
		</xsl:choose>
	</xsl:template>

	<!-- Sona Line Functions-->
	<xsl:template name="GetLineCoordinatesODF">
		<xsl:param name="shape" select="."/>
		<xsl:variable name="x1">
			<xsl:choose>
				<xsl:when test ="@svg:x1=0">
					<xsl:value-of select="'0pt'"/>
				</xsl:when>
				<xsl:otherwise>
					<xsl:call-template name ="ConvertMeasure">
						<xsl:with-param name ="length" select ="@svg:x1"/>
						<xsl:with-param name ="unit" select ="'point'"/>
					</xsl:call-template>
				</xsl:otherwise>
			</xsl:choose>
		</xsl:variable>
		<xsl:variable name="y1">
			<xsl:choose>
				<xsl:when test ="@svg:y1=0">
					<xsl:value-of select="'0pt'"/>
				</xsl:when>
				<xsl:otherwise>
					<xsl:call-template name ="ConvertMeasure">
						<xsl:with-param name ="length" select ="@svg:y1"/>
						<xsl:with-param name ="unit" select ="'point'"/>
					</xsl:call-template>
				</xsl:otherwise>
			</xsl:choose>
		</xsl:variable>
		<xsl:variable name="x2">
			<xsl:choose>
				<xsl:when test ="@svg:x2=0">
					<xsl:value-of select="'0pt'"/>
				</xsl:when>
				<xsl:otherwise>
					<xsl:call-template name ="ConvertMeasure">
						<xsl:with-param name ="length" select ="@svg:x2"/>
						<xsl:with-param name ="unit" select ="'point'"/>
					</xsl:call-template>
				</xsl:otherwise>
			</xsl:choose>
		</xsl:variable>
		<xsl:variable name="y2">
			<xsl:choose>
				<xsl:when test ="@svg:y2=0">
					<xsl:value-of select="'0pt'"/>
				</xsl:when>
				<xsl:otherwise>
					<xsl:call-template name ="ConvertMeasure">
						<xsl:with-param name ="length" select ="@svg:y2"/>
						<xsl:with-param name ="unit" select ="'point'"/>
					</xsl:call-template>
				</xsl:otherwise>
			</xsl:choose>
		</xsl:variable>
		<xsl:text>position:absolute;</xsl:text>
		<xsl:choose>
			<xsl:when test="$x2 &gt;= $x1 and $y1 &gt; $y2">
				<xsl:text>margin-left:</xsl:text>
				<xsl:choose>
					<xsl:when test="number($x1) and $x1 != 0">
						<xsl:value-of select="$x1"/>
						<xsl:text>pt</xsl:text>
					</xsl:when>
					<xsl:otherwise>0</xsl:otherwise>
				</xsl:choose>
				<xsl:text>;</xsl:text>
				<!--<xsl:variable name="margin-top">
          <xsl:call-template name ="ConvertMeasure">
            <xsl:with-param name ="length" select ="$y2"/>
            <xsl:with-param name ="unit" select ="'point'"/>
          </xsl:call-template>
        </xsl:variable>-->
				<xsl:text>margin-top:</xsl:text>
				<xsl:choose>
					<xsl:when test="number($y2) and $y2 != 0">
						<xsl:value-of select="$y2"/>
						<xsl:text>pt</xsl:text>
					</xsl:when>
					<xsl:otherwise>0</xsl:otherwise>
				</xsl:choose>
				<xsl:text>;</xsl:text>
				<xsl:text>width:</xsl:text>
				<xsl:choose>
					<xsl:when test="number($x2 - $x1) and ($x2 - $x1) != 0">
						<xsl:value-of select="($x2 - $x1)"/>
						<xsl:text>pt</xsl:text>
					</xsl:when>
					<xsl:otherwise>0</xsl:otherwise>
				</xsl:choose>
				<xsl:text>;</xsl:text>
				<!--<xsl:variable name="height">
          <xsl:call-template name ="ConvertMeasure">
            <xsl:with-param name ="length" select ="concat(substring-before($y1,'in')-substring-before($y2,'in'),'in')"/>
            <xsl:with-param name ="unit" select ="'point'"/>
          </xsl:call-template>
        </xsl:variable>-->
				<xsl:text>height:</xsl:text>
				<xsl:choose>
					<xsl:when test="number($y1 - $y2) and $y1 - $y2 != 0">
						<xsl:value-of select="$y1 - $y2"/>
						<xsl:text>pt</xsl:text>
					</xsl:when>
					<xsl:otherwise>0</xsl:otherwise>
				</xsl:choose>
				<xsl:text>;</xsl:text>
				<xsl:text>flip:y;</xsl:text>
			</xsl:when>
			<xsl:when test="$x1 &gt; $x2 and $y2 &gt;= $y1">
				<!--<xsl:variable name="margin-left">
          <xsl:call-template name ="ConvertMeasure">
            <xsl:with-param name ="length" select ="$x2"/>
            <xsl:with-param name ="unit" select ="'point'"/>
          </xsl:call-template>
        </xsl:variable>-->
				<xsl:text>margin-left:</xsl:text>
				<xsl:choose>
					<xsl:when test="number($x2) and $x2 != 0">
						<xsl:value-of select="$x2"/>
						<xsl:text>pt</xsl:text>
					</xsl:when>
					<xsl:otherwise>0</xsl:otherwise>
				</xsl:choose>
				<xsl:text>;</xsl:text>
				<!--<xsl:variable name="margin-top">
          <xsl:call-template name ="ConvertMeasure">
            <xsl:with-param name ="length" select ="$y1"/>
            <xsl:with-param name ="unit" select ="'point'"/>
          </xsl:call-template>
        </xsl:variable>-->
				<xsl:text>margin-top:</xsl:text>
				<xsl:choose>
					<xsl:when test="number($y1) and $y1 != 0">
						<xsl:value-of select="$y1"/>
						<xsl:text>pt</xsl:text>
					</xsl:when>
					<xsl:otherwise>0</xsl:otherwise>
				</xsl:choose>
				<xsl:text>;</xsl:text>
				<!--<xsl:variable name="width">
          <xsl:call-template name ="ConvertMeasure">
            <xsl:with-param name ="length" select ="concat(substring-before($x1,'in')-substring-before($x2,'in'),'in')"/>
            <xsl:with-param name ="unit" select ="'point'"/>
          </xsl:call-template>
        </xsl:variable>-->
				<xsl:text>width:</xsl:text>
				<xsl:choose>
					<xsl:when test="number($x1 - $x2) and ($x1 - $x2) != 0">
						<xsl:value-of select="($x1 - $x2)"/>
						<xsl:text>pt</xsl:text>
					</xsl:when>
					<xsl:otherwise>0</xsl:otherwise>
				</xsl:choose>
				<xsl:text>;</xsl:text>
				<!--<xsl:variable name="height">
          <xsl:call-template name ="ConvertMeasure">
            <xsl:with-param name ="length" select ="concat(substring-before($y2,'in')-substring-before($y1,'in'),'in')"/>
            <xsl:with-param name ="unit" select ="'point'"/>
          </xsl:call-template>
        </xsl:variable>-->
				<xsl:text>height:</xsl:text>
				<xsl:choose>
					<xsl:when test="number($y2 - $y1) and ($y2 - $y1) != 0">
						<xsl:value-of select="($y2 - $y1)"/>
						<xsl:text>pt</xsl:text>
					</xsl:when>
					<xsl:otherwise>0</xsl:otherwise>
				</xsl:choose>
				<xsl:text>;</xsl:text>
				<xsl:text>flip:x;</xsl:text>

			</xsl:when>
			<xsl:when test="$x1 &gt; $x2 and $y1 &gt; $y2">
				<!--<xsl:variable name="margin-left">
          <xsl:call-template name ="ConvertMeasure">
            <xsl:with-param name ="length" select ="$x2"/>
            <xsl:with-param name ="unit" select ="'point'"/>
          </xsl:call-template>
        </xsl:variable>-->
				<xsl:text>margin-left:</xsl:text>
				<xsl:choose>
					<xsl:when test="number($x2) and $x2 != 0">
						<xsl:value-of select="$x2"/>
						<xsl:text>pt</xsl:text>
					</xsl:when>
					<xsl:otherwise>0</xsl:otherwise>
				</xsl:choose>
				<xsl:text>;</xsl:text>
				<!--<xsl:variable name="margin-top">
          <xsl:call-template name ="ConvertMeasure">
            <xsl:with-param name ="length" select ="$y2"/>
            <xsl:with-param name ="unit" select ="'point'"/>
          </xsl:call-template>
        </xsl:variable>-->
				<xsl:text>margin-top:</xsl:text>
				<xsl:choose>
					<xsl:when test="number($y2) and $y2 != 0">
						<xsl:value-of select="$y2"/>
						<xsl:text>pt</xsl:text>
					</xsl:when>
					<xsl:otherwise>0</xsl:otherwise>
				</xsl:choose>
				<xsl:text>;</xsl:text>
				<!--<xsl:variable name="width">
          <xsl:call-template name ="ConvertMeasure">
            <xsl:with-param name ="length" select ="concat(substring-before($x1,'in')-substring-before($x2,'in'),'in')"/>
            <xsl:with-param name ="unit" select ="'point'"/>
          </xsl:call-template>
        </xsl:variable>-->
				<xsl:text>width:</xsl:text>
				<xsl:choose>
					<xsl:when test="number($x1 - $x2) and ($x1 - $x2) != 0">
						<xsl:value-of select="($x1 - $x2)"/>
						<xsl:text>pt</xsl:text>
					</xsl:when>
					<xsl:otherwise>0</xsl:otherwise>
				</xsl:choose>
				<xsl:text>;</xsl:text>
				<!--<xsl:variable name="height">
          <xsl:call-template name ="ConvertMeasure">
            <xsl:with-param name ="length" select ="concat(substring-before($y1,'in')-substring-before($y2,'in'),'in')"/>
            <xsl:with-param name ="unit" select ="'point'"/>
          </xsl:call-template>
        </xsl:variable>-->
				<xsl:text>height:</xsl:text>
				<xsl:choose>
					<xsl:when test="number($y1 - $y2) and ($y1 - $y2) != 0">
						<xsl:value-of select="($y1 - $y2)"/>
						<xsl:text>pt</xsl:text>
					</xsl:when>
					<xsl:otherwise>0</xsl:otherwise>
				</xsl:choose>
				<xsl:text>;</xsl:text>
				<xsl:text>flip:x y;</xsl:text>

			</xsl:when>
			<xsl:otherwise>
				<!--<xsl:variable name="margin-left">
          <xsl:call-template name ="ConvertMeasure">
            <xsl:with-param name ="length" select ="$x1"/>
            <xsl:with-param name ="unit" select ="'point'"/>
          </xsl:call-template>
        </xsl:variable>-->
				<xsl:text>margin-left:</xsl:text>
				<xsl:choose>
					<xsl:when test="number($x1) and $x1 != 0">
						<xsl:value-of select="$x1"/>
						<xsl:text>pt</xsl:text>
					</xsl:when>
					<xsl:otherwise>0</xsl:otherwise>
				</xsl:choose>
				<xsl:text>;</xsl:text>
				<!--<xsl:variable name="margin-top">
          <xsl:call-template name ="ConvertMeasure">
            <xsl:with-param name ="length" select ="$y1"/>
            <xsl:with-param name ="unit" select ="'point'"/>
          </xsl:call-template>
        </xsl:variable>-->
				<xsl:text>margin-top:</xsl:text>
				<xsl:choose>
					<xsl:when test="number($y1) and $y1 != 0">
						<xsl:value-of select="$y1"/>
						<xsl:text>pt</xsl:text>
					</xsl:when>
					<xsl:otherwise>0</xsl:otherwise>
				</xsl:choose>
				<xsl:text>;</xsl:text>
				<!--<xsl:variable name="width">
          <xsl:call-template name ="ConvertMeasure">
            <xsl:with-param name ="length" select ="concat(substring-before($x2,'in')-substring-before($x1,'in'),'in')"/>
            <xsl:with-param name ="unit" select ="'point'"/>
          </xsl:call-template>
        </xsl:variable>-->
				<xsl:text>width:</xsl:text>
				<xsl:choose>
					<xsl:when test="number($x2 - $x1) and ($x2 - $x1) != 0">
						<xsl:value-of select="($x2 - $x1)"/>
						<xsl:text>pt</xsl:text>
					</xsl:when>
					<xsl:otherwise>0</xsl:otherwise>
				</xsl:choose>
				<xsl:text>;</xsl:text>
				<!--<xsl:variable name="height">
          <xsl:call-template name ="ConvertMeasure">
            <xsl:with-param name ="length" select ="concat(substring-before($y2,'in')-substring-before($y1,'in'),'in')"/>
            <xsl:with-param name ="unit" select ="'point'"/>
          </xsl:call-template>
        </xsl:variable>-->
				<xsl:text>height:</xsl:text>
				<xsl:choose>
					<xsl:when test="number($y2 - $y1) and ($y2 - $y1) != 0">
						<xsl:value-of select="($y2 - $y1)"/>
						<xsl:text>pt</xsl:text>
					</xsl:when>
					<xsl:otherwise>0</xsl:otherwise>
				</xsl:choose>
				<xsl:text>;</xsl:text>
			</xsl:otherwise>

		</xsl:choose>
	</xsl:template>
	<!-- Sona Arrow Stroke Function-->
	<xsl:template name="GetLineStroke">
		<xsl:param name ="shapeStyle"></xsl:param>
		<v:stroke>
			<!--code added by yeswanth.s : 'linestyle' in DOCX-->
			<xsl:variable name="styleBorderLine">
				<xsl:call-template name="GetGraphicProperties">
					<xsl:with-param name="shapeStyle" select="$shapeStyle"/>
					<xsl:with-param name="attribName">style:border-line-width</xsl:with-param>
				</xsl:call-template>
			</xsl:variable>
			<xsl:if test="$styleBorderLine != ''">
				<xsl:attribute name="linestyle">
					<xsl:choose>
						<xsl:when test="$styleBorderLine">

							<xsl:variable name="innerLineWidth">
								<xsl:call-template name="point-measure">
									<xsl:with-param name="length" select="substring-before($styleBorderLine,' ' )"/>
								</xsl:call-template>
							</xsl:variable>

							<xsl:variable name="outerLineWidth">
								<xsl:call-template name="point-measure">
									<xsl:with-param name="length"
									  select="substring-after(substring-after($styleBorderLine,' ' ),' ' )"/>
								</xsl:call-template>
							</xsl:variable>

							<xsl:if test="$innerLineWidth = $outerLineWidth">thinThin</xsl:if>
							<xsl:if test="$innerLineWidth > $outerLineWidth">thinThick</xsl:if>
							<xsl:if test="$outerLineWidth > $innerLineWidth  ">thickThin</xsl:if>

						</xsl:when>
					</xsl:choose>
				</xsl:attribute>
			</xsl:if>

			<!--end-->
			
			<!--Dash Styles-->
			<xsl:if test="$shapeStyle/style:graphic-properties/@draw:stroke!='' and $shapeStyle/style:graphic-properties/@draw:stroke!='none'">
				<xsl:variable name="drawStrokeDash" select="$shapeStyle/style:graphic-properties/@draw:stroke-dash"></xsl:variable>
				<xsl:variable name ="strokeStyle" select ="document('styles.xml')/office:document-styles/office:styles/draw:stroke-dash[@draw:name=$drawStrokeDash]"></xsl:variable>
				<xsl:variable name="drawStroke" select="$shapeStyle/style:graphic-properties/@draw:stroke"/>
        <!-- Sona: Arrow feature continuation -->
        <xsl:variable name="Unit1">
          <xsl:call-template name="GetUnit">
            <xsl:with-param name="length" select="$strokeStyle/@draw:dots1-length"/>
          </xsl:call-template>
        </xsl:variable>
        <xsl:variable name="Unit2">
          <xsl:call-template name="GetUnit">
            <xsl:with-param name="length" select="$strokeStyle/@draw:dots2-length"/>
          </xsl:call-template>
        </xsl:variable>
				<xsl:attribute name ="dashstyle">
					<xsl:choose>
						<xsl:when test ="$strokeStyle/@draw:dots1=2 and $strokeStyle/@draw:dots2=1">
							<xsl:value-of select="'longDashDotDot'"/>
						</xsl:when>
            <xsl:when test =" not($strokeStyle/@draw:dots1-length)and substring-before($strokeStyle/@draw:dots2-length,$Unit2) &gt;'0.05'">
							<xsl:value-of select="'longDashDot'"/>
						</xsl:when>
            <xsl:when test ="not($strokeStyle/@draw:dots1-length)and substring-before($strokeStyle/@draw:dots2-length,$Unit2) &lt;= '0.05'">
							<xsl:value-of select="'dashDot'"/>
						</xsl:when>
            <!-- Square Dot-->
            <xsl:when test ="$strokeStyle/@draw:dots1-length=$strokeStyle/@draw:dots2-length and substring-before($strokeStyle/@draw:dots1-length,$Unit1) &lt;= '0.02'">
							<xsl:value-of select="'1 1'"/>
						</xsl:when>
            <xsl:when test ="$strokeStyle/@draw:dots1-length=$strokeStyle/@draw:dots2-length and substring-before($strokeStyle/@draw:dots1-length,$Unit1) &lt;= '0.05'">
							<xsl:value-of select="'dash'"/>
						</xsl:when>
            <xsl:when test ="$strokeStyle/@draw:dots1-length=$strokeStyle/@draw:dots2-length and substring-before($strokeStyle/@draw:dots1-length,$Unit1) &gt; '0.05'">
							<xsl:value-of select="'longDash'"/>
						</xsl:when>
            <!-- Sona : code update for square or round dots-->
            <xsl:when test ="not($strokeStyle/@draw:dots1-length)and not($strokeStyle/@draw:dots2) and $strokeStyle/@draw:dots1">
							<xsl:value-of select="'1 1'"/>
						</xsl:when>
            <!--Sona: Defect #2019464-->
            <xsl:otherwise>
							<xsl:choose>
								<xsl:when test="$drawStroke = 'dash'">
									<xsl:value-of select="'dash'"/>
								</xsl:when>
								<xsl:otherwise>
              <xsl:value-of select="'solid'"/>
            </xsl:otherwise>
					</xsl:choose>
						</xsl:otherwise>
					</xsl:choose>
				</xsl:attribute>
        <xsl:if test ="not($strokeStyle/@draw:dots1-length)and not($strokeStyle/@draw:dots2) and $strokeStyle/@draw:dots1">
					<xsl:attribute name ="endcap">
						<xsl:value-of select ="'round'"/>
					</xsl:attribute>
				</xsl:if>
			</xsl:if>

			<!--Arrow Styles-->
			<!-- Start Arrow-->
      <!-- Sonata: Defect #2151408-->
			<xsl:if test="$shapeStyle/style:graphic-properties/@draw:marker-start and $shapeStyle/style:graphic-properties/@draw:marker-start !=''">
				<xsl:variable name ="drawArrowTypeStart" select="$shapeStyle/style:graphic-properties/@draw:marker-start"></xsl:variable>
				<xsl:variable name ="startArrow" select ="document('styles.xml')/office:document-styles/office:styles/draw:marker[@draw:name=$drawArrowTypeStart]"></xsl:variable>
				<xsl:attribute name ="startarrow">
					<xsl:choose>
						<xsl:when test ="$startArrow/@svg:d='m10 0-10 30h20z'">
							<xsl:value-of select="'block'"/>
						</xsl:when>
						<xsl:when test ="$startArrow/@svg:d='m0 2108v17 17l12 42 30 34 38 21 43 4 29-8 30-21 25-26 13-34 343-1532 339 1520 13 42 29 34 39 21 42 4 42-12 34-30 21-42v-39-12l-4 4-440-1998-9-42-25-39-38-25-43-8-42 8-38 25-26 39-8 42z'">
							<xsl:value-of select="'open'"/>
						</xsl:when>
						<xsl:when test ="$startArrow/@svg:d='m1013 1491 118 89-567-1580-564 1580 114-85 136-68 148-46 161-17 161 13 153 46z'">
							<xsl:value-of select="'classic'"/>
						</xsl:when>
						<xsl:when test ="$startArrow/@svg:d='m462 1118-102-29-102-51-93-72-72-93-51-102-29-102-13-105 13-102 29-106 51-102 72-89 93-72 102-50 102-34 106-9 101 9 106 34 98 50 93 72 72 89 51 102 29 106 13 102-13 105-29 102-51 102-72 93-93 72-98 51-106 29-101 13z'">
							<xsl:value-of select="'oval'"/>
						</xsl:when>
						<xsl:when test ="$startArrow/@svg:d='m0 564 564 567 567-567-567-564z'">
							<xsl:value-of select="'diamond'"/>
						</xsl:when>
            <!-- Sona: Defect #2019464-->
            <xsl:otherwise>
              <xsl:value-of select="'block'"/>
            </xsl:otherwise>
					</xsl:choose>
				</xsl:attribute>
        <!-- Sonata: Defect #2151408-->
        <xsl:if test ="$shapeStyle/style:graphic-properties/@draw:marker-start-width and $shapeStyle/style:graphic-properties/@draw:marker-start-width !=''">
					<xsl:variable name="Unit">
						<xsl:call-template name="GetUnit">
							<xsl:with-param name="length" select="$shapeStyle/style:graphic-properties/@draw:marker-start-width"/>
						</xsl:call-template>
					</xsl:variable>
					<xsl:call-template name ="setArrowSize">
						<xsl:with-param name ="size" select ="substring-before($shapeStyle/style:graphic-properties/@draw:marker-start-width,$Unit)" />
						<xsl:with-param name ="arrType" select ="'start'"></xsl:with-param>
					</xsl:call-template >
				</xsl:if>
			</xsl:if>
			<!-- End Arrow-->
      <!-- Sonata: Defect #2151408-->
			<xsl:if test="$shapeStyle/style:graphic-properties/@draw:marker-end and $shapeStyle/style:graphic-properties/@draw:marker-end !=''">
				<xsl:variable name ="drawArrowTypeEnd" select="$shapeStyle/style:graphic-properties/@draw:marker-end"></xsl:variable>
				<xsl:variable name ="endArrow" select ="document('styles.xml')/office:document-styles/office:styles/draw:marker[@draw:name=$drawArrowTypeEnd]"></xsl:variable>
				<xsl:attribute name ="endarrow">
					<xsl:choose>
						<xsl:when test ="$endArrow/@svg:d='m10 0-10 30h20z'">
							<xsl:value-of select="'block'"/>
						</xsl:when>
						<xsl:when test ="$endArrow/@svg:d='m0 2108v17 17l12 42 30 34 38 21 43 4 29-8 30-21 25-26 13-34 343-1532 339 1520 13 42 29 34 39 21 42 4 42-12 34-30 21-42v-39-12l-4 4-440-1998-9-42-25-39-38-25-43-8-42 8-38 25-26 39-8 42z'">
							<xsl:value-of select="'open'"/>
						</xsl:when>
						<xsl:when test ="$endArrow/@svg:d='m1013 1491 118 89-567-1580-564 1580 114-85 136-68 148-46 161-17 161 13 153 46z'">
							<xsl:value-of select="'classic'"/>
						</xsl:when>
						<xsl:when test ="$endArrow/@svg:d='m462 1118-102-29-102-51-93-72-72-93-51-102-29-102-13-105 13-102 29-106 51-102 72-89 93-72 102-50 102-34 106-9 101 9 106 34 98 50 93 72 72 89 51 102 29 106 13 102-13 105-29 102-51 102-72 93-93 72-98 51-106 29-101 13z'">
							<xsl:value-of select="'oval'"/>
						</xsl:when>
						<xsl:when test ="$endArrow/@svg:d='m0 564 564 567 567-567-567-564z'">
							<xsl:value-of select="'diamond'"/>
						</xsl:when>
            <!-- Sona: Defect #2019464-->
            <xsl:otherwise>
              <xsl:value-of select="'block'"/>
            </xsl:otherwise>
					</xsl:choose>
				</xsl:attribute>
        <!-- Sonata: Defect #2151408-->
				<xsl:if test ="$shapeStyle/style:graphic-properties/@draw:marker-end-width and $shapeStyle/style:graphic-properties/@draw:marker-end-width !=''">
					<xsl:variable name="Unit">
						<xsl:call-template name="GetUnit">
							<xsl:with-param name="length" select="$shapeStyle/style:graphic-properties/@draw:marker-end-width"/>
						</xsl:call-template>
					</xsl:variable>
					<xsl:call-template name ="setArrowSize">
						<xsl:with-param name ="size" select ="substring-before($shapeStyle/style:graphic-properties/@draw:marker-end-width,$Unit)" />
						<xsl:with-param name ="arrType" select ="'end'"></xsl:with-param>
					</xsl:call-template >
				</xsl:if>
			</xsl:if>
		</v:stroke>
	</xsl:template>
	<xsl:template name ="setArrowSize">
		<xsl:param name ="size" />
		<xsl:param name ="arrType"></xsl:param>

		<xsl:choose>
			<xsl:when test="$arrType='start'">
				<xsl:choose>
					<xsl:when test ="($size &lt;= $sm-sm)">
						<xsl:attribute name ="startarrowwidth">
							<xsl:value-of select ="'narrow'"/>
						</xsl:attribute>
						<xsl:attribute name ="startarrowlength">
							<xsl:value-of select ="'short'"/>
						</xsl:attribute>
					</xsl:when>
					<xsl:when test ="($size &gt; $sm-sm) and ($size &lt;= $sm-med)">
						<xsl:attribute name ="startarrowwidth">
							<xsl:value-of select ="'narrow'"/>
						</xsl:attribute>
					</xsl:when>
					<xsl:when test ="($size &gt; $sm-med) and ($size &lt;= $sm-lg)">
						<xsl:attribute name ="startarrowwidth">
							<xsl:value-of select ="'narrow'"/>
						</xsl:attribute>
						<xsl:attribute name ="startarrowlength">
							<xsl:value-of select ="'long'"/>
						</xsl:attribute>
					</xsl:when>
					<xsl:when test ="($size &gt; $sm-lg) and ($size &lt;= $med-sm)">
						<xsl:attribute name ="startarrowlength">
							<xsl:value-of select ="'short'"/>
						</xsl:attribute>
					</xsl:when>
					<!--<xsl:when test ="($size &gt; $med-med) and ($size &lt;= $med-lg)">
        <xsl:attribute name ="startarrowwidth">
          <xsl:value-of select ="'med'"/>
        </xsl:attribute>
        <xsl:attribute name ="startarrowlength">
          <xsl:value-of select ="'lg'"/>
        </xsl:attribute>
      </xsl:when>-->
					<xsl:when test ="($size &gt; $med-lg) and ($size &lt;= $lg-sm)">
						<xsl:attribute name ="startarrowlength">
							<xsl:value-of select ="'long'"/>
						</xsl:attribute>
					</xsl:when>
					<xsl:when test ="($size &gt; $lg-sm) and ($size &lt;= $lg-med)">
						<xsl:attribute name ="startarrowwidth">
							<xsl:value-of select ="'wide'"/>
						</xsl:attribute>
						<xsl:attribute name ="startarrowlength">
							<xsl:value-of select ="'short'"/>
						</xsl:attribute>
					</xsl:when>
					<xsl:when test ="($size &gt; $lg-med) and ($size &lt;= $lg-lg)">
						<xsl:attribute name ="startarrowwidth">
							<xsl:value-of select ="'wide'"/>
						</xsl:attribute>
					</xsl:when>
					<xsl:when test ="($size &gt; $lg-lg)">
						<xsl:attribute name ="startarrowwidth">
							<xsl:value-of select ="'wide'"/>
						</xsl:attribute>
						<xsl:attribute name ="startarrowlength">
							<xsl:value-of select ="'long'"/>
						</xsl:attribute>
					</xsl:when>
				</xsl:choose>
			</xsl:when>
			<xsl:otherwise>
				<xsl:choose>
					<xsl:when test ="($size &lt;= $sm-sm)">
						<xsl:attribute name ="endarrowwidth">
							<xsl:value-of select ="'narrow'"/>
						</xsl:attribute>
						<xsl:attribute name ="endarrowlength">
							<xsl:value-of select ="'short'"/>
						</xsl:attribute>
					</xsl:when>
					<xsl:when test ="($size &gt; $sm-sm) and ($size &lt;= $sm-med)">
						<xsl:attribute name ="endarrowwidth">
							<xsl:value-of select ="'narrow'"/>
						</xsl:attribute>
					</xsl:when>
					<xsl:when test ="($size &gt; $sm-med) and ($size &lt;= $sm-lg)">
						<xsl:attribute name ="endarrowwidth">
							<xsl:value-of select ="'narrow'"/>
						</xsl:attribute>
						<xsl:attribute name ="endarrowlength">
							<xsl:value-of select ="'long'"/>
						</xsl:attribute>
					</xsl:when>
					<xsl:when test ="($size &gt; $sm-lg) and ($size &lt;= $med-sm)">
						<xsl:attribute name ="endarrowlength">
							<xsl:value-of select ="'short'"/>
						</xsl:attribute>
					</xsl:when>
					<!--<xsl:when test ="($size &gt; $med-med) and ($size &lt;= $med-lg)">
        <xsl:attribute name ="startarrowwidth">
          <xsl:value-of select ="'med'"/>
        </xsl:attribute>
        <xsl:attribute name ="startarrowlength">
          <xsl:value-of select ="'lg'"/>
        </xsl:attribute>
      </xsl:when>-->
					<xsl:when test ="($size &gt; $med-lg) and ($size &lt;= $lg-sm)">
						<xsl:attribute name ="endarrowlength">
							<xsl:value-of select ="'long'"/>
						</xsl:attribute>
					</xsl:when>
					<xsl:when test ="($size &gt; $lg-sm) and ($size &lt;= $lg-med)">
						<xsl:attribute name ="endarrowwidth">
							<xsl:value-of select ="'wide'"/>
						</xsl:attribute>
						<xsl:attribute name ="endarrowlength">
							<xsl:value-of select ="'short'"/>
						</xsl:attribute>
					</xsl:when>
					<xsl:when test ="($size &gt; $lg-med) and ($size &lt;= $lg-lg)">
						<xsl:attribute name ="endarrowwidth">
							<xsl:value-of select ="'wide'"/>
						</xsl:attribute>
					</xsl:when>
					<xsl:when test ="($size &gt; $lg-lg)">
						<xsl:attribute name ="endarrowwidth">
							<xsl:value-of select ="'wide'"/>
						</xsl:attribute>
						<xsl:attribute name ="endarrowlength">
							<xsl:value-of select ="'long'"/>
						</xsl:attribute>
					</xsl:when>
				</xsl:choose>
			</xsl:otherwise>
		</xsl:choose>

	</xsl:template>
</xsl:stylesheet>