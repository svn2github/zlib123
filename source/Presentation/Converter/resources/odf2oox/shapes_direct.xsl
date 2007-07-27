<?xml version="1.0" encoding="utf-8" ?>
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
  xmlns:script="urn:oasis:names:tc:opendocument:xmlns:script:1.0"
  xmlns:xlink="http://www.w3.org/1999/xlink"
  exclude-result-prefixes="odf style text number draw page r">
  <xsl:import href="common.xsl"/>

  <!-- Shape constants-->
  <xsl:variable name ="dot">
    <xsl:value-of select ="'0.07'"/>
  </xsl:variable>
  <xsl:variable name ="dash">
    <xsl:value-of select ="'0.282'"/>
  </xsl:variable>
  <xsl:variable name ="longDash">
    <xsl:value-of select ="'0.564'"/>
  </xsl:variable>
  <xsl:variable name ="distance">
    <xsl:value-of select ="'0.211'"/>
  </xsl:variable>
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

  <!-- Template for Shapes in direct conversion -->
  <xsl:template name ="shapes">
		<xsl:param name ="fileName" />
		<xsl:param name ="styleName" />
    <!-- Rectangle -->
    <xsl:for-each select ="draw:rect">
      <xsl:call-template name ="CreateShape">
		<xsl:with-param name ="fileName" select ="$fileName" />
        <xsl:with-param name ="shapeName" select="'Rectangle '" />
      </xsl:call-template>
    </xsl:for-each>

    <!-- Text box-->
    <xsl:for-each select ="draw:frame">
      <xsl:if test ="(draw:text-box) and not(@presentation:style-name) and not(@presentation:class)">
        <xsl:call-template name ="CreateShape">
		  <xsl:with-param name ="fileName" select ="$fileName" />
          <xsl:with-param name ="shapeName" select="'TextBox '" />
        </xsl:call-template>
      </xsl:if>
    </xsl:for-each>

   <!--  Ellipse   -->
   <!-- <xsl:if test="@draw:kind != ''">  -->
   <!--  code given by pradeep for the bug 1746358   -->
   <xsl:for-each select="draw:ellipse[not(@draw:kind)]">
	   <xsl:call-template name="CreateShape">
		  <xsl:with-param name="shapeName" select="'Oval '" />
	  </xsl:call-template>
  </xsl:for-each>
   <!--  End of code   -->
   <!-- </xsl:if>  -->
	<!-- Custom shapes -->
    <xsl:for-each select ="draw:custom-shape">
      <xsl:variable name ="shapeName">
        <xsl:choose>
          <!-- Text Box -->
          <xsl:when test ="(draw:enhanced-geometry/@draw:type='mso-spt202') and (draw:enhanced-geometry/@draw:enhanced-path='M 0 0 L 21600 0 21600 21600 0 21600 0 0 Z N')">
            <xsl:value-of select ="'TextBox Custom '"/>
          </xsl:when>
          <!-- Oval -->
          <xsl:when test ="(draw:enhanced-geometry/@draw:type='ellipse')">
            <xsl:value-of select ="'Oval Custom '"/>
          </xsl:when>
          <!-- Isosceles Triangle -->
          <xsl:when test ="(draw:enhanced-geometry/@draw:type='isosceles-triangle')">
            <xsl:value-of select ="'Isosceles Triangle '"/>
          </xsl:when>
          <!-- Right Arrow (Added by Mathi as on 4/7/2007 -->
          <xsl:when test = "(draw:enhanced-geometry/@draw:type='right-arrow')">
            <xsl:value-of select ="'Right Arrow '"/>
          </xsl:when>
          <!-- Left Arrow (Added by Mathi as on 5/7/2007) -->
          <xsl:when test ="(draw:enhanced-geometry/@draw:type='left-arrow')">
            <xsl:value-of select ="'Left Arrow '"/>
          </xsl:when>
			<!-- Up Arrow (Added by Mathi as on 19/7/2007)-->
			<xsl:when test ="(draw:enhanced-geometry/@draw:type='up-arrow')">
				<xsl:value-of select ="'Up Arrow '"/>
			</xsl:when>
			<!-- Down Arrow (Added by Mathi as on 19/7/2007)-->
			<xsl:when test ="(draw:enhanced-geometry/@draw:type='down-arrow')">
				<xsl:value-of select ="'Down Arrow '"/>
			</xsl:when>
			<!-- Left-Right Arrow (Added by Mathi as on 19/7/2007)-->
			<xsl:when test ="(draw:enhanced-geometry/@draw:type='left-right-arrow')">
				<xsl:value-of select ="'Left-Right Arrow '"/>
			</xsl:when>
			<!-- Up-Down Arrow (Added by Mathi as on 19/7/2007)-->
			<xsl:when test ="(draw:enhanced-geometry/@draw:type='up-down-arrow')">
				<xsl:value-of select ="'Up-Down Arrow '"/>
			</xsl:when>
          <!-- Right Triangle -->
          <xsl:when test ="(draw:enhanced-geometry/@draw:type='right-triangle')">
            <xsl:value-of select ="'Right Triangle '"/>
          </xsl:when>
          <!-- Parallelogram -->
          <xsl:when test ="(draw:enhanced-geometry/@draw:type='parallelogram')">
            <xsl:value-of select ="'Parallelogram '"/>
          </xsl:when>
		  <!-- Trapezoid (Added by A.Mathi as on 24/07/2007) -->
		  <xsl:when test ="(draw:enhanced-geometry/@draw:type='trapezoid')">
			<xsl:value-of select ="'Trapezoid '"/>
		  </xsl:when>
          <!-- Diamond -->
          <xsl:when test ="(draw:enhanced-geometry/@draw:type='diamond')and
									 (draw:enhanced-geometry/@draw:enhanced-path='M 10800 0 L 21600 10800 10800 21600 0 10800 10800 0 Z N')">
            <xsl:value-of select ="'Diamond '"/>
          </xsl:when>
          <!-- Regular Pentagon -->
          <xsl:when test ="(draw:enhanced-geometry/@draw:type='pentagon') and
									 (draw:enhanced-geometry/@draw:enhanced-path='M 10800 0 L 0 8260 4230 21600 17370 21600 21600 8260 10800 0 Z N')">
            <xsl:value-of select ="'Regular Pentagon '"/>
          </xsl:when>
          <!-- Hexagon -->
          <xsl:when test ="(draw:enhanced-geometry/@draw:type='hexagon')">
            <xsl:value-of select ="'Hexagon '"/>
          </xsl:when>
          <!-- Octagon -->
          <xsl:when test ="(draw:enhanced-geometry/@draw:type='octagon')">
            <xsl:value-of select ="'Octagon '"/>
          </xsl:when>
          <!-- Cube -->
          <xsl:when test ="(draw:enhanced-geometry/@draw:type='cube')">
            <xsl:value-of select ="'Cube '"/>
          </xsl:when>
          <!-- Can -->
          <xsl:when test ="(draw:enhanced-geometry/@draw:type='can')">
            <xsl:value-of select ="'Can '"/>
          </xsl:when>
		  <!-- Circular Arrow -->
		  <xsl:when test ="(draw:enhanced-geometry/@draw:type='circular-arrow')">
			<xsl:value-of select ="'circular-arrow'"/>
		  </xsl:when>
		  <!-- LeftUp Arrow -->
		  <xsl:when test ="(draw:enhanced-geometry/@draw:type='mso-spt89') and
								 (draw:enhanced-geometry/@draw:enhanced-path='M 0 ?f5 L ?f2 ?f0 ?f2 ?f7 ?f7 ?f7 ?f7 ?f2 ?f0 ?f2 ?f5 0 21600 ?f2 ?f1 ?f2 ?f1 ?f1 ?f2 ?f1 ?f2 21600 Z N')">
			<xsl:value-of select ="'leftUpArrow'"/>
		  </xsl:when>
		  <!-- BentUp Arrow -->
		  <xsl:when test ="(draw:enhanced-geometry/@draw:type='non-primitive') and
								 (draw:enhanced-geometry/@draw:enhanced-path='M 517 247 L 517 415 264 415 264 0 0 0 0 680 517 680 517 854 841 547 517 247 Z N')">
		  <xsl:value-of select ="'bentUpArrow'"/>
		  </xsl:when>
		  <!-- Cross (Added by Mathi on 19/7/2007)-->
			<xsl:when test ="(draw:enhanced-geometry/@draw:type='cross')">
				<xsl:value-of select ="'plus '"/>
			</xsl:when>
		  <!-- "No Symbol" (Added by A.Mathi as on 19/07/2007)-->
			<xsl:when test ="(draw:enhanced-geometry/@draw:type='forbidden')">
				<xsl:value-of select ="'noSmoking '"/>
			</xsl:when>
		  <!-- Bent Arrow (Added by A.Mathi as on 23/07/2007) -->
			<xsl:when test ="(draw:enhanced-geometry/@draw:enhanced-path='M 517 247 L 517 415 264 415 264 0 0 0 0 680 517 680 517 854 841 547 517 247 Z N')">
				<xsl:value-of select ="'bentArrow '"/>
			</xsl:when>
		  <!--  Folded Corner (Added by A.Mathi as on 19/07/2007) -->
			<xsl:when test ="(draw:enhanced-geometry/@draw:type='paper')">
				<xsl:value-of select ="'foldedCorner '"/>
			</xsl:when>
		  <!--  Lightning Bolt (Added by A.Mathi as on 20/07/2007) -->
			<xsl:when test ="(draw:enhanced-geometry/@draw:enhanced-path='M 640 233 L 221 293 506 12 367 0 29 406 431 347 145 645 99 520 0 861 326 765 209 711 640 233 640 233 Z N')">
				<xsl:value-of select ="'lightningBolt '"/>
			</xsl:when>
		  <!--  Explosion 1 (Added by A.Mathi as on 20/07/2007) -->
			<xsl:when test ="(draw:enhanced-geometry/@draw:type='bang')">
				<xsl:value-of select ="'irregularSeal1 '"/>
			</xsl:when>
		  <!-- Left Bracket (Added by A.Mathi as on 20/07/2007) -->
			<xsl:when test ="(draw:enhanced-geometry/@draw:type='left-bracket')">
				<xsl:value-of select ="'Left Bracket '"/>
			</xsl:when>
		  <!-- Right Bracket (Added by A.Mathi as on 20/07/2007) -->
			<xsl:when test ="(draw:enhanced-geometry/@draw:type='right-bracket')">
				<xsl:value-of select ="'Right Bracket '"/>
			</xsl:when>
		  <!-- Left Brace (Added by A.Mathi as on 20/07/2007) -->
			<xsl:when test ="(draw:enhanced-geometry/@draw:type='left-brace')">
				<xsl:value-of select ="'Left Brace '"/>
			</xsl:when>
		  <!-- Right Brace (Added by A.Mathi as on 23/07/2007) -->
			<xsl:when test ="(draw:enhanced-geometry/@draw:type='right-brace')">
				<xsl:value-of select ="'Right Brace '"/>
			</xsl:when>
		  <!-- Rectangular Callout (Added by A.Mathi as on 23/07/2007) -->
			<xsl:when test ="(draw:enhanced-geometry/@draw:type='rectangular-callout')">
				<xsl:value-of select ="'Rectangular Callout '"/>
			</xsl:when>
		  <!-- Rounded Rectangular Callout (Added by A.Mathi as on 23/07/2007) -->
			<xsl:when test ="(draw:enhanced-geometry/@draw:type='round-rectangular-callout')">
				<xsl:value-of select ="'wedgeRoundRectCallout '"/>
			</xsl:when>
		  <!-- Oval Callout (Added by A.Mathi as on 23/07/2007) -->
			<xsl:when test ="(draw:enhanced-geometry/@draw:type='round-callout')">
				<xsl:value-of select ="'wedgeEllipseCallout '"/>
			</xsl:when>
		  <!-- Cloud Callout (Added by A.Mathi as on 23/07/2007) -->
			<xsl:when test ="(draw:enhanced-geometry/@draw:type='cloud-callout')">
				<xsl:value-of select ="'Cloud Callout '"/>
			</xsl:when>
		  <!-- U-Turn Arrow (Added by A.Mathi as on 23/07/2007) -->
			<xsl:when test ="(draw:enhanced-geometry/@draw:type='circular-arrow')">
				<xsl:value-of select ="'U-Turn Arrow '"/>
          </xsl:when>
          <!-- Rectangle -->
          <xsl:when test ="(draw:enhanced-geometry/@draw:type='rectangle') and (draw:enhanced-geometry/@draw:enhanced-path='M 0 0 L 21600 0 21600 21600 0 21600 0 0 Z N')">
            <xsl:value-of select ="'Rectangle Custom '"/>
          </xsl:when>
          <!-- Rounded Rectangle -->
          <xsl:when test ="(draw:enhanced-geometry/@draw:type='round-rectangle')">
            <xsl:value-of select ="'Rounded Rectangle '"/>
          </xsl:when>
          <!-- Snip Single Corner Rectangle -->
          <xsl:when test ="(draw:enhanced-geometry/@draw:type='flowchart-card') and 
									 (draw:enhanced-geometry/@draw:mirror-horizontal='true') and
									 (draw:enhanced-geometry/@draw:enhanced-path='M 4300 0 L 21600 0 21600 21600 0 21600 0 4300 4300 0 Z N')">
            <xsl:value-of select ="'Snip Single Corner Rectangle '"/>
          </xsl:when>
          <!-- Snip Same Side Corner Rectangle -->
          <xsl:when test ="(draw:enhanced-geometry/@draw:type='mso-spt100') and 
									 (draw:enhanced-geometry/@draw:enhanced-path='M 50801 0 L 863599 0 L 914400 50801 L 914400 304800 L 914400 304800 L 0 304800 L 0 304800 L 0 50801 Z N')">
            <xsl:value-of select ="'Snip Same Side Corner Rectangle '"/>
          </xsl:when>
          <!-- Snip Diagonal Corner Rectangle -->
          <xsl:when test ="(draw:enhanced-geometry/@draw:type='mso-spt100') and
									 (draw:enhanced-geometry/@draw:enhanced-path='M 0 0 L 876299 0 L 914400 38101 L 914400 228600 L 914400 228600 L 38101 228600 L 0 190499 L 0 0 Z N')">
            <xsl:value-of select ="'Snip Diagonal Corner Rectangle '"/>
          </xsl:when>
          <!-- Snip and Round Single Corner Rectangle -->
          <xsl:when test ="(draw:enhanced-geometry/@draw:type='mso-spt100') and
									 (draw:enhanced-geometry/@draw:enhanced-path='M 38101 0 L 876299 0 L 914400 38101 L 914400 228600 L 0 228600 L 0 38101 W 0 0 76202 76202 0 38101 38101 0 Z N')">
            <xsl:value-of select ="'Snip and Round Single Corner Rectangle '"/>
          </xsl:when>
          <!-- Round Single Corner Rectangle -->
          <xsl:when test ="(draw:enhanced-geometry/@draw:type='mso-spt100') and
									 (draw:enhanced-geometry/@draw:enhanced-path='M 0 0 L 876299 0 W 838198 0 914400 76202 876299 0 914400 38101 L 914400 228600 L 0 228600 Z N')">
            <xsl:value-of select ="'Round Single Corner Rectangle '"/>
          </xsl:when>
          <!-- Round Same Side Corner Rectangle -->
          <xsl:when test ="(draw:enhanced-geometry/@draw:type='mso-spt100') and
									 (draw:enhanced-geometry/@draw:enhanced-path='M 152403 0 L 761997 0 W 609594 0 914400 304806 761997 0 914400 152403 L 914400 914400 L 0 914400 L 0 152403 W 0 0 304806 304806 0 152403 152403 0 Z N')">
            <xsl:value-of select ="'Round Same Side Corner Rectangle '"/>
          </xsl:when>
          <!-- Round Diagonal Corner Rectangle -->
          <xsl:when test ="(draw:enhanced-geometry/@draw:type='mso-spt100') and
									 (draw:enhanced-geometry/@draw:enhanced-path='M 254005 0 L 2286000 0 L 2286000 1269995 W 1777990 1015990 2286000 1524000 2286000 1269995 2031995 1524000 L 0 1524000 L 0 254005 W 0 0 508010 508010 0 254005 254005 0 Z N')">
            <xsl:value-of select ="'Round Diagonal Corner Rectangle '"/>
          </xsl:when>

          <!-- Flow Chart: Process -->
          <xsl:when test ="(draw:enhanced-geometry/@draw:type='flowchart-process')">
            <xsl:value-of select ="'Flowchart: Process '"/>
          </xsl:when>
          <!-- Flow Chart: Alternate Process -->
          <xsl:when test ="(draw:enhanced-geometry/@draw:type='flowchart-alternate-process')">
            <xsl:value-of select ="'Flowchart: Alternate Process '"/>
          </xsl:when>
          <!-- Flow Chart: Decision -->
          <xsl:when test ="(draw:enhanced-geometry/@draw:type='flowchart-decision')">
            <xsl:value-of select ="'Flowchart: Decision '"/>
          </xsl:when>
          <!-- Flow Chart: Data -->
          <xsl:when test ="(draw:enhanced-geometry/@draw:type='flowchart-data')">
            <xsl:value-of select ="'Flowchart: Data '"/>
          </xsl:when>
          <!-- Flow Chart: Predefined-process -->
          <xsl:when test ="(draw:enhanced-geometry/@draw:type='flowchart-predefined-process')">
            <xsl:value-of select ="'Flowchart: Predefined Process '"/>
          </xsl:when>
          <!-- Flow Chart: Internal-storage -->
          <xsl:when test ="(draw:enhanced-geometry/@draw:type='flowchart-internal-storage')">
            <xsl:value-of select ="'Flowchart: Internal Storage '"/>
          </xsl:when>
          <!-- Flow Chart: Document -->
          <xsl:when test ="(draw:enhanced-geometry/@draw:type='flowchart-document')">
            <xsl:value-of select ="'Flowchart: Document '"/>
          </xsl:when>
          <!-- Flow Chart: Multidocument -->
          <xsl:when test ="(draw:enhanced-geometry/@draw:type='flowchart-multidocument')">
            <xsl:value-of select ="'Flowchart: Multi document '"/>
          </xsl:when>
          <!-- Flow Chart: Terminator -->
          <xsl:when test ="(draw:enhanced-geometry/@draw:type='flowchart-terminator')">
            <xsl:value-of select ="'Flowchart: Terminator '"/>
          </xsl:when>
          <!-- Flow Chart: Preparation -->
          <xsl:when test ="(draw:enhanced-geometry/@draw:type='flowchart-preparation')">
            <xsl:value-of select ="'Flowchart: Preparation '"/>
          </xsl:when>
          <!-- Flow Chart: Manual-input -->
          <xsl:when test ="(draw:enhanced-geometry/@draw:type='flowchart-manual-input')">
            <xsl:value-of select ="'Flowchart: Manual Input '"/>
          </xsl:when>
          <!-- Flow Chart: Manual-operation -->
          <xsl:when test ="(draw:enhanced-geometry/@draw:type='flowchart-manual-operation')">
            <xsl:value-of select ="'Flowchart: Manual Operation '"/>
          </xsl:when>
          <!-- Flow Chart: Connector -->
          <xsl:when test ="(draw:enhanced-geometry/@draw:type='flowchart-connector')">
            <xsl:value-of select ="'Flowchart: Connector '"/>
          </xsl:when>
          <!-- Flow Chart: Off-page-connector -->
          <xsl:when test ="(draw:enhanced-geometry/@draw:type='flowchart-off-page-connector')">
            <xsl:value-of select ="'Flowchart: Off-page Connector '"/>
          </xsl:when>
          <!-- Flow Chart: Card -->
          <xsl:when test ="(draw:enhanced-geometry/@draw:type='flowchart-card')">
            <xsl:value-of select ="'Flowchart: Card '"/>
          </xsl:when>
          <!-- Flow Chart: Punched-tape -->
          <xsl:when test ="(draw:enhanced-geometry/@draw:type='flowchart-punched-tape')">
            <xsl:value-of select ="'Flowchart: Punched Tape '"/>
          </xsl:when>
          <!-- Flow Chart: Summing-junction -->
          <xsl:when test ="(draw:enhanced-geometry/@draw:type='flowchart-summing-junction')">
            <xsl:value-of select ="'Flowchart: Summing Junction '"/>
          </xsl:when>
          <!-- Flow Chart: Or -->
          <xsl:when test ="(draw:enhanced-geometry/@draw:type='flowchart-or')">
            <xsl:value-of select ="'Flowchart: Or '"/>
          </xsl:when>
          <!-- Flow Chart: Collate -->
          <xsl:when test ="(draw:enhanced-geometry/@draw:type='flowchart-collate')">
            <xsl:value-of select ="'Flowchart: Collate '"/>
          </xsl:when>
          <!-- Flow Chart: Sort -->
          <xsl:when test ="(draw:enhanced-geometry/@draw:type='flowchart-sort')">
            <xsl:value-of select ="'Flowchart: Sort '"/>
          </xsl:when>
          <!-- Flow Chart: Extract -->
          <xsl:when test ="(draw:enhanced-geometry/@draw:type='flowchart-extract')">
            <xsl:value-of select ="'Flowchart: Extract '"/>
          </xsl:when>
          <!-- Flow Chart: Merge -->
          <xsl:when test ="(draw:enhanced-geometry/@draw:type='flowchart-merge')">
            <xsl:value-of select ="'Flowchart: Merge '"/>
          </xsl:when>
          <!-- Flow Chart: Stored-data -->
          <xsl:when test ="(draw:enhanced-geometry/@draw:type='flowchart-stored-data')">
            <xsl:value-of select ="'Flowchart: Stored Data '"/>
          </xsl:when>
          <!-- Flow Chart: Delay-->
          <xsl:when test ="(draw:enhanced-geometry/@draw:type='flowchart-delay')">
            <xsl:value-of select ="'Flowchart: Delay '"/>
          </xsl:when>
          <!-- Flow Chart: Sequential-access -->
          <xsl:when test ="(draw:enhanced-geometry/@draw:type='flowchart-sequential-access')">
            <xsl:value-of select ="'Flowchart: Sequential Access Storage '"/>
          </xsl:when>
          <!-- Flow Chart: Direct-access-storage -->
          <xsl:when test ="(draw:enhanced-geometry/@draw:type='flowchart-direct-access-storage')">
            <xsl:value-of select ="'Flowchart: Direct Access Storage'"/>
          </xsl:when>
          <!-- Flow Chart: Magnetic-disk -->
          <xsl:when test ="(draw:enhanced-geometry/@draw:type='flowchart-magnetic-disk')">
            <xsl:value-of select ="'Flowchart: Magnetic Disk '"/>
          </xsl:when>
          <!-- Flow Chart: Display -->
          <xsl:when test ="(draw:enhanced-geometry/@draw:type='flowchart-display')">
            <xsl:value-of select ="'Flowchart: Display '"/>
          </xsl:when>
        </xsl:choose>
      </xsl:variable>

      <!-- code for the bug 1746360 -->
      <xsl:if test="$shapeName != ''">
        <xsl:call-template name ="CreateShape">
					<xsl:with-param name ="fileName" select ="$fileName" />
          <xsl:with-param name ="shapeName" select="$shapeName" />
          <xsl:with-param name ="customShape" select ="'true'" />
        </xsl:call-template>
      </xsl:if>
      <!-- End of code for the bug 1746360 -->

    </xsl:for-each>

    <!-- Line -->
    <xsl:for-each select ="draw:line">
      <xsl:call-template name ="drawLine">
				<xsl:with-param name ="fileName" select ="$fileName" />
        <xsl:with-param name ="connectorType" select ="'Straight Connector '" />
      </xsl:call-template>
    </xsl:for-each>

    <!-- Connectors -->
    <xsl:for-each select ="draw:connector">
      <xsl:variable name ="type">
        <xsl:choose>
          <xsl:when test ="@draw:type='line'">
            <xsl:value-of select ="'Straight Arrow Connector '"/>
          </xsl:when>
          <xsl:when test ="@draw:type='curve'">
            <xsl:value-of select ="'Curved Connector '"/>
          </xsl:when>
          <xsl:otherwise>
            <xsl:value-of select ="'Elbow Connector '"/>
          </xsl:otherwise>
        </xsl:choose>
      </xsl:variable>
      <xsl:call-template name ="drawLine">
		<xsl:with-param name ="fileName" select ="$fileName" />
        <xsl:with-param name ="connectorType" select ="$type" />
      </xsl:call-template>
    </xsl:for-each>

    <!-- Grouping/Ungrouping-->
    <xsl:for-each select="draw:g">
		<xsl:call-template name="shapes">
			<xsl:with-param name ="fileName" select ="$fileName" />
		</xsl:call-template >
    </xsl:for-each>

  </xsl:template>
  <!-- Create p:sp node for shape -->
  <xsl:template name ="CreateShape">
		<xsl:param name ="fileName" />
    <xsl:param name ="shapeName" />
    <xsl:param name ="customShape" />
    <p:sp>
      <p:nvSpPr>
        <p:cNvPr>
          <xsl:attribute name ="id">
            <xsl:value-of select ="position()"/>
          </xsl:attribute>
          <xsl:attribute name ="name">
            <xsl:value-of select ="concat($shapeName, position())"/>
          </xsl:attribute>
          <!-- ADDED BY LOHITH - HYPER LINKS FOR SHAPES-->
          <xsl:variable name="PostionCount">
            <xsl:value-of select="position()"/>
          </xsl:variable>
          <xsl:variable name="ShapeType">
            <xsl:if test="contains(.,'rect')">
              <xsl:value-of select="'RectAtachFileId'"/>
            </xsl:if>
            <xsl:if test="not(contains(.,'rect')) and not(contains($shapeName,'Text'))">
              <xsl:value-of select="'ShapeFileId'"/>
            </xsl:if>
            <xsl:if test="contains($shapeName,'Text')">
              <xsl:if test="not(contains($shapeName,'Custom'))">
                <xsl:value-of select="'TxtBoxAtchFileId'"/>
              </xsl:if>
              <xsl:if test="contains($shapeName,'Custom')">
                <xsl:value-of select="'ShapeFileId'"/>
              </xsl:if>
            </xsl:if>
          </xsl:variable>
          <xsl:if test="office:event-listeners">
            <xsl:for-each select ="office:event-listeners/presentation:event-listener">
              <xsl:if test="@script:event-name[contains(.,'dom:click')]">
                <a:hlinkClick>
                  <xsl:choose>
                    <!--Go to previous slide-->
                    <xsl:when test="@presentation:action[ contains(.,'previous-page')]">
                      <xsl:attribute name="action">
                        <xsl:value-of select="'ppaction://hlinkshowjump?jump=previousslide'"/>
                      </xsl:attribute>
                    </xsl:when>
                    <!--Go to Next slide-->
                    <xsl:when test="@presentation:action[ contains(.,'next-page')]">
                      <xsl:attribute name="action">
                        <xsl:value-of select="'ppaction://hlinkshowjump?jump=nextslide'"/>
                      </xsl:attribute>
                    </xsl:when>
                    <!--Go to First slide-->
                    <xsl:when test="@presentation:action[ contains(.,'first-page')]">
                      <xsl:attribute name="action">
                        <xsl:value-of select="'ppaction://hlinkshowjump?jump=firstslide'"/>
                      </xsl:attribute>
                    </xsl:when>
                    <!--Go to Last slide-->
                    <xsl:when test="@presentation:action[ contains(.,'last-page')]">
                      <xsl:attribute name="action">
                        <xsl:value-of select="'ppaction://hlinkshowjump?jump=lastslide'"/>
                      </xsl:attribute>
                    </xsl:when>
                    <!--End Presentation-->
                    <xsl:when test="@presentation:action[ contains(.,'stop')]">
                      <xsl:attribute name="action">
                        <xsl:value-of select="'ppaction://hlinkshowjump?jump=endshow'"/>
                      </xsl:attribute>
                    </xsl:when>
                    <!--Run program-->
                    <xsl:when test="@xlink:href and @presentation:action[ contains(.,'execute')]">
                      <xsl:attribute name="action">
                        <xsl:value-of select="'ppaction://program'"/>
                      </xsl:attribute>
                    </xsl:when>
                    <!--Go to slide-->
                    <xsl:when test="@xlink:href[ contains(.,'#page')]">
                      <xsl:attribute name="action">
                        <xsl:value-of select="'ppaction://hlinksldjump'"/>
                      </xsl:attribute>
                    </xsl:when>
                    <!--Go to document-->
                    <xsl:when test="@xlink:href and @presentation:action[ contains(.,'show')] ">
                      <xsl:if test="not(@xlink:href[ contains(.,'#page')])">
                        <xsl:attribute name="action">
                          <xsl:value-of select="'ppaction://hlinkfile'"/>
                        </xsl:attribute>
                      </xsl:if>
                    </xsl:when>
                  </xsl:choose>
                  <!--set value for attribute r:id-->
                  <xsl:choose>
                    <!--For jump to next,previous,first,last-->
                    <xsl:when test="@presentation:action[ contains(.,'page') or contains(.,'stop')]">
                      <xsl:attribute name="r:id">
                        <xsl:value-of select="''"/>
                      </xsl:attribute>
                    </xsl:when>
                    <!--For Run program & got to slide-->
                    <xsl:when test="@xlink:href[contains(.,'#page')]">
                      <xsl:attribute name="r:id">
                        <xsl:value-of select="concat($ShapeType,$PostionCount)"/>
                      </xsl:attribute>
                    </xsl:when>
                    <!--For Go to document-->
                    <xsl:when test="@xlink:href and @presentation:action[ contains(.,'show')] ">
                      <xsl:if test="not(@xlink:href[ contains(.,'#page')])">
                        <xsl:attribute name="r:id">
                          <xsl:value-of select="concat($ShapeType,$PostionCount)"/>
                        </xsl:attribute>
                      </xsl:if>
                    </xsl:when>
                    <!--For Go to Run Programs-->
                    <xsl:when test="@xlink:href and @presentation:action[ contains(.,'execute')]">
                      <xsl:attribute name="r:id">
                        <xsl:value-of select="concat($ShapeType,$PostionCount)"/>
                      </xsl:attribute>
                    </xsl:when>
                  </xsl:choose>
                  <!-- Play Sound -->
                  <xsl:if test="@presentation:action[ contains(.,'sound')]">
                    <a:snd>
                      <xsl:variable name="varMediaFilePath">
                        <xsl:if test="presentation:sound/@xlink:href [ contains(.,'../')]">
                          <xsl:value-of select="presentation:sound/@xlink:href" />
                        </xsl:if>
                        <xsl:if test="not(presentation:sound/@xlink:href [ contains(.,'../')])">
                          <xsl:value-of select="substring-after(presentation:sound/@xlink:href,'/')" />
                        </xsl:if>
                      </xsl:variable>
                      <xsl:variable name="varFileRelId">
                        <xsl:value-of select="translate(translate(translate(translate(translate($varMediaFilePath,'/','_'),'..','_'),'.','_'),':','_'),'%20D','_')"/>
                      </xsl:variable>
                      <xsl:attribute name="r:embed">
                        <xsl:value-of select="$varFileRelId"/>
                      </xsl:attribute>
                      <xsl:attribute name="name">
                        <xsl:value-of select="concat('SoundFileForShape',$PostionCount)"/>
                      </xsl:attribute>
                      <xsl:attribute name="builtIn">
                        <xsl:value-of select='1'/>
                      </xsl:attribute>
                      <pzip:extract pzip:source="{$varMediaFilePath}" pzip:target="{concat('ppt/media/',$varFileRelId,'.wav')}" />
                    </a:snd>
                  </xsl:if>
                </a:hlinkClick>
              </xsl:if>
            </xsl:for-each>
          </xsl:if>
          <!-- ADDED BY LOHITH - HYPER LINKS FOR SHAPES-->
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
					<xsl:with-param name ="fileName" select ="$fileName" />
					<xsl:with-param name ="gr" >
					<xsl:if test ="@draw:style-name">
						<xsl:value-of select ="@draw:style-name"/>
					</xsl:if>
					<xsl:if test ="@presentation:style-name">
						<xsl:value-of select ="@presentation:style-name"/>
					</xsl:if >
					</xsl:with-param >
        </xsl:call-template>
      </p:spPr>
      <p:style>
        <a:lnRef idx="2">
          <a:schemeClr val="dk1" />
        </a:lnRef>
        <a:fillRef idx="1">
          <a:srgbClr val="99ccff">
            <xsl:call-template name ="getShade">
							<xsl:with-param name ="fileName" select ="$fileName" />
							<xsl:with-param name ="gr" >
								<xsl:if test ="@draw:style-name">
									<xsl:value-of select ="@draw:style-name"/>
								</xsl:if>
								<xsl:if test ="@presentation:style-name">
									<xsl:value-of select ="@presentation:style-name"/>
								</xsl:if >
							</xsl:with-param >
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
					<xsl:with-param name ="fileName" select ="$fileName" />
					<xsl:with-param name ="gr" >
						<xsl:if test ="@draw:style-name">
							<xsl:value-of select ="@draw:style-name"/>
						</xsl:if>
						<xsl:if test ="@presentation:style-name">
							<xsl:value-of select ="@presentation:style-name"/>
						</xsl:if >
					</xsl:with-param >
          <xsl:with-param name ="customShape" select="$customShape" />
        </xsl:call-template>

        <xsl:choose>
          <xsl:when test ="draw:text-box">
            <xsl:for-each select ="draw:text-box">
              <!-- Paremeter added by vijayeta,get master page name, dated:11-7-07-->
              <xsl:variable name ="masterPageName" select ="./parent::node()/parent::node()/@draw:master-page-name"/>
              <xsl:call-template name ="processShapeText">
								<xsl:with-param name ="fileName" select ="$fileName" />								

								<xsl:with-param name ="gr" >
									<xsl:if test ="parent::node()/@draw:style-name">
										<xsl:value-of select ="parent::node()/@draw:style-name"/>
									</xsl:if>
									<xsl:if test ="parent::node()/@presentation:style-name">
										<xsl:value-of select ="parent::node()/@presentation:style-name"/>
									</xsl:if >
								</xsl:with-param >
                <xsl:with-param name ="masterPageName" select ="$masterPageName" />
              </xsl:call-template>
            </xsl:for-each>
          </xsl:when>
          <xsl:otherwise>
            <xsl:call-template name ="processShapeText">
							<xsl:with-param name ="fileName" select ="$fileName" />
							<xsl:with-param name ="gr" >
								<xsl:if test ="@draw:style-name">
									<xsl:value-of select ="@draw:style-name"/>
								</xsl:if>
								<xsl:if test ="@presentation:style-name">
									<xsl:value-of select ="@presentation:style-name"/>
								</xsl:if >
							</xsl:with-param >
            </xsl:call-template>
          </xsl:otherwise>
        </xsl:choose>
      </p:txBody>
    </p:sp>
  </xsl:template>
  <!-- Draw shape -->
  <xsl:template name ="drawShape">
    <a:xfrm>
      <xsl:variable name ="angle">
        <xsl:value-of select="substring-after(substring-before(substring-before(@draw:transform,'translate'),')'),'(')" />
      </xsl:variable>
      <xsl:variable name ="x2">
        <xsl:value-of select="substring-before(substring-before(substring-before(substring-after(substring-after(@draw:transform,'translate'),'('),')'),' '),'cm')" />
      </xsl:variable>
      <xsl:variable name ="y2">
        <xsl:value-of select="substring-before(substring-after(substring-before(substring-after(substring-after(@draw:transform,'translate'),'('),')'),' '),'cm')" />
      </xsl:variable>
      <xsl:if test="@draw:transform">
        <xsl:attribute name ="rot">
          <xsl:value-of select ="concat('draw-transform:ROT:',substring-before(@svg:width,'cm'), ':',
																   substring-before(@svg:height,'cm'), ':', 
																   $x2, ':', 
                                   $y2, ':', 
																   $angle)"/>
        </xsl:attribute>
      </xsl:if>
      <a:off>
        <xsl:if test="@draw:transform">
          <xsl:attribute name ="x">
          <xsl:value-of select ="concat('draw-transform:X:',substring-before(@svg:width,'cm'), ':',
																   substring-before(@svg:height,'cm'), ':', 
																   $x2, ':', 
                                   $y2, ':', 
																   $angle)"/>
          </xsl:attribute>
          
          <xsl:attribute name ="y">
                 <xsl:value-of select ="concat('draw-transform:Y:',substring-before(@svg:width,'cm'), ':',
																   substring-before(@svg:height,'cm'), ':', 
																   $x2, ':', 
                                   $y2, ':', 
																   $angle)"/>
          </xsl:attribute>
        </xsl:if>
        <xsl:if test="not(@draw:transform)">
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
        </xsl:if>
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
		<xsl:param name ="fileName" />
    <xsl:param name ="connectorType" />
    <p:cxnSp>
      <p:nvCxnSpPr>
        <p:cNvPr>
          <xsl:attribute name ="id">
            <xsl:value-of select ="position()"/>
          </xsl:attribute>
          <xsl:attribute name ="name">
            <xsl:value-of select ="concat($connectorType, position())" />
          </xsl:attribute>
          <!-- ADDED BY LOHITH - HYPER LINKS FOR SHAPES-->
          <xsl:variable name="PostionCount">
            <xsl:value-of select="position()"/>
          </xsl:variable>
          <xsl:if test="office:event-listeners">
            <xsl:for-each select ="office:event-listeners/presentation:event-listener">
              <xsl:if test="@script:event-name[contains(.,'dom:click')]">
                <a:hlinkClick>
                  <xsl:choose>
                    <!-- Go to previous slide-->
                    <xsl:when test="@presentation:action[ contains(.,'previous-page')]">
                      <xsl:attribute name="action">
                        <xsl:value-of select="'ppaction://hlinkshowjump?jump=previousslide'"/>
                      </xsl:attribute>
                    </xsl:when>
                    <!-- Go to Next slide-->
                    <xsl:when test="@presentation:action[ contains(.,'next-page')]">
                      <xsl:attribute name="action">
                        <xsl:value-of select="'ppaction://hlinkshowjump?jump=nextslide'"/>
                      </xsl:attribute>
                    </xsl:when>
                    <!-- Go to First slide-->
                    <xsl:when test="@presentation:action[ contains(.,'first-page')]">
                      <xsl:attribute name="action">
                        <xsl:value-of select="'ppaction://hlinkshowjump?jump=firstslide'"/>
                      </xsl:attribute>
                    </xsl:when>
                    <!-- Go to Last slide-->
                    <xsl:when test="@presentation:action[ contains(.,'last-page')]">
                      <xsl:attribute name="action">
                        <xsl:value-of select="'ppaction://hlinkshowjump?jump=lastslide'"/>
                      </xsl:attribute>
                    </xsl:when>
                    <!-- End Presentation -->
                    <xsl:when test="@presentation:action[ contains(.,'stop')]">
                      <xsl:attribute name="action">
                        <xsl:value-of select="'ppaction://hlinkshowjump?jump=endshow'"/>
                      </xsl:attribute>
                    </xsl:when>
                    <!-- Run program-->
                    <xsl:when test="@xlink:href and @presentation:action[ contains(.,'execute')]">
                      <xsl:attribute name="action">
                        <xsl:value-of select="'ppaction://program'"/>
                      </xsl:attribute>
                    </xsl:when>
                    <!-- Go to slide -->
                    <xsl:when test="@xlink:href[ contains(.,'#page')]">
                      <xsl:attribute name="action">
                        <xsl:value-of select="'ppaction://hlinksldjump'"/>
                      </xsl:attribute>
                    </xsl:when>
                    <!-- Go to document -->
                    <xsl:when test="@xlink:href and @presentation:action[ contains(.,'show')] ">
                      <xsl:if test="not(@xlink:href[ contains(.,'#page')])">
                        <xsl:attribute name="action">
                          <xsl:value-of select="'ppaction://hlinkfile'"/>
                        </xsl:attribute>
                      </xsl:if>
                    </xsl:when>
                  </xsl:choose>
                  <!-- set value for attribute r:id-->
                  <xsl:choose>
                    <!-- For jump to next,previous,first,last-->
                    <xsl:when test="@presentation:action[ contains(.,'page') or contains(.,'stop')]">
                      <xsl:attribute name="r:id">
                        <xsl:value-of select="''"/>
                      </xsl:attribute>
                    </xsl:when>
                    <!-- For Run program & got to slide-->
                    <xsl:when test="@xlink:href[contains(.,'#page')]">
                      <xsl:attribute name="r:id">
                        <xsl:value-of select="concat('LineFileId',$PostionCount)"/>
                      </xsl:attribute>
                    </xsl:when>
                    <!-- For Go to document -->
                    <xsl:when test="@xlink:href and @presentation:action[ contains(.,'show')] ">
                      <xsl:if test="not(@xlink:href[ contains(.,'#page')])">
                        <xsl:attribute name="r:id">
                          <xsl:value-of select="concat('LineFileId',$PostionCount)"/>
                        </xsl:attribute>
                      </xsl:if>
                    </xsl:when>
                    <!-- For Go to Run Programs -->
                    <xsl:when test="@xlink:href and @presentation:action[ contains(.,'execute')]">
                      <xsl:attribute name="r:id">
                        <xsl:value-of select="concat('LineFileId',$PostionCount)"/>
                      </xsl:attribute>
                    </xsl:when>
                  </xsl:choose>
                  <!-- Play Sound -->
                  <xsl:if test="@presentation:action[ contains(.,'sound')]">
                    <a:snd>
                      <xsl:variable name="varMediaFilePath">
                        <xsl:if test="presentation:sound/@xlink:href [ contains(.,'../')]">
                          <xsl:value-of select="presentation:sound/@xlink:href" />
                        </xsl:if>
                        <xsl:if test="not(presentation:sound/@xlink:href [ contains(.,'../')])">
                          <xsl:value-of select="substring-after(presentation:sound/@xlink:href,'/')" />
                        </xsl:if>
                      </xsl:variable>
                      <xsl:variable name="varFileRelId">
                        <xsl:value-of select="translate(translate(translate(translate(translate($varMediaFilePath,'/','_'),'..','_'),'.','_'),':','_'),'%20D','_')"/>
                      </xsl:variable>
                      <xsl:attribute name="r:embed">
                        <xsl:value-of select="$varFileRelId"/>
                      </xsl:attribute>
                      <xsl:attribute name="name">
                        <xsl:value-of select="concat('SoundFileForShapeLine',$PostionCount)"/>
                      </xsl:attribute>
                      <xsl:attribute name="builtIn">
                        <xsl:value-of select='1'/>
                      </xsl:attribute>
                      <pzip:extract pzip:source="{$varMediaFilePath}" pzip:target="{concat('ppt/media/',$varFileRelId,'.wav')}" />
                    </a:snd>
                  </xsl:if>
                </a:hlinkClick>
              </xsl:if>
            </xsl:for-each>
          </xsl:if>
          <!-- ADDED BY LOHITH - HYPER LINKS FOR SHAPES-->
        </p:cNvPr>
        <p:cNvCxnSpPr/>
        <p:nvPr/>
      </p:nvCxnSpPr>
      <p:spPr>

        <a:xfrm>
          <xsl:attribute name ="rot">
            <xsl:value-of select ="concat('cxnSp:rot:',substring-before(@svg:x1,'cm'), ':',
																   substring-before(@svg:x2,'cm'), ':', 
																   substring-before(@svg:y1,'cm'), ':', 
																   substring-before(@svg:y2,'cm'))"/>
          </xsl:attribute>
          <xsl:attribute name ="flipH">
            <xsl:value-of select ="concat('cxnSp:flipH:',substring-before(@svg:x1,'cm'), ':',
																   substring-before(@svg:x2,'cm'), ':', 
																   substring-before(@svg:y1,'cm'), ':', 
																   substring-before(@svg:y2,'cm'))"/>
          </xsl:attribute>
          <xsl:attribute name ="flipV">
            <xsl:value-of select ="concat('cxnSp:flipV:',substring-before(@svg:x1,'cm'), ':',
																   substring-before(@svg:x2,'cm'), ':', 
																   substring-before(@svg:y1,'cm'), ':', 
																   substring-before(@svg:y2,'cm'))"/>
          </xsl:attribute>
          <a:off >
            <xsl:attribute name ="x">

              <xsl:value-of select ="concat('cxnSp:x:',substring-before(@svg:x1,'cm'), ':',
																   substring-before(@svg:x2,'cm'), ':', 
																   substring-before(@svg:y1,'cm'), ':', 
																   substring-before(@svg:y2,'cm'))"/>
            </xsl:attribute>
            <xsl:attribute name ="y">
              <xsl:value-of select ="concat('cxnSp:y:',substring-before(@svg:x1,'cm'), ':',
																   substring-before(@svg:x2,'cm'), ':', 
																   substring-before(@svg:y1,'cm'), ':', 
																   substring-before(@svg:y2,'cm'))"/>
            </xsl:attribute>
          </a:off>
          <a:ext>
            <xsl:attribute name ="cx">
              <xsl:value-of select ="concat('cxnSp:cx:',substring-before(@svg:x1,'cm'), ':',
																   substring-before(@svg:x2,'cm'), ':', 
																   substring-before(@svg:y1,'cm'), ':', 
																   substring-before(@svg:y2,'cm'))"/>
            </xsl:attribute>
            <xsl:attribute name ="cy">
              <xsl:value-of select ="concat('cxnSp:cy:',substring-before(@svg:x1,'cm'), ':',
																   substring-before(@svg:x2,'cm'), ':', 
																   substring-before(@svg:y1,'cm'), ':', 
																   substring-before(@svg:y2,'cm'))"/>
            </xsl:attribute>

          </a:ext>
        </a:xfrm>

        <xsl:call-template name ="getPresetGeom">
          <xsl:with-param name ="prstGeom" select ="$connectorType" />
        </xsl:call-template>

        <xsl:call-template name ="getGraphicProperties">
					<xsl:with-param name ="fileName" select ="$fileName" />
					<xsl:with-param name ="gr" >
						<xsl:if test ="@draw:style-name">
							<xsl:value-of select ="@draw:style-name"/>
						</xsl:if>
						<xsl:if test ="@presentation:style-name">
							<xsl:value-of select ="@presentation:style-name"/>
						</xsl:if >
					</xsl:with-param >
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
		<xsl:param name ="fileName" />
    <xsl:param name ="gr" />

		<xsl:for-each select ="document($fileName)//office:automatic-styles/style:style[@style:name=$gr]/style:graphic-properties">
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

      <!-- Oval -->
      <xsl:when test ="contains($prstGeom, 'Oval')">
        <a:prstGeom prst="ellipse">
          <a:avLst/>
        </a:prstGeom>
      </xsl:when>
      <!-- Isosceles Triangle -->
      <xsl:when test ="contains($prstGeom, 'Isosceles Triangle')">
        <a:prstGeom prst="triangle">
          <a:avLst/>
        </a:prstGeom>
      </xsl:when>
		<!-- Left-Right Arrow (Added by Mathi as on 19/7/2007)-->
		<xsl:when test ="contains($prstGeom, 'Left-Right Arrow')">
			<a:prstGeom prst="leftRightArrow">
				<a:avLst/>
			</a:prstGeom>
		</xsl:when>
      <!-- Right Arrow (Added by Mathi as on 4/7/2007) -->
      <xsl:when test ="contains($prstGeom, 'Right Arrow')">
        <a:prstGeom prst="rightArrow">
          <a:avLst/>
        </a:prstGeom>
      </xsl:when>
      <!-- Left Arrow (Added by Mathi as on 5/7/2007)-->
      <xsl:when test ="contains($prstGeom, 'Left Arrow')">
        <a:prstGeom prst="leftArrow">
				<a:avLst/>
			</a:prstGeom>
		</xsl:when>
		<!--Up Arrow (Added by Mathi as on 19/7/2007)-->
		<xsl:when test ="contains($prstGeom, 'Up Arrow')">
			<a:prstGeom prst="upArrow">
				<a:avLst/>
			</a:prstGeom>
		</xsl:when>
		<!-- Up-Down Arrow (Added by Mathi as on 19/7/2007)-->
		<xsl:when test ="contains($prstGeom, 'Up-Down Arrow')">
			<a:prstGeom prst="upDownArrow">
				<a:avLst/>
			</a:prstGeom>
		</xsl:when>
		<!-- Down Arrow (Added by Mathi as on 19/7/2007)-->
		<xsl:when test ="contains($prstGeom, 'Down Arrow')">
			<a:prstGeom prst="downArrow">
          <a:avLst/>
        </a:prstGeom>
      </xsl:when>
      <!-- Right Triangle -->
      <xsl:when test ="contains($prstGeom, 'Right Triangle')">
        <a:prstGeom prst="rtTriangle">
          <a:avLst/>
        </a:prstGeom>
      </xsl:when>
      <!-- Parallelogram -->
      <xsl:when test ="contains($prstGeom, 'Parallelogram')">
        <a:prstGeom prst="parallelogram">
          <a:avLst/>
        </a:prstGeom>
      </xsl:when>
	  <!-- Trapezoid (Added by A.Mathi as on 24/07/2007) -->
      <xsl:when test ="contains($prstGeom, 'Trapezoid')">
        <a:prstGeom prst="trapezoid">
          <a:avLst/>
        </a:prstGeom>
      </xsl:when>
      <!-- Diamond -->
      <xsl:when test ="contains($prstGeom, 'Diamond')">
        <a:prstGeom prst="diamond">
          <a:avLst/>
        </a:prstGeom>
      </xsl:when>
      <!-- Regular Pentagon -->
      <xsl:when test ="contains($prstGeom, 'Regular Pentagon')">
        <a:prstGeom prst="pentagon">
          <a:avLst/>
        </a:prstGeom>
      </xsl:when>
      <!-- Hexagon -->
      <xsl:when test ="contains($prstGeom, 'Hexagon')">
        <a:prstGeom prst="hexagon">
          <a:avLst/>
        </a:prstGeom>
      </xsl:when>
      <!-- Octagon -->
      <xsl:when test ="contains($prstGeom, 'Octagon')">
        <a:prstGeom prst="octagon">
          <a:avLst>
            <a:gd name="adj" fmla="val 32414"/>
          </a:avLst>
        </a:prstGeom>
      </xsl:when>
      <!-- Cube -->
      <xsl:when test ="contains($prstGeom, 'Cube')">
        <a:prstGeom prst="cube">
          <a:avLst/>
        </a:prstGeom>
      </xsl:when>
      <!-- Can -->
      <xsl:when test ="contains($prstGeom, 'Can')">
        <a:prstGeom prst="can">
          <a:avLst/>
        </a:prstGeom>
      </xsl:when>
	  <!-- Cross (Added by Mathi on 19/7/2007)-->
		<xsl:when test ="contains($prstGeom, 'plus')">
			<a:prstGeom prst="plus">
				<a:avLst/>
			</a:prstGeom>
		</xsl:when>
	  <!-- "No" Symbol (Added by A.Mathi as on 19/07/2007)-->
		<xsl:when test ="contains($prstGeom, 'noSmoking')">
			<a:prstGeom prst="noSmoking">
				<a:avLst/>
			</a:prstGeom>
		</xsl:when>
	  <!--  Folded Corner (Added by A.Mathi as on 19/07/2007) -->
		<xsl:when test ="contains($prstGeom, 'foldedCorner')">
			<a:prstGeom prst="foldedCorner">
				<a:avLst/>
			</a:prstGeom>
		</xsl:when>
	  <!--  Lightning Bolt (Added by A.Mathi as on 20/07/2007) -->
		<xsl:when test ="contains($prstGeom, 'lightningBolt')">
			<a:prstGeom prst="lightningBolt">
				<a:avLst/>
			</a:prstGeom>
		</xsl:when>
	  <!--  Explosion 1 (Added by A.Mathi as on 20/07/2007) -->
		<xsl:when test ="contains($prstGeom, 'irregularSeal1')">
			<a:prstGeom prst="irregularSeal1">
				<a:avLst/>
			</a:prstGeom>
		</xsl:when>
	  <!-- Left Bracket (Added by A.Mathi as on 20/07/2007) -->
		<xsl:when test ="contains($prstGeom, 'Left Bracket')">
			<a:prstGeom prst="leftBracket">
				<a:avLst/>
			</a:prstGeom>
		</xsl:when>
	 <!-- Right Bracket (Added by A.Mathi as on 20/07/2007) -->
		<xsl:when test ="contains($prstGeom, 'Right Bracket')">
			<a:prstGeom prst="rightBracket">
				<a:avLst/>
			</a:prstGeom>
		</xsl:when>
	 <!-- Left Brace (Added by A.Mathi as on 20/07/2007) -->
		<xsl:when test ="contains($prstGeom, 'Left Brace')">
			<a:prstGeom prst="leftBrace">
				<a:avLst/>
			</a:prstGeom>
		</xsl:when>
	 <!-- Right Brace (Added by A.Mathi as on 23/07/2007) -->
		<xsl:when test ="contains($prstGeom, 'Right Brace')">
			<a:prstGeom prst="rightBrace">
				<a:avLst/>
			</a:prstGeom>
		</xsl:when>
	 <!-- Rectangular Callout (Added by A.Mathi as on 23/07/2007) -->
		<xsl:when test ="contains($prstGeom, 'Rectangular Callout')">
			<a:prstGeom prst="wedgeRectCallout">
				<a:avLst/>
			</a:prstGeom>
		</xsl:when>
	 <!-- Rounded Rectangular Callout (Added by A.Mathi as on 23/07/2007) -->
		<xsl:when test ="contains($prstGeom, 'wedgeRoundRectCallout')">
			<a:prstGeom prst="wedgeRoundRectCallout">
				<a:avLst/>
			</a:prstGeom>
		</xsl:when>
	 <!-- Oval Callout (Added by A.Mathi as on 23/07/2007) -->
		<xsl:when test ="contains($prstGeom, 'wedgeEllipseCallout')">
			<a:prstGeom prst="wedgeEllipseCallout">
				<a:avLst/>
			</a:prstGeom>
		</xsl:when>
	 <!-- Cloud Callout (Added by A.Mathi as on 23/07/2007) -->
		<xsl:when test ="contains($prstGeom, 'Cloud Callout')">
			<a:prstGeom prst="cloudCallout">
				<a:avLst/>
			</a:prstGeom>
		</xsl:when>
	 <!-- Bent Arrow (Added by A.Mathi as on 23/07/2007) -->
		<xsl:when test ="contains($prstGeom, 'bentArrow')">
			<a:prstGeom prst="bentArrow">
				<a:avLst/>
			</a:prstGeom>
		</xsl:when>
	<!-- U-Turn Arrow (Added by A.Mathi as on 23/07/2007) -->
		<xsl:when test ="contains($prstGeom, 'U-Turn Arrow')">
			<a:prstGeom prst="uturnArrow">
				<a:avLst/>
			</a:prstGeom>
		</xsl:when>
      <!-- Rounded Rectangle -->
      <xsl:when test ="contains($prstGeom, 'Rounded Rectangle')">
        <a:prstGeom prst="roundRect">
          <a:avLst/>
        </a:prstGeom>
      </xsl:when>
      <!-- Snip Single Corner Rectangle -->
      <xsl:when test ="contains($prstGeom, 'Snip Single Corner Rectangle')">
        <a:prstGeom prst="snip1Rect">
          <a:avLst/>
        </a:prstGeom>
      </xsl:when>
      <!-- Snip Same Side Corner -->
      <xsl:when test ="contains($prstGeom, 'Snip Same Side Corner Rectangle')">
        <a:prstGeom prst="snip2SameRect">
          <a:avLst/>
        </a:prstGeom>
      </xsl:when>
		<!-- Circular Arrow -->
		<xsl:when test ="contains($prstGeom, 'circular-arrow')">
			<a:prstGeom prst="circularArrow">
				<a:avLst/>
			</a:prstGeom>
		</xsl:when>
		<!-- leftUp Arrow -->
		<xsl:when test ="contains($prstGeom, 'leftUpArrow')">
			<a:prstGeom prst="leftUpArrow">
				<a:avLst/>
			</a:prstGeom>
		</xsl:when>
		<!-- bentUp Arrow -->
		<xsl:when test ="contains($prstGeom, 'bentUpArrow')">
			<a:prstGeom prst="bentUpArrow">
          <a:avLst/>
        </a:prstGeom>
      </xsl:when>
      <!-- Snip Diagonal Corner -->
      <xsl:when test ="contains($prstGeom, 'Snip Diagonal Corner Rectangle')">
        <a:prstGeom prst="snip2DiagRect">
          <a:avLst/>
        </a:prstGeom>
      </xsl:when>
      <!-- and Round Single Corner -->
      <xsl:when test ="contains($prstGeom, 'Snip and Round Single Corner Rectangle')">
        <a:prstGeom prst="snipRoundRect">
          <a:avLst/>
        </a:prstGeom>
      </xsl:when>
      <!-- Round Single Corner -->
      <xsl:when test ="contains($prstGeom, 'Round Single Corner Rectangle')">
        <a:prstGeom prst="round1Rect">
          <a:avLst/>
        </a:prstGeom>
      </xsl:when>
      <!-- Round Same Side Corner -->
      <xsl:when test ="contains($prstGeom, 'Round Same Side Corner Rectangle')">
        <a:prstGeom prst="round2SameRect">
          <a:avLst/>
        </a:prstGeom>
      </xsl:when>
      <!-- Round Diagonal Corner -->
      <xsl:when test ="contains($prstGeom, 'Round Diagonal Corner Rectangle')">
        <a:prstGeom prst="round2DiagRect">
          <a:avLst/>
        </a:prstGeom>
      </xsl:when>
      <!-- Rectangles -->
      <xsl:when test ="(contains($prstGeom, 'Rectangle')) or (contains($prstGeom, 'TextBox'))">
        <a:prstGeom prst="rect">
          <a:avLst/>
        </a:prstGeom>
      </xsl:when>

      <!-- Connectors-->
      <!-- Straight line-->
      <xsl:when test ="contains($prstGeom, 'Straight Connector')">
        <a:prstGeom prst="line">
          <a:avLst/>
        </a:prstGeom>
      </xsl:when>
      <!-- Straight Arrow Connector -->
      <xsl:when test ="contains($prstGeom, 'Straight Arrow Connector')">
        <a:prstGeom prst="straightConnector1">
          <a:avLst/>
        </a:prstGeom>
      </xsl:when>
      <!-- Elbow Connector -->
      <xsl:when test ="contains($prstGeom, 'Elbow Connector')">
        <a:prstGeom prst="bentConnector3">
          <a:avLst>
            <a:gd name="adj1" fmla="val 50000"/>
          </a:avLst>
        </a:prstGeom>

      </xsl:when>
      <!-- Curved Connector -->
      <xsl:when test ="contains($prstGeom, 'Curved Connector')">
        <a:prstGeom prst="curvedConnector3">
          <a:avLst>
            <a:gd name="adj1" fmla="val 60417"/>
          </a:avLst>
        </a:prstGeom>
      </xsl:when>

      <!-- Flow Chart: Process -->
      <xsl:when test ="contains($prstGeom, 'Flowchart: Process')">
        <a:prstGeom prst="flowChartProcess">
          <a:avLst/>
        </a:prstGeom>
      </xsl:when>
      <!-- Flow Chart: Alternate Process -->
      <xsl:when test ="contains($prstGeom, 'Flowchart: Alternate Process')">
        <a:prstGeom prst="flowChartAlternateProcess">
          <a:avLst/>
        </a:prstGeom>
      </xsl:when>
      <!-- Flow Chart: Decision -->
      <xsl:when test ="contains($prstGeom, 'Flowchart: Decision')">
        <a:prstGeom prst="flowChartDecision">
          <a:avLst/>
        </a:prstGeom>
      </xsl:when>
      <!-- Flow Chart: Data -->
      <xsl:when test ="contains($prstGeom, 'Flowchart: Data')">
        <a:prstGeom prst="flowChartInputOutput">
          <a:avLst/>
        </a:prstGeom>
      </xsl:when>
      <!-- Flow Chart: Predefined-process -->
      <xsl:when test ="contains($prstGeom, 'Flowchart: Predefined Process')">
        <a:prstGeom prst="flowChartPredefinedProcess">
          <a:avLst/>
        </a:prstGeom>
      </xsl:when>
      <!-- Flow Chart: Internal-storage -->
      <xsl:when test ="contains($prstGeom, 'Flowchart: Internal Storage')">
        <a:prstGeom prst="flowChartInternalStorage">
          <a:avLst/>
        </a:prstGeom>
      </xsl:when>
      <!-- Flow Chart: Document -->
      <xsl:when test ="contains($prstGeom, 'Flowchart: Document')">
        <a:prstGeom prst="flowChartDocument">
          <a:avLst/>
        </a:prstGeom>
      </xsl:when>
      <!-- Flow Chart: Multidocument -->
      <xsl:when test ="contains($prstGeom, 'Flowchart: Multi document')">
        <a:prstGeom prst="flowChartMultidocument">
          <a:avLst/>
        </a:prstGeom>
      </xsl:when>
      <!-- Flow Chart: Terminator -->
      <xsl:when test ="contains($prstGeom, 'Flowchart: Terminator')">
        <a:prstGeom prst="flowChartTerminator">
          <a:avLst/>
        </a:prstGeom>
      </xsl:when>
      <!-- Flow Chart: Preparation -->
      <xsl:when test ="contains($prstGeom, 'Flowchart: Preparation')">
        <a:prstGeom prst="flowChartPreparation">
          <a:avLst/>
        </a:prstGeom>
      </xsl:when>
      <!-- Flow Chart: Manual-input -->
      <xsl:when test ="contains($prstGeom, 'Flowchart: Manual Input')">
        <a:prstGeom prst="flowChartManualInput">
          <a:avLst/>
        </a:prstGeom>
      </xsl:when>
      <!-- Flow Chart: Manual-operation -->
      <xsl:when test ="contains($prstGeom, 'Flowchart: Manual Operation')">
        <a:prstGeom prst="flowChartManualOperation">
          <a:avLst/>
        </a:prstGeom>
      </xsl:when>
      <!-- Flow Chart: Connector -->
      <xsl:when test ="contains($prstGeom, 'Flowchart: Connector')">
        <a:prstGeom prst="flowChartConnector">
          <a:avLst/>
        </a:prstGeom>
      </xsl:when>
      <!-- Flow Chart: Off-page-connector -->
      <xsl:when test ="contains($prstGeom, 'Flowchart: Off-page Connector')">
        <a:prstGeom prst="flowChartOffpageConnector">
          <a:avLst/>
        </a:prstGeom>
      </xsl:when>
      <!-- Flow Chart: Card -->
      <xsl:when test ="contains($prstGeom, 'Flowchart: Card')">
        <a:prstGeom prst="flowChartPunchedCard">
          <a:avLst/>
        </a:prstGeom>
      </xsl:when>
      <!-- Flow Chart: Punched-tape -->
      <xsl:when test ="contains($prstGeom, 'Flowchart: Punched Tape')">
        <a:prstGeom prst="flowChartPunchedTape">
          <a:avLst/>
        </a:prstGeom>
      </xsl:when>
      <!-- Flow Chart: Summing-junction -->
      <xsl:when test ="contains($prstGeom, 'Flowchart: Summing Junction')">
        <a:prstGeom prst="flowChartSummingJunction">
          <a:avLst/>
        </a:prstGeom>
      </xsl:when>
      <!-- Flow Chart: Or -->
      <xsl:when test ="contains($prstGeom, 'Flowchart: Or')">
        <a:prstGeom prst="flowChartOr">
          <a:avLst/>
        </a:prstGeom>
      </xsl:when>
      <!-- Flow Chart: Collate -->
      <xsl:when test ="contains($prstGeom, 'Flowchart: Collate')">
        <a:prstGeom prst="flowChartCollate">
          <a:avLst/>
        </a:prstGeom>
      </xsl:when>
      <!-- Flow Chart: Sort -->
      <xsl:when test ="contains($prstGeom, 'Flowchart: Sort')">
        <a:prstGeom prst="flowChartSort">
          <a:avLst/>
        </a:prstGeom>
      </xsl:when>
      <!-- Flow Chart: Extract -->
      <xsl:when test ="contains($prstGeom, 'Flowchart: Extract')">
        <a:prstGeom prst="flowChartExtract">
          <a:avLst/>
        </a:prstGeom>
      </xsl:when>
      <!-- Flow Chart: Merge -->
      <xsl:when test ="contains($prstGeom, 'Flowchart: Merge')">
        <a:prstGeom prst="flowChartMerge">
          <a:avLst/>
        </a:prstGeom>
      </xsl:when>
      <!-- Flow Chart: Stored-data -->
      <xsl:when test ="contains($prstGeom, 'Flowchart: Stored Data')">
        <a:prstGeom prst="flowChartOnlineStorage">
          <a:avLst/>
        </a:prstGeom>
      </xsl:when>
      <!-- Flow Chart: Delay-->
      <xsl:when test ="contains($prstGeom, 'Flowchart: Delay')">
        <a:prstGeom prst="flowChartDelay">
          <a:avLst/>
        </a:prstGeom>
      </xsl:when>
      <!-- Flow Chart: Sequential-access -->
      <xsl:when test ="contains($prstGeom, 'Flowchart: Sequential Access Storage')">
        <a:prstGeom prst="flowChartMagneticTape">
          <a:avLst/>
        </a:prstGeom>
      </xsl:when>
      <!-- Flow Chart: Direct-access-storage -->
      <xsl:when test ="contains($prstGeom, 'Flowchart: Direct Access Storage')">
        <a:prstGeom prst="flowChartMagneticDrum">
          <a:avLst/>
        </a:prstGeom>
      </xsl:when>
      <!-- Flow Chart: Magnetic-disk -->
      <xsl:when test ="contains($prstGeom, 'Flowchart: Magnetic Disk')">
        <a:prstGeom prst="flowChartMagneticDisk">
          <a:avLst/>
        </a:prstGeom>
      </xsl:when>
      <!-- Flow Chart: Display -->
      <xsl:when test ="contains($prstGeom, 'Flowchart: Display')">
        <a:prstGeom prst="flowChartDisplay">
          <a:avLst/>
        </a:prstGeom>
      </xsl:when>
    </xsl:choose>
  </xsl:template>
  <!-- Get fill details-->
  <xsl:template name ="getFillDetails">
    <xsl:param name ="parentStyle" />
	  <xsl:variable name ="opacity" >
		  <xsl:value-of select ="@draw:opacity"/>
	  </xsl:variable>
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
                <xsl:with-param name ="opacity" select ="$opacity" />
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
            <!-- Code for the Bug 1746356 -->
            <xsl:if test ="not(@svg:stroke-width)">
              <xsl:attribute name ="w">
                <xsl:value-of select ="'0'"/>
              </xsl:attribute>
            </xsl:if >
            <!-- Code for the Bug 1746356 -->
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
		<xsl:variable name ="lineOpacity"	>
			<xsl:value-of select ="@svg:stroke-opacity"/>
		</xsl:variable>
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
                  <xsl:with-param name ="opacity" select ="$lineOpacity" />
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
  <!-- Get color code -->
  <xsl:template name ="getFillColor">
    <xsl:param name ="fill-color" />
    <xsl:param name ="opacity" />
    <xsl:if test ="$fill-color != ''">
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
    </xsl:if>
  </xsl:template>
  <!-- Get shade for fill reference-->
  <xsl:template name ="getShade">
		<xsl:param name ="fileName" />
    <xsl:param name ="gr" />
		<xsl:if test ="document($fileName)//office:automatic-styles/style:style[@style:name=$gr]/style:graphic-properties/@draw:shadow-opacity">
      <xsl:variable name ="shade">
				<xsl:value-of select ="document($fileName)//office:automatic-styles/style:style[@style:name=$gr]/style:graphic-properties/@draw:shadow-opacity"/>
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
      <xsl:when test ="($size &lt;= $sm-sm)">
        <xsl:attribute name ="w">
          <xsl:value-of select ="'sm'"/>
        </xsl:attribute>
        <xsl:attribute name ="len">
          <xsl:value-of select ="'sm'"/>
        </xsl:attribute>
      </xsl:when>
      <xsl:when test ="($size &gt; $sm-sm) and ($size &lt;= $sm-med)">
        <xsl:attribute name ="w">
          <xsl:value-of select ="'sm'"/>
        </xsl:attribute>
        <xsl:attribute name ="len">
          <xsl:value-of select ="'med'"/>
        </xsl:attribute>
      </xsl:when>
      <xsl:when test ="($size &gt; $sm-med) and ($size &lt;= $sm-lg)">
        <xsl:attribute name ="w">
          <xsl:value-of select ="'sm'"/>
        </xsl:attribute>
        <xsl:attribute name ="len">
          <xsl:value-of select ="'lg'"/>
        </xsl:attribute>
      </xsl:when>
      <xsl:when test ="($size &gt; $sm-lg) and ($size &lt;= $med-sm)">
        <xsl:attribute name ="w">
          <xsl:value-of select ="'med'"/>
        </xsl:attribute>
        <xsl:attribute name ="len">
          <xsl:value-of select ="'sm'"/>
        </xsl:attribute>
      </xsl:when>
      <xsl:when test ="($size &gt; $med-med) and ($size &lt;= $med-lg)">
        <xsl:attribute name ="w">
          <xsl:value-of select ="'med'"/>
        </xsl:attribute>
        <xsl:attribute name ="len">
          <xsl:value-of select ="'lg'"/>
        </xsl:attribute>
      </xsl:when>
      <xsl:when test ="($size &gt; $med-lg) and ($size &lt;= $lg-sm)">
        <xsl:attribute name ="w">
          <xsl:value-of select ="'lg'"/>
        </xsl:attribute>
        <xsl:attribute name ="len">
          <xsl:value-of select ="'sm'"/>
        </xsl:attribute>
      </xsl:when>
      <xsl:when test ="($size &gt; $lg-sm) and ($size &lt;= $lg-med)">
        <xsl:attribute name ="w">
          <xsl:value-of select ="'lg'"/>
        </xsl:attribute>
        <xsl:attribute name ="len">
          <xsl:value-of select ="'med'"/>
        </xsl:attribute>
      </xsl:when>
      <xsl:when test ="($size &gt; $lg-med) and ($size &lt;= $lg-lg)">
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
    <xsl:if test ="document('styles.xml')/office:document-styles/office:styles/draw:stroke-dash[@draw:name=$stroke-dash]">
      <xsl:variable name ="dots1" select="document('styles.xml')/office:document-styles/office:styles/draw:stroke-dash[@draw:name=$stroke-dash]/@draw:dots1"/>
      <xsl:variable name ="dots1-length" select ="substring-before(document('styles.xml')/office:document-styles/office:styles/draw:stroke-dash[@draw:name=$stroke-dash]/@draw:dots1-length, 'cm')" />
      <xsl:variable name ="dots2" select="document('styles.xml')/office:document-styles/office:styles/draw:stroke-dash[@draw:name=$stroke-dash]/@draw:dots2"/>
      <xsl:variable name ="dots2-length" select ="substring-before(document('styles.xml')/office:document-styles/office:styles/draw:stroke-dash[@draw:name=$stroke-dash]/@draw:dots2-length, 'cm')"/>
      <xsl:variable name ="distance" select ="substring-before(document('styles.xml')/office:document-styles/office:styles/draw:stroke-dash[@draw:name=$stroke-dash]/@draw:distance, 'cm')" />

      <xsl:choose>
        <xsl:when test ="($dots1=1) and ($dots2=1)">
          <xsl:choose>
            <xsl:when test ="($dots1-length &lt;= $dot) and ($dots2-length &lt;= $dot)">
              <xsl:value-of select ="'sysDot'" />
            </xsl:when>
            <xsl:when test ="(($dots1-length &lt;= $dot) and ($dots2-length &lt;= $dash)) or
										 (($dots1-length &lt;= $dash) and ($dots2-length &lt;= $dot)) ">
              <xsl:value-of select ="'dashDot'" />
            </xsl:when>
            <xsl:when test ="($dots1-length &lt;= $dash) and ($dots2-length &lt;= $dash)">
              <xsl:value-of select ="'dash'" />
            </xsl:when>
            <xsl:when test ="(($dots1-length &lt;= $dot) and (($dots2-length &gt;= $dash) and ($dots2-length &lt;= $longDash))) or
										 (($dots2-length &lt;= $dot) and (($dots1-length &gt;= $dash) and ($dots1-length &lt;= $longDash))) ">
              <xsl:value-of select ="'lgDashDot'" />
            </xsl:when>
            <xsl:when test ="(($dots1-length &gt;= $dash) and ($dots1-length &lt;= $longDash)) or 
										 (($dots2-length &gt;= $dash) and ($dots2-length &lt;= $longDash))">
              <xsl:value-of select ="'lgDash'" />
            </xsl:when>
            <xsl:otherwise>
              <xsl:value-of select ="'sysDash'" />
            </xsl:otherwise>
          </xsl:choose>
        </xsl:when>
        <xsl:when test ="(($dots1=2) and ($dots2=1)) or (($dots1=1) and ($dots2=2))">
          <!--<xsl:if test ="(($dots1-length &lt;= $dot) and ($dots2-length &gt;= $dash) and ($dots2-length &lt;= $longDash))">-->
          <xsl:value-of select ="'lgDashDotDot'" />
          <!--</xsl:if>-->
        </xsl:when>
        <!--<xsl:when test ="($dots1=1) and ($dots2=2)">
					-->
        <!--<xsl:if test ="(($dots2-length &lt;= $dot) and ($dots1-length &gt;= $dash) and ($dots1-length &lt;= $longDash))">-->
        <!--
						<xsl:value-of select ="'lgDashDotDot'" />
					-->
        <!--</xsl:if>-->
        <!--
				</xsl:when>-->
        <xsl:when test ="(($dots1 &gt;= 1) and not($dots2))">
          <xsl:choose>
            <xsl:when test ="($dots1-length &lt;= $dot)">
              <xsl:value-of select ="'sysDash'" />
            </xsl:when>
            <xsl:when test ="($dots1-length &gt;= $dot) and ($dots1-length &lt;= $dash)">
              <xsl:value-of select ="'dash'" />
            </xsl:when>
            <xsl:when test ="($dots1-length &gt;= $dash) and ($dots1-length &lt;= $longDash)">
              <xsl:value-of select ="'lgDash'" />
            </xsl:when>
            <xsl:otherwise>
              <xsl:value-of select ="'sysDash'" />
            </xsl:otherwise>
          </xsl:choose>
        </xsl:when>
        <xsl:when test ="(($dots2 &gt;= 1) and not($dots1))">
          <xsl:choose>
            <xsl:when test ="($dots2-length &lt;= $dot)">
              <xsl:value-of select ="'sysDash'" />
            </xsl:when>
            <xsl:when test ="($dots2-length &gt;= $dot) and ($dots2-length &lt;= $dash)">
              <xsl:value-of select ="'dash'" />
            </xsl:when>
            <xsl:when test ="($dots2-length &gt;= $dash) and ($dots2-length &lt;= $longDash)">
              <xsl:value-of select ="'lgDash'" />
            </xsl:when>
            <xsl:otherwise>
              <xsl:value-of select ="'sysDash'" />
            </xsl:otherwise>
          </xsl:choose>
        </xsl:when>
        <xsl:otherwise>
          <xsl:value-of select ="'sysDash'" />
        </xsl:otherwise>
      </xsl:choose>
    </xsl:if >
  </xsl:template>
  <!-- Get text layout for shape-->
  <xsl:template name ="getParagraphProperties">
		<xsl:param name ="fileName" />
    <xsl:param name ="gr" />
    <xsl:param name ="customShape" />

    <a:bodyPr>
      <!-- Text direction-->
			<xsl:for-each select ="document($fileName)//office:automatic-styles/style:style[@style:name=$gr]/style:paragraph-properties">
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

			<xsl:for-each select ="document($fileName)/child::node()[1]/office:automatic-styles/style:style[@style:name=$gr]/style:graphic-properties">
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

          <xsl:attribute name ="lIns">
            <xsl:call-template name ="convertToPoints">
              <xsl:with-param name ="unit" select ="'cm'"/>
              <xsl:with-param name ="length" select ="$default-padding-left"/>
            </xsl:call-template>
          </xsl:attribute>
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

          <xsl:attribute name ="tIns">
            <xsl:call-template name ="convertToPoints">
              <xsl:with-param name ="unit" select ="'cm'"/>
              <xsl:with-param name ="length" select ="$default-padding-top"/>
            </xsl:call-template>
          </xsl:attribute>
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

          <xsl:attribute name ="rIns">
            <xsl:call-template name ="convertToPoints">
              <xsl:with-param name ="unit" select ="'cm'"/>
              <xsl:with-param name ="length" select ="$default-padding-right"/>
            </xsl:call-template>
          </xsl:attribute>
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

          <xsl:attribute name ="bIns">
            <xsl:call-template name ="convertToPoints">
              <xsl:with-param name ="unit" select ="'cm'"/>
              <xsl:with-param name ="length" select ="$default-padding-bottom"/>
            </xsl:call-template>
          </xsl:attribute>
        </xsl:if>


        <!-- Text wrapping in shape-->
        <!--<xsl:choose>
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
					<xsl:when test ="(@draw:auto-grow-height = 'true') or (@draw:auto-grow-width = 'true')">
						<xsl:attribute name ="wrap">
							<xsl:value-of select ="'square'"/>
						</xsl:attribute>
					</xsl:when>
					<xsl:otherwise>
						<xsl:attribute name ="wrap">
							<xsl:value-of select ="'none'"/>
						</xsl:attribute>
					</xsl:otherwise>  
				</xsl:choose>

				-->
        <!-- AutoFit - Resize to accommodate the text in shape.-->
        <!--
				<xsl:if test ="(@draw:auto-grow-height = 'true') or (@draw:auto-grow-width = 'true') or ($default-autogrow-height = 'true') or ($default-autogrow-width = 'true')">
					<a:spAutoFit/>
				</xsl:if>-->

        <xsl:choose>
          <xsl:when test ="@fo:wrap-option='no-wrap'">
            <xsl:attribute name ="wrap">
              <xsl:value-of select ="'square'"/>
            </xsl:attribute>
          </xsl:when>
          <xsl:when test ="$default-wrap-option='no-wrap'">
            <xsl:attribute name ="wrap">
              <xsl:value-of select ="'square'"/>
            </xsl:attribute>
          </xsl:when>
          <xsl:when test ="(@fo:wrap-option='wrap') or ( ($default-wrap-option='wrap') and ($customShape = 'true') )  or
                            ((@draw:auto-grow-height = 'false') and (@draw:auto-grow-width = 'false'))">
            <xsl:attribute name ="wrap">
              <xsl:value-of select ="'none'"/>
            </xsl:attribute>
          </xsl:when>
          <xsl:otherwise>
            <xsl:attribute name ="wrap">
              <xsl:value-of select ="'square'"/>
            </xsl:attribute>
          </xsl:otherwise>
        </xsl:choose>

        <!-- AutoFit - Resize to accommodate the text in shape.-->
        <xsl:if test ="(@draw:auto-grow-height = 'true') or (@draw:auto-grow-width = 'true') or ($default-autogrow-height = 'true') or ($default-autogrow-width = 'true') or not(@draw:auto-grow-height)">
          <a:spAutoFit/>
        </xsl:if>
      </xsl:for-each>
    </a:bodyPr>

  </xsl:template>
  <!-- Get arrow type-->
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
        <!-- Coded by pradeep for the bug 1742753 -->
        <!-- square to Diamond -->
        <xsl:when test ="$marker = 'm0 0h10v10h-10z'">
          <xsl:value-of select ="'diamond'"/>
        </xsl:when>
        <!-- End of Code by pradeep for the bug 1742753 -->
        <xsl:otherwise>
          <xsl:value-of select ="'triangle'"/>
        </xsl:otherwise>
      </xsl:choose>
    </xsl:if>

	  <xsl:if test ="not(document('styles.xml')/office:document-styles/office:styles/draw:marker[@draw:name=$ArrowType])">
		  <xsl:value-of  select ="'none'"/>	  
	  </xsl:if >
  </xsl:template>
  <!-- Text formatting-->
  <xsl:template name ="processShapeText" >
	<xsl:param name ="fileName" />
    <xsl:param name ="gr" />
    <!-- Paremeter added by vijayeta,get master page name, dated:11-7-07-->
    <xsl:param name ="masterPageName"/>
    <xsl:variable name ="prClsName">		
		<xsl:value-of select ="document($fileName)//office:automatic-styles/style:style[@style:name=$gr]/@style:parent-style-name"/>		
    </xsl:variable>
    <xsl:variable name ="defFontSize">
      <xsl:value-of  select ="office:document-styles/office:styles/style:style/style:text-properties/@fo:font-size"/>
    </xsl:variable>
    <xsl:for-each select ="node()">
      <xsl:if test ="name()='text:p'" >
        <xsl:if test ="child::node()">
        <a:p xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
          <xsl:call-template name ="paraProperties" >
            <xsl:with-param name ="paraId"  select="@text:style-name"/>
            <xsl:with-param name ="isBulleted" select ="'false'"/>
            <xsl:with-param name ="level" select ="'0'"/>
            <xsl:with-param name="framePresentaionStyleId" select="parent::node()/parent::node()/./@presentation:style-name" />
            <xsl:with-param name ="isNumberingEnabled" select ="'false'"/>
            <xsl:with-param name ="slideMaster" select ="$fileName"/>
          </xsl:call-template >
          <xsl:for-each select ="child::node()[position()]">
            <xsl:choose >
              <xsl:when test ="name()='text:span'">
                <xsl:if test="node()">
                  <a:r>
                    <a:rPr lang="en-US" smtClean="0">
                      <!--Font Size -->
                      <xsl:variable name ="textId">
                        <xsl:value-of select ="@text:style-name"/>
                      </xsl:variable>
                      <xsl:if test ="not($textId ='')">
                        <xsl:call-template name ="fontStyles">
                          <xsl:with-param name ="Tid" select ="$textId" />
                          <xsl:with-param name ="prClassName" select ="$prClsName"/>
                          <xsl:with-param name ="slideMaster" select ="$fileName"/>
                          <xsl:with-param name ="fileName" select ="$fileName"/>
                        </xsl:call-template>
                      </xsl:if>
                      <!--<xsl:copy-of select="$varTextHyperLinks"/>-->
                    </a:rPr >
                    <a:t>
                      <xsl:call-template name ="insertTab" />
                    </a:t>
                  </a:r>
                  <xsl:if test ="text:line-break">
                    <xsl:call-template name ="processBR">
                      <xsl:with-param name ="T" select ="@text:style-name" />
                      <xsl:with-param name ="prClassName" select ="$prClsName"/>
                    </xsl:call-template>
                  </xsl:if>
                </xsl:if>
                <!-- Added by lohith - bug fix 1731885 -->
                <xsl:if test="not(node())">
                  <a:endParaRPr lang="en-US" smtClean="0">
                    <!--Font Size -->
                    <xsl:variable name ="textId">
                      <xsl:value-of select ="@text:style-name"/>
                    </xsl:variable>
                    <xsl:if test ="not($textId ='')">
                      <xsl:call-template name ="fontStyles">
                        <xsl:with-param name ="Tid" select ="$textId" />
                        <xsl:with-param name ="prClassName" select ="$prClsName"/>
                        <xsl:with-param name ="slideMaster" select ="$fileName"/>
                      </xsl:call-template>
                    </xsl:if>
                  </a:endParaRPr>
                </xsl:if>

              </xsl:when >
              <xsl:when test ="name()='text:line-break'">
                <xsl:call-template name ="processBR">
                  <xsl:with-param name ="T" select ="@text:style-name" />
                  <xsl:with-param name ="prClassName" select ="$prClsName"/>
                </xsl:call-template>
              </xsl:when>
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
                        <xsl:with-param name ="prClassName" select ="$prClsName"/>
                        <xsl:with-param name ="slideMaster" select ="$fileName"/>
                      </xsl:call-template>
                    </xsl:if>
                    <!--<xsl:copy-of select="$varTextHyperLinks"/>-->
                    <xsl:if test ="$textId =''">
                      <a:latin charset="0"  >
                        <xsl:attribute name ="typeface">
                          <xsl:call-template name ="getDefaultFonaName">
                            <xsl:with-param name ="prClassName" select ="$prClsName"/>
                          </xsl:call-template>
                        </xsl:attribute>
                      </a:latin >
                    </xsl:if >
                  </a:rPr >
                  <a:t>
                    <xsl:call-template name ="insertTab" />
                  </a:t>
                </a:r>
              </xsl:when >
            </xsl:choose>
          </xsl:for-each>
          <!--<xsl:copy-of select="$varFrameHyperLinks"/>-->
        </a:p>
        </xsl:if>
        <!-- Change made by vijayeta
        Bug fix	1731885
        Dated 20th July
        Added a condition to test if text:p has child nodes or not, if not, set a default font to the empty line-->
        <xsl:if test ="not(child::node())">
          <a:p>
            <a:endParaRPr lang="en-US" dirty="0" smtClean="0">
              <a:latin charset="0"  >
                <xsl:attribute name ="typeface">
                  <xsl:call-template name ="getDefaultFonaName">
                    <xsl:with-param name ="prClassName" select ="$prClsName"/>
                    <xsl:with-param name ="masterPageName" select ="$masterPageName"/>
                  </xsl:call-template>
                </xsl:attribute>
              </a:latin >
            </a:endParaRPr>
          </a:p>
        </xsl:if>
        <!-- Change made by vijayeta, Bug fix	1731885-->
      </xsl:if>
      <xsl:if test ="name()='text:list'">
        <a:p xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
          <!--Code inserted by Vijayets for Bullets and numbering-->
          <!--Check if Levels are prese nt-->
          <xsl:variable name ="lvl">
            <xsl:if test ="./text:list-item/text:list">
              <xsl:call-template name ="getListLevel">
                <xsl:with-param name ="levelCount"/>
              </xsl:call-template>
            </xsl:if>
            <xsl:if test ="not(./text:list-item/text:list)">
              <xsl:value-of select ="'0'"/>
            </xsl:if>
          </xsl:variable >
          <xsl:variable name="paragraphId" >
            <xsl:call-template name ="getParaStyleNameForTextBox">
              <xsl:with-param name ="lvl" select ="$lvl"/>
            </xsl:call-template>
          </xsl:variable>
          <xsl:variable name ="isNumberingEnabled">
            <xsl:if test ="document($fileName)//style:style[@style:name=$paragraphId]/style:paragraph-properties/@text:enable-numbering">
              <xsl:value-of select ="document($fileName)//style:style[@style:name=$paragraphId]/style:paragraph-properties/@text:enable-numbering"/>
            </xsl:if>
            <xsl:if test ="not(document($fileName)//style:style[@style:name=$paragraphId]/style:paragraph-properties/@text:enable-numbering)">
              <xsl:value-of select ="'true'"/>
            </xsl:if>
          </xsl:variable>
          <xsl:call-template name ="paraProperties" >
            <xsl:with-param name ="paraId" >
              <xsl:value-of select ="$paragraphId"/>
            </xsl:with-param >
            <!-- list property also included-->
            <xsl:with-param name ="listId">
              <xsl:value-of select ="@text:style-name"/>
            </xsl:with-param >
            <!-- Parameters added by vijayeta,Set bulleting as true/false,and set level -->
            <xsl:with-param name ="isBulleted" select ="'true'"/>
            <xsl:with-param name ="level" select ="$lvl"/>
            <xsl:with-param name ="isNumberingEnabled" select ="$isNumberingEnabled"/>
            <!-- Paremeter added by vijayeta,get master page name, dated:11-7-07-->
            <xsl:with-param name ="masterPageName" select ="$masterPageName"/>
            <xsl:with-param name ="slideMaster" select ="$fileName"/>
          </xsl:call-template >
          <!--End of Code inserted by Vijayets for Bullets and numbering-->
          <xsl:for-each select ="child::node()[position()]">
            <xsl:choose >
              <xsl:when test ="name()='text:list-item'">
                <xsl:variable name ="currentNodeStyle">
                  <xsl:call-template name ="getTextNodeForFontStyle">
                    <xsl:with-param name ="prClassName" select ="$prClsName"/>
                    <xsl:with-param name ="lvl" select ="$lvl"/>
                    <!-- Paremeter added by vijayeta,get master page name, dated:11-7-07-->
                    <xsl:with-param name ="masterPageName" select ="$masterPageName"/>
                    <xsl:with-param name ="fileName" select ="$fileName"/>
                  </xsl:call-template>
                </xsl:variable>
                <xsl:copy-of select ="$currentNodeStyle"/>
              </xsl:when >
            </xsl:choose>
          </xsl:for-each>
          <!--<xsl:copy-of select="$varFrameHyperLinks"/>-->
        </a:p >
      </xsl:if>
    </xsl:for-each>
  </xsl:template>
  <!-- Blank lines in text-->
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
  <!-- Get graphic properties from styles.xml -->
  <xsl:template name ="getDefaultStyle">
    <xsl:param name ="parentStyle" />
    <xsl:param name ="attributeName" />
    <xsl:for-each select ="document('styles.xml')/office:document-styles/office:styles/style:style[@style:name = $parentStyle]/style:graphic-properties">
      <xsl:choose>
        <xsl:when test ="$attributeName = 'fill'">
          <xsl:value-of select ="@draw:fill"/>
        </xsl:when>
        <xsl:when test ="$attributeName = 'padding-left'">
          <xsl:if test ="@fo:padding-left">
            <xsl:value-of select ="@fo:padding-left"/>
          </xsl:if>
          <xsl:if test ="not(@fo:padding-left)">
            <xsl:value-of select ="'0'"/>
          </xsl:if>
        </xsl:when>
        <xsl:when test ="$attributeName = 'padding-top'">
          <xsl:if test ="@fo:padding-top">
            <xsl:value-of select ="@fo:padding-top"/>
          </xsl:if>
          <xsl:if test ="not(@fo:padding-top)">
            <xsl:value-of select ="'0'"/>
          </xsl:if>
        </xsl:when>
        <xsl:when test ="$attributeName = 'padding-right'">
          <xsl:if test ="@fo:padding-right">
            <xsl:value-of select ="@fo:padding-right"/>
          </xsl:if>
          <xsl:if test ="not(@fo:padding-right)">
            <xsl:value-of select ="'0'"/>
          </xsl:if>
        </xsl:when>
        <xsl:when test ="$attributeName = 'padding-bottom'">
          <xsl:if test ="@fo:padding-bottom">
            <xsl:value-of select ="@fo:padding-bottom"/>
          </xsl:if>
          <xsl:if test ="not(@fo:padding-bottom)">
            <xsl:value-of select ="'0'"/>
          </xsl:if>
        </xsl:when>
        <xsl:when test ="$attributeName = 'wrap-option'">
          <xsl:if test ="@fo:wrap-option">
            <xsl:value-of select ="@fo:wrap-option"/>
          </xsl:if>
          <xsl:if test ="not(@fo:wrap-option)">
            <xsl:value-of select ="'wrap'"/>
          </xsl:if>
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

  <!-- Vijayeta,for bullets and numbering,get para style name depending on levels-->
  <xsl:template name ="getParaStyleNameForTextBox">
    <xsl:param name ="lvl"/>
    <xsl:choose>
      <xsl:when test ="$lvl='0'">
        <xsl:value-of select ="./text:list-item/text:p/@text:style-name"/>
      </xsl:when >
      <xsl:when test ="$lvl='1'">
        <xsl:value-of select ="./text:list-item/text:list/text:list-item/text:p/@text:style-name"/>
      </xsl:when>
      <xsl:when test ="$lvl='2'">
        <xsl:value-of select ="./text:list-item/text:list/text:list-item/text:list/text:list-item/text:p/@text:style-name"/>
      </xsl:when>
      <xsl:when test ="$lvl='3'">
        <xsl:value-of select ="./text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:p/@text:style-name"/>
      </xsl:when>
      <xsl:when test ="$lvl='4'">
        <xsl:value-of select ="./text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:p/@text:style-name"/>
      </xsl:when>
      <xsl:when test ="$lvl='5'">
        <xsl:value-of select ="./text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:p/@text:style-name"/>
      </xsl:when>
      <xsl:when test ="$lvl='6'">
        <xsl:value-of select ="./text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:p/@text:style-name"/>
      </xsl:when>
      <xsl:when test ="$lvl='7'">
        <xsl:value-of select ="./text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:p/@text:style-name"/>
      </xsl:when>
      <xsl:when test ="$lvl='8'">
        <xsl:value-of select ="./text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:p/@text:style-name"/>
      </xsl:when>
      <xsl:when test ="$lvl='9'">
        <xsl:value-of select ="./text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:p/@text:style-name"/>
      </xsl:when>
    </xsl:choose>
  </xsl:template>
  <!-- Inner levels,styles-->
  <xsl:template name ="getTextNodeForFontStyle">
    <xsl:param name ="prClassName"/>
    <xsl:param name ="lvl" />
    <xsl:param name ="masterPageName"/>
    <xsl:param name ="fileName"/>
    <xsl:choose>
      <xsl:when test ="./text:p">
        <xsl:for-each select ="./text:p">
          <xsl:for-each select ="child::node()[position()]">
            <xsl:if test ="name()='text:span'">
              <xsl:if test ="not(./text:line-break)">
                <xsl:variable name ="textId">
                  <xsl:value-of select ="@text:style-name"/>
                </xsl:variable>
                <xsl:if test ="child::node()">
                  <a:r>
                    <a:rPr lang="en-US" smtClean="0">
                      <!--Font Size -->
                      <xsl:if test ="not($textId ='')">
                        <xsl:call-template name ="fontStyles">
                          <xsl:with-param name ="Tid" select ="$textId" />
                          <xsl:with-param name ="prClassName" select ="$prClassName"/>
                          <xsl:with-param name ="lvl" select ="$lvl"/>
                          <xsl:with-param name ="masterPageName" select ="$masterPageName"/>
                        </xsl:call-template>
                      </xsl:if>
                      <!--<xsl:copy-of select="$varTextHyperLinks"/>-->
                    </a:rPr >
                    <a:t>
                      <xsl:call-template name ="insertTab" />
                    </a:t>
                  </a:r>
                </xsl:if>
                <xsl:if test ="not(child::node())">
                  <a:endParaRPr lang="en-US" dirty="0">
                    <xsl:attribute name ="sz">
                      <xsl:call-template name ="convertToPoints">
                        <xsl:with-param name ="unit" select ="'pt'"/>
                        <xsl:with-param name ="length" select ="document($fileName)//office:automatic-styles/style:style[@style:name =$textId]/style:text-properties/@fo:font-size"/>
                      </xsl:call-template>
                    </xsl:attribute>
                    <a:latin charset="0" >
                      <xsl:variable name ="typeFace">
                        <xsl:value-of select =" document($fileName)//office:automatic-styles/style:style[@style:name =$textId]/style:text-properties/@fo:font-family" />
                      </xsl:variable>
                      <xsl:attribute name ="typeface">
                        <xsl:value-of select ="translate($typeFace, &quot;'&quot;,'')" />
                      </xsl:attribute>
                    </a:latin >
                  </a:endParaRPr>
                </xsl:if>
              </xsl:if>
              <xsl:if test ="./text:line-break">
                <xsl:call-template name ="processBR">
                  <xsl:with-param name ="T" select ="@text:style-name" />
                  <xsl:with-param name ="prClassName" select ="$prClassName"/>
                </xsl:call-template>
              </xsl:if>
            </xsl:if>
            <xsl:if test ="name()='text:line-break'">
              <xsl:call-template name ="processBR">
                <xsl:with-param name ="T" select ="@text:style-name" />
                <xsl:with-param name ="prClassName" select ="$prClassName"/>
              </xsl:call-template>
            </xsl:if>
            <xsl:if test ="not(name()='text:span' or name()='text:line-break')">
              <a:r>
                <a:rPr lang="en-US" smtClean="0" kern="0">
                  <xsl:variable name ="DefFontSize">
                    <xsl:call-template name ="getDefaultFontSize">
                      <xsl:with-param name ="className" select ="$prClassName"/>
                    </xsl:call-template >
                  </xsl:variable>
                  <xsl:if  test ="$DefFontSize!=''">
                    <xsl:attribute name ="sz">
                      <xsl:call-template name ="convertToPoints">
                        <xsl:with-param name ="unit" select ="'pt'"/>
                        <xsl:with-param name ="length" select ="$DefFontSize"/>
                      </xsl:call-template>
                    </xsl:attribute>
                  </xsl:if>
                  <a:latin charset="0" >
                    <xsl:attribute name ="typeface">
                      <xsl:call-template name ="getDefaultFonaName">
                        <xsl:with-param name ="className" select ="$prClassName"/>
                        <!-- Node added by vijayeta,to insert font sizes to inner levels-->
                        <xsl:with-param name ="lvl" select ="0"/>
                      </xsl:call-template >
                    </xsl:attribute>
                  </a:latin >
                  <!--<xsl:copy-of select="$varTextHyperLinks"/>-->
                </a:rPr >
                <a:t>
                  <xsl:call-template name ="insertTab" />
                </a:t>
              </a:r>
            </xsl:if >
          </xsl:for-each>
        </xsl:for-each>
      </xsl:when>
      <xsl:when test ="./text:list/text:list-item/text:p">
        <xsl:for-each select ="./text:list/text:list-item/text:p">
          <xsl:for-each select ="child::node()[position()]">
            <xsl:if test ="name()='text:span'">
              <xsl:if test ="not(./text:line-break)">
                <xsl:variable name ="textId">
                  <xsl:value-of select ="@text:style-name"/>
                </xsl:variable>
                <xsl:if test ="child::node()">
                  <a:r>
                    <a:rPr lang="en-US" smtClean="0">
                      <!--Font Size -->
                      <xsl:if test ="not($textId ='')">
                        <xsl:call-template name ="fontStyles">
                          <xsl:with-param name ="Tid" select ="$textId" />
                          <xsl:with-param name ="prClassName" select ="$prClassName"/>
                          <xsl:with-param name ="lvl" select ="$lvl"/>
                          <xsl:with-param name ="masterPageName" select ="$masterPageName"/>
                        </xsl:call-template>
                      </xsl:if>
                      <!--<xsl:copy-of select="$varTextHyperLinks"/>-->
                    </a:rPr >
                    <a:t>
                      <xsl:call-template name ="insertTab" />
                    </a:t>
                  </a:r>
                </xsl:if>
                <xsl:if test ="not(child::node())">
                  <a:endParaRPr lang="en-US" dirty="0">
                    <xsl:attribute name ="sz">
                      <xsl:call-template name ="convertToPoints">
                        <xsl:with-param name ="unit" select ="'pt'"/>
                        <xsl:with-param name ="length" select ="document($fileName)//office:automatic-styles/style:style[@style:name =$textId]/style:text-properties/@fo:font-size"/>
                      </xsl:call-template>
                    </xsl:attribute>
                    <a:latin charset="0" >
                      <xsl:variable name ="typeFace">
                        <xsl:value-of select =" document($fileName)//office:automatic-styles/style:style[@style:name =$textId]/style:text-properties/@fo:font-family" />
                      </xsl:variable>
                      <xsl:attribute name ="typeface">
                        <xsl:value-of select ="translate($typeFace, &quot;'&quot;,'')" />
                      </xsl:attribute>
                    </a:latin >
                  </a:endParaRPr>
                </xsl:if>
              </xsl:if>             
              <xsl:if test ="./text:line-break">
                <xsl:call-template name ="processBR">
                  <xsl:with-param name ="T" select ="@text:style-name" />
                  <xsl:with-param name ="prClassName" select ="$prClassName"/>
                </xsl:call-template>
              </xsl:if>
            </xsl:if>
            <xsl:if test ="name()='text:line-break'">
              <xsl:call-template name ="processBR">
                <xsl:with-param name ="T" select ="@text:style-name" />
                <xsl:with-param name ="prClassName" select ="$prClassName"/>
              </xsl:call-template>
            </xsl:if>
            <xsl:if test ="not(name()='text:span' or name()='text:line-break')">
              <a:r>
                <a:rPr lang="en-US" smtClean="0">
                  <!--Font Size -->
                  <xsl:variable name ="textId">
                    <xsl:value-of select ="@text:style-name"/>
                  </xsl:variable>
                  <xsl:if test ="not($textId ='')">
                    <xsl:call-template name ="fontStyles">
                      <xsl:with-param name ="Tid" select ="$textId" />
                      <xsl:with-param name ="prClassName" select ="$prClassName"/>
                    </xsl:call-template>
                  </xsl:if>
                </a:rPr >
                <a:t>
                  <xsl:call-template name ="insertTab" />
                </a:t>
              </a:r>
            </xsl:if >
          </xsl:for-each>
        </xsl:for-each>
      </xsl:when>
      <xsl:when test ="./text:list/text:list-item/text:list/text:list-item/text:p">
        <xsl:for-each select ="./text:list/text:list-item/text:list/text:list-item/text:p">
          <xsl:for-each select ="child::node()[position()]">
            <xsl:if test ="name()='text:span'">
              <xsl:if test ="not(./text:line-break)">
                <xsl:variable name ="textId">
                  <xsl:value-of select ="@text:style-name"/>
                </xsl:variable>
                <xsl:if test ="child::node()">
                  <a:r>
                    <a:rPr lang="en-US" smtClean="0">
                      <!--Font Size -->
                      <xsl:if test ="not($textId ='')">
                        <xsl:call-template name ="fontStyles">
                          <xsl:with-param name ="Tid" select ="$textId" />
                          <xsl:with-param name ="prClassName" select ="$prClassName"/>
                          <xsl:with-param name ="lvl" select ="$lvl"/>
                          <xsl:with-param name ="masterPageName" select ="$masterPageName"/>
                        </xsl:call-template>
                      </xsl:if>
                      <!--<xsl:copy-of select="$varTextHyperLinks"/>-->
                    </a:rPr >
                    <a:t>
                      <xsl:call-template name ="insertTab" />
                    </a:t>
                  </a:r>
                </xsl:if>
                <xsl:if test ="not(child::node())">
                  <a:endParaRPr lang="en-US" dirty="0">
                    <xsl:attribute name ="sz">
                      <xsl:call-template name ="convertToPoints">
                        <xsl:with-param name ="unit" select ="'pt'"/>
                        <xsl:with-param name ="length" select ="document($fileName)//office:automatic-styles/style:style[@style:name =$textId]/style:text-properties/@fo:font-size"/>
                      </xsl:call-template>
                    </xsl:attribute>
                    <a:latin charset="0" >
                      <xsl:variable name ="typeFace">
                        <xsl:value-of select =" document($fileName)//office:automatic-styles/style:style[@style:name =$textId]/style:text-properties/@fo:font-family" />
                      </xsl:variable>
                      <xsl:attribute name ="typeface">
                        <xsl:value-of select ="translate($typeFace, &quot;'&quot;,'')" />
                      </xsl:attribute>
                    </a:latin >
                  </a:endParaRPr>
                </xsl:if>
              </xsl:if>
              <xsl:if test ="./text:line-break">
                <xsl:call-template name ="processBR">
                  <xsl:with-param name ="T" select ="@text:style-name" />
                  <xsl:with-param name ="prClassName" select ="$prClassName"/>
                </xsl:call-template>
              </xsl:if>
            </xsl:if>
            <xsl:if test ="name()='text:line-break'">
              <xsl:call-template name ="processBR">
                <xsl:with-param name ="T" select ="@text:style-name" />
                <xsl:with-param name ="prClassName" select ="$prClassName"/>
              </xsl:call-template>
            </xsl:if>
            <xsl:if test ="not(name()='text:span' or name()='text:line-break')">
              <a:r>
                <a:rPr lang="en-US" smtClean="0">
                  <!--Font Size -->
                  <xsl:variable name ="textId">
                    <xsl:value-of select ="@text:style-name"/>
                  </xsl:variable>
                  <xsl:if test ="not($textId ='')">
                    <xsl:call-template name ="fontStyles">
                      <xsl:with-param name ="Tid" select ="$textId" />
                      <xsl:with-param name ="prClassName" select ="$prClassName"/>
                    </xsl:call-template>
                  </xsl:if>
                </a:rPr >
                <a:t>
                  <xsl:call-template name ="insertTab" />
                </a:t>
              </a:r>
            </xsl:if >
          </xsl:for-each>
        </xsl:for-each>
      </xsl:when>
      <xsl:when test ="./text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:p">
        <xsl:for-each select ="./text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:p">
          <xsl:for-each select ="child::node()[position()]">
            <xsl:if test ="name()='text:span'">
              <xsl:if test ="not(./text:line-break)">
                <xsl:variable name ="textId">
                  <xsl:value-of select ="@text:style-name"/>
                </xsl:variable>
                <xsl:if test ="child::node()">
                  <a:r>
                    <a:rPr lang="en-US" smtClean="0">
                      <!--Font Size -->
                      <xsl:if test ="not($textId ='')">
                        <xsl:call-template name ="fontStyles">
                          <xsl:with-param name ="Tid" select ="$textId" />
                          <xsl:with-param name ="prClassName" select ="$prClassName"/>
                          <xsl:with-param name ="lvl" select ="$lvl"/>
                          <xsl:with-param name ="masterPageName" select ="$masterPageName"/>
                        </xsl:call-template>
                      </xsl:if>
                      <!--<xsl:copy-of select="$varTextHyperLinks"/>-->
                    </a:rPr >
                    <a:t>
                      <xsl:call-template name ="insertTab" />
                    </a:t>
                  </a:r>
                </xsl:if>
                <xsl:if test ="not(child::node())">
                  <a:endParaRPr lang="en-US" dirty="0">
                    <xsl:attribute name ="sz">
                      <xsl:call-template name ="convertToPoints">
                        <xsl:with-param name ="unit" select ="'pt'"/>
                        <xsl:with-param name ="length" select ="document($fileName)//office:automatic-styles/style:style[@style:name =$textId]/style:text-properties/@fo:font-size"/>
                      </xsl:call-template>
                    </xsl:attribute>
                    <a:latin charset="0" >
                      <xsl:variable name ="typeFace">
                        <xsl:value-of select =" document($fileName)//office:automatic-styles/style:style[@style:name =$textId]/style:text-properties/@fo:font-family" />
                      </xsl:variable>
                      <xsl:attribute name ="typeface">
                        <xsl:value-of select ="translate($typeFace, &quot;'&quot;,'')" />
                      </xsl:attribute>
                    </a:latin >
                  </a:endParaRPr>
                </xsl:if>
              </xsl:if>
              <xsl:if test ="./text:line-break">
                <xsl:call-template name ="processBR">
                  <xsl:with-param name ="T" select ="@text:style-name" />
                  <xsl:with-param name ="prClassName" select ="$prClassName"/>
                </xsl:call-template>
              </xsl:if>
            </xsl:if>
            <xsl:if test ="name()='text:line-break'">
              <xsl:call-template name ="processBR">
                <xsl:with-param name ="T" select ="@text:style-name" />
                <xsl:with-param name ="prClassName" select ="$prClassName"/>
              </xsl:call-template>
            </xsl:if>
            <xsl:if test ="not(name()='text:span' or name()='text:line-break')">
              <a:r>
                <a:rPr lang="en-US" smtClean="0">
                  <!--Font Size -->
                  <xsl:variable name ="textId">
                    <xsl:value-of select ="@text:style-name"/>
                  </xsl:variable>
                  <xsl:if test ="not($textId ='')">
                    <xsl:call-template name ="fontStyles">
                      <xsl:with-param name ="Tid" select ="$textId" />
                      <xsl:with-param name ="prClassName" select ="$prClassName"/>
                    </xsl:call-template>
                  </xsl:if>
                </a:rPr >
                <a:t>
                  <xsl:call-template name ="insertTab" />
                </a:t>
              </a:r>
            </xsl:if >
          </xsl:for-each>
        </xsl:for-each>
      </xsl:when>
      <xsl:when test ="./text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:p">
        <xsl:for-each select ="./text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:p">
          <xsl:for-each select ="child::node()[position()]">
            <xsl:if test ="name()='text:line-break'">
              <xsl:call-template name ="processBR">
                <xsl:with-param name ="T" select ="@text:style-name" />
                <xsl:with-param name ="prClassName" select ="$prClassName"/>
              </xsl:call-template>
            </xsl:if>
            <xsl:if test ="name()='text:span'">
              <xsl:if test ="not(./text:line-break)">
                <xsl:variable name ="textId">
                  <xsl:value-of select ="@text:style-name"/>
                </xsl:variable>
                <xsl:if test ="child::node()">
                  <a:r>
                    <a:rPr lang="en-US" smtClean="0">
                      <!--Font Size -->
                      <xsl:if test ="not($textId ='')">
                        <xsl:call-template name ="fontStyles">
                          <xsl:with-param name ="Tid" select ="$textId" />
                          <xsl:with-param name ="prClassName" select ="$prClassName"/>
                          <xsl:with-param name ="lvl" select ="$lvl"/>
                          <xsl:with-param name ="masterPageName" select ="$masterPageName"/>
                        </xsl:call-template>
                      </xsl:if>
                      <!--<xsl:copy-of select="$varTextHyperLinks"/>-->
                    </a:rPr >
                    <a:t>
                      <xsl:call-template name ="insertTab" />
                    </a:t>
                  </a:r>
                </xsl:if>
                <xsl:if test ="not(child::node())">
                  <a:endParaRPr lang="en-US" dirty="0">
                    <xsl:attribute name ="sz">
                      <xsl:call-template name ="convertToPoints">
                        <xsl:with-param name ="unit" select ="'pt'"/>
                        <xsl:with-param name ="length" select ="document($fileName)//office:automatic-styles/style:style[@style:name =$textId]/style:text-properties/@fo:font-size"/>
                      </xsl:call-template>
                    </xsl:attribute>
                    <a:latin charset="0" >
                      <xsl:variable name ="typeFace">
                        <xsl:value-of select =" document($fileName)//office:automatic-styles/style:style[@style:name =$textId]/style:text-properties/@fo:font-family" />
                      </xsl:variable>
                      <xsl:attribute name ="typeface">
                        <xsl:value-of select ="translate($typeFace, &quot;'&quot;,'')" />
                      </xsl:attribute>
                    </a:latin >
                  </a:endParaRPr>
                </xsl:if>
              </xsl:if>
              <xsl:if test ="./text:line-break">
                <xsl:call-template name ="processBR">
                  <xsl:with-param name ="T" select ="@text:style-name" />
                  <xsl:with-param name ="prClassName" select ="$prClassName"/>
                </xsl:call-template>
              </xsl:if>
            </xsl:if>
            <xsl:if test ="not(name()='text:span' or name()='text:line-break')">
              <a:r>
                <a:rPr lang="en-US" smtClean="0">
                  <!--Font Size -->
                  <xsl:variable name ="textId">
                    <xsl:value-of select ="@text:style-name"/>
                  </xsl:variable>
                  <xsl:if test ="not($textId ='')">
                    <xsl:call-template name ="fontStyles">
                      <xsl:with-param name ="Tid" select ="$textId" />
                      <xsl:with-param name ="prClassName" select ="$prClassName"/>
                    </xsl:call-template>
                  </xsl:if>
                </a:rPr >
                <a:t>
                  <xsl:call-template name ="insertTab" />
                </a:t>
              </a:r>
            </xsl:if >
          </xsl:for-each>
        </xsl:for-each>
      </xsl:when>
      <xsl:when test ="./text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:p">
        <xsl:for-each select ="./text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:p">
          <xsl:for-each select ="child::node()[position()]">
            <xsl:if test ="name()='text:span'">
              <xsl:if test ="not(./text:line-break)">
                <xsl:variable name ="textId">
                  <xsl:value-of select ="@text:style-name"/>
                </xsl:variable>
                <xsl:if test ="child::node()">
                  <a:r>
                    <a:rPr lang="en-US" smtClean="0">
                      <!--Font Size -->
                      <xsl:if test ="not($textId ='')">
                        <xsl:call-template name ="fontStyles">
                          <xsl:with-param name ="Tid" select ="$textId" />
                          <xsl:with-param name ="prClassName" select ="$prClassName"/>
                          <xsl:with-param name ="lvl" select ="$lvl"/>
                          <xsl:with-param name ="masterPageName" select ="$masterPageName"/>
                        </xsl:call-template>
                      </xsl:if>
                      <!--<xsl:copy-of select="$varTextHyperLinks"/>-->
                    </a:rPr >
                    <a:t>
                      <xsl:call-template name ="insertTab" />
                    </a:t>
                  </a:r>
                </xsl:if>
                <xsl:if test ="not(child::node())">
                  <a:endParaRPr lang="en-US" dirty="0">
                    <xsl:attribute name ="sz">
                      <xsl:call-template name ="convertToPoints">
                        <xsl:with-param name ="unit" select ="'pt'"/>
                        <xsl:with-param name ="length" select ="document($fileName)//office:automatic-styles/style:style[@style:name =$textId]/style:text-properties/@fo:font-size"/>
                      </xsl:call-template>
                    </xsl:attribute>
                    <a:latin charset="0" >
                      <xsl:variable name ="typeFace">
                        <xsl:value-of select =" document($fileName)//office:automatic-styles/style:style[@style:name =$textId]/style:text-properties/@fo:font-family" />
                      </xsl:variable>
                      <xsl:attribute name ="typeface">
                        <xsl:value-of select ="translate($typeFace, &quot;'&quot;,'')" />
                      </xsl:attribute>
                    </a:latin >
                  </a:endParaRPr>
                </xsl:if>
              </xsl:if>
              <xsl:if test ="./text:line-break">
                <xsl:call-template name ="processBR">
                  <xsl:with-param name ="T" select ="@text:style-name" />
                  <xsl:with-param name ="prClassName" select ="$prClassName"/>
                </xsl:call-template>
              </xsl:if>
            </xsl:if>
            <xsl:if test ="name()='text:line-break'">
              <xsl:call-template name ="processBR">
                <xsl:with-param name ="T" select ="@text:style-name" />
                <xsl:with-param name ="prClassName" select ="$prClassName"/>
              </xsl:call-template>
            </xsl:if>
            <xsl:if test ="not(name()='text:span' or name()='text:line-break')">
              <a:r>
                <a:rPr lang="en-US" smtClean="0">
                  <!--Font Size -->
                  <xsl:variable name ="textId">
                    <xsl:value-of select ="@text:style-name"/>
                  </xsl:variable>
                  <xsl:if test ="not($textId ='')">
                    <xsl:call-template name ="fontStyles">
                      <xsl:with-param name ="Tid" select ="$textId" />
                      <xsl:with-param name ="prClassName" select ="$prClassName"/>
                    </xsl:call-template>
                  </xsl:if>
                </a:rPr >
                <a:t>
                  <xsl:call-template name ="insertTab" />
                </a:t>
              </a:r>
            </xsl:if >
          </xsl:for-each>
        </xsl:for-each>
      </xsl:when>
      <xsl:when test ="./text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:p">
        <xsl:for-each select ="./text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:p">
          <xsl:for-each select ="child::node()[position()]">
            <xsl:if test ="name()='text:span'">
              <xsl:if test ="not(./text:line-break)">
                <xsl:variable name ="textId">
                  <xsl:value-of select ="@text:style-name"/>
                </xsl:variable>
                <xsl:if test ="child::node()">
                  <a:r>
                    <a:rPr lang="en-US" smtClean="0">
                      <!--Font Size -->
                      <xsl:if test ="not($textId ='')">
                        <xsl:call-template name ="fontStyles">
                          <xsl:with-param name ="Tid" select ="$textId" />
                          <xsl:with-param name ="prClassName" select ="$prClassName"/>
                          <xsl:with-param name ="lvl" select ="$lvl"/>
                          <xsl:with-param name ="masterPageName" select ="$masterPageName"/>
                        </xsl:call-template>
                      </xsl:if>
                      <!--<xsl:copy-of select="$varTextHyperLinks"/>-->
                    </a:rPr >
                    <a:t>
                      <xsl:call-template name ="insertTab" />
                    </a:t>
                  </a:r>
                </xsl:if>
                <xsl:if test ="not(child::node())">
                  <a:endParaRPr lang="en-US" dirty="0">
                    <xsl:attribute name ="sz">
                      <xsl:call-template name ="convertToPoints">
                        <xsl:with-param name ="unit" select ="'pt'"/>
                        <xsl:with-param name ="length" select ="document($fileName)//office:automatic-styles/style:style[@style:name =$textId]/style:text-properties/@fo:font-size"/>
                      </xsl:call-template>
                    </xsl:attribute>
                    <a:latin charset="0" >
                      <xsl:variable name ="typeFace">
                        <xsl:value-of select =" document($fileName)//office:automatic-styles/style:style[@style:name =$textId]/style:text-properties/@fo:font-family" />
                      </xsl:variable>
                      <xsl:attribute name ="typeface">
                        <xsl:value-of select ="translate($typeFace, &quot;'&quot;,'')" />
                      </xsl:attribute>
                    </a:latin >
                  </a:endParaRPr>
                </xsl:if>
              </xsl:if>
              <xsl:if test ="./text:line-break">
                <xsl:call-template name ="processBR">
                  <xsl:with-param name ="T" select ="@text:style-name" />
                  <xsl:with-param name ="prClassName" select ="$prClassName"/>
                </xsl:call-template>
              </xsl:if>
            </xsl:if>
            <xsl:if test ="name()='text:line-break'">
              <xsl:call-template name ="processBR">
                <xsl:with-param name ="T" select ="@text:style-name" />
                <xsl:with-param name ="prClassName" select ="$prClassName"/>
              </xsl:call-template>
            </xsl:if>
            <xsl:if test ="not(name()='text:span' or name()='text:line-break')">
              <a:r>
                <a:rPr lang="en-US" smtClean="0">
                  <!--Font Size -->
                  <xsl:variable name ="textId">
                    <xsl:value-of select ="@text:style-name"/>
                  </xsl:variable>
                  <xsl:if test ="not($textId ='')">
                    <xsl:call-template name ="fontStyles">
                      <xsl:with-param name ="Tid" select ="$textId" />
                      <xsl:with-param name ="prClassName" select ="$prClassName"/>
                    </xsl:call-template>
                  </xsl:if>
                </a:rPr >
                <a:t>
                  <xsl:call-template name ="insertTab" />
                </a:t>
              </a:r>
            </xsl:if >
          </xsl:for-each>
        </xsl:for-each>
      </xsl:when>
      <xsl:when test ="./text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:p">
        <xsl:for-each select ="./text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:p">
          <xsl:for-each select ="child::node()[position()]">
            <xsl:if test ="name()='text:span'">
              <xsl:if test ="not(./text:line-break)">
                <xsl:variable name ="textId">
                  <xsl:value-of select ="@text:style-name"/>
                </xsl:variable>
                <xsl:if test ="child::node()">
                  <a:r>
                    <a:rPr lang="en-US" smtClean="0">
                      <!--Font Size -->
                      <xsl:if test ="not($textId ='')">
                        <xsl:call-template name ="fontStyles">
                          <xsl:with-param name ="Tid" select ="$textId" />
                          <xsl:with-param name ="prClassName" select ="$prClassName"/>
                          <xsl:with-param name ="lvl" select ="$lvl"/>
                          <xsl:with-param name ="masterPageName" select ="$masterPageName"/>
                        </xsl:call-template>
                      </xsl:if>
                      <!--<xsl:copy-of select="$varTextHyperLinks"/>-->
                    </a:rPr >
                    <a:t>
                      <xsl:call-template name ="insertTab" />
                    </a:t>
                  </a:r>
                </xsl:if>
                <xsl:if test ="not(child::node())">
                  <a:endParaRPr lang="en-US" dirty="0">
                    <xsl:attribute name ="sz">
                      <xsl:call-template name ="convertToPoints">
                        <xsl:with-param name ="unit" select ="'pt'"/>
                        <xsl:with-param name ="length" select ="document($fileName)//office:automatic-styles/style:style[@style:name =$textId]/style:text-properties/@fo:font-size"/>
                      </xsl:call-template>
                    </xsl:attribute>
                    <a:latin charset="0" >
                      <xsl:variable name ="typeFace">
                        <xsl:value-of select =" document($fileName)//office:automatic-styles/style:style[@style:name =$textId]/style:text-properties/@fo:font-family" />
                      </xsl:variable>
                      <xsl:attribute name ="typeface">
                        <xsl:value-of select ="translate($typeFace, &quot;'&quot;,'')" />
                      </xsl:attribute>
                    </a:latin >
                  </a:endParaRPr>
                </xsl:if>
              </xsl:if>
              <xsl:if test ="./text:line-break">
                <xsl:call-template name ="processBR">
                  <xsl:with-param name ="T" select ="@text:style-name" />
                  <xsl:with-param name ="prClassName" select ="$prClassName"/>
                </xsl:call-template>
              </xsl:if>
            </xsl:if>
            <xsl:if test ="name()='text:line-break'">
              <xsl:call-template name ="processBR">
                <xsl:with-param name ="T" select ="@text:style-name" />
                <xsl:with-param name ="prClassName" select ="$prClassName"/>
              </xsl:call-template>
            </xsl:if>
            <xsl:if test ="not(name()='text:span' or name()='text:line-break')">
              <a:r>
                <a:rPr lang="en-US" smtClean="0">
                  <!--Font Size -->
                  <xsl:variable name ="textId">
                    <xsl:value-of select ="@text:style-name"/>
                  </xsl:variable>
                  <xsl:if test ="not($textId ='')">
                    <xsl:call-template name ="fontStyles">
                      <xsl:with-param name ="Tid" select ="$textId" />
                      <xsl:with-param name ="prClassName" select ="$prClassName"/>
                    </xsl:call-template>
                  </xsl:if>
                </a:rPr >
                <a:t>
                  <xsl:call-template name ="insertTab" />
                </a:t>
              </a:r>
            </xsl:if >
          </xsl:for-each>
        </xsl:for-each>
      </xsl:when>
      <xsl:when test ="./text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:p">
        <xsl:for-each select ="./text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:p">
          <xsl:for-each select ="child::node()[position()]">
            <xsl:if test ="name()='text:span'">
              <xsl:if test ="not(./text:line-break)">
                <xsl:variable name ="textId">
                  <xsl:value-of select ="@text:style-name"/>
                </xsl:variable>
                <xsl:if test ="child::node()">
                  <a:r>
                    <a:rPr lang="en-US" smtClean="0">
                      <!--Font Size -->
                      <xsl:if test ="not($textId ='')">
                        <xsl:call-template name ="fontStyles">
                          <xsl:with-param name ="Tid" select ="$textId" />
                          <xsl:with-param name ="prClassName" select ="$prClassName"/>
                          <xsl:with-param name ="lvl" select ="$lvl"/>
                          <xsl:with-param name ="masterPageName" select ="$masterPageName"/>
                        </xsl:call-template>
                      </xsl:if>
                      <!--<xsl:copy-of select="$varTextHyperLinks"/>-->
                    </a:rPr >
                    <a:t>
                      <xsl:call-template name ="insertTab" />
                    </a:t>
                  </a:r>
                </xsl:if>
                <xsl:if test ="not(child::node())">
                  <a:endParaRPr lang="en-US" dirty="0">
                    <xsl:attribute name ="sz">
                      <xsl:call-template name ="convertToPoints">
                        <xsl:with-param name ="unit" select ="'pt'"/>
                        <xsl:with-param name ="length" select ="document($fileName)//office:automatic-styles/style:style[@style:name =$textId]/style:text-properties/@fo:font-size"/>
                      </xsl:call-template>
                    </xsl:attribute>
                    <a:latin charset="0" >
                      <xsl:variable name ="typeFace">
                        <xsl:value-of select =" document($fileName)//office:automatic-styles/style:style[@style:name =$textId]/style:text-properties/@fo:font-family" />
                      </xsl:variable>
                      <xsl:attribute name ="typeface">
                        <xsl:value-of select ="translate($typeFace, &quot;'&quot;,'')" />
                      </xsl:attribute>
                    </a:latin >
                  </a:endParaRPr>
                </xsl:if>
              </xsl:if>
              <xsl:if test ="./text:line-break">
                <xsl:call-template name ="processBR">
                  <xsl:with-param name ="T" select ="@text:style-name" />
                  <xsl:with-param name ="prClassName" select ="$prClassName"/>
                </xsl:call-template>
              </xsl:if>
            </xsl:if>
            <xsl:if test ="name()='text:line-break'">
              <xsl:call-template name ="processBR">
                <xsl:with-param name ="T" select ="@text:style-name" />
                <xsl:with-param name ="prClassName" select ="$prClassName"/>
              </xsl:call-template>
            </xsl:if>
            <xsl:if test ="not(name()='text:span' or name()='text:line-break')">
              <a:r>
                <a:rPr lang="en-US" smtClean="0">
                  <!--Font Size -->
                  <xsl:variable name ="textId">
                    <xsl:value-of select ="@text:style-name"/>
                  </xsl:variable>
                  <xsl:if test ="not($textId ='')">
                    <xsl:call-template name ="fontStyles">
                      <xsl:with-param name ="Tid" select ="$textId" />
                      <xsl:with-param name ="prClassName" select ="$prClassName"/>
                    </xsl:call-template>
                  </xsl:if>
                </a:rPr >
                <a:t>
                  <xsl:call-template name ="insertTab" />
                </a:t>
              </a:r>
            </xsl:if >
          </xsl:for-each>
        </xsl:for-each>
      </xsl:when>
    </xsl:choose>
  </xsl:template>
  <!-- End of code by Vijayeta,for bullets and numbering,get para style name depending on levels-->
</xsl:stylesheet >