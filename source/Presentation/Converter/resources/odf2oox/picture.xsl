<?xml version="1.0" encoding="UTF-8" standalone="yes" ?>
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
  exclude-result-prefixes="odf style text number draw page svg presentation fo script xlink">
  <xsl:import href ="common.xsl"/>
  <xsl:import href ="shapes_direct.xsl"/>
  <xsl:template name ="InsertPicture">    
    <xsl:param name ="imageNo" />
    <xsl:param name ="picNo" />
    <xsl:param name ="master" />
 <xsl:param name ="NvPrId" />
    <!-- warn if Audio or Video -->
    <xsl:message terminate="no">translation.odf2oox.audioVideoTypeImage</xsl:message>
    <xsl:variable name ="imageSerialNo">
      <xsl:value-of select ="concat('sl',$imageNo,'Image' ,$picNo)"/>
    </xsl:variable>
    <pzip:copy pzip:source="{@xlink:href}"
				   pzip:target="{concat('ppt/media/',substring-after(@xlink:href,'/'))}"/>
    <p:pic>
      <p:nvPicPr>
        <p:cNvPr id="{$picNo + 1 }" name="Placeholder 3">
          <xsl:attribute name ="descr">
            <xsl:if test ="parent::node()/@draw:name">
              <xsl:value-of select ="parent::node()/@draw:name"/>
            </xsl:if>
            <xsl:if test ="not(parent::node()/@draw:name)">
              <xsl:value-of  select ="substring-after(@xlink:href,'/')"/>
            </xsl:if >
          </xsl:attribute>
        </p:cNvPr >
        <p:cNvPicPr>
          <a:picLocks noGrp="1" noChangeAspect="1" />
        </p:cNvPicPr>
        <xsl:if test="$master='1'">
          <p:nvPr userDrawn="1"/>
        </xsl:if>
        <xsl:if test="not($master)">
          <p:nvPr>
            <p:ph idx="{$picNo+5}" />
          </p:nvPr>
        </xsl:if>
      </p:nvPicPr>
      <p:blipFill>
        <a:blip>
          <xsl:attribute name ="r:embed">
            <xsl:value-of select ="$imageSerialNo"/>
          </xsl:attribute>
        </a:blip >
        <a:stretch>
          <a:fillRect />
        </a:stretch>
      </p:blipFill>
      <p:spPr>
        <a:xfrm>
          <xsl:variable name ="angle">
            <xsl:value-of select="substring-after(substring-before(substring-before(parent::node()/@draw:transform,'translate'),')'),'(')" />
          </xsl:variable>
          <xsl:variable name ="x2">
            <xsl:value-of select="substring-before(substring-before(substring-before(substring-after(substring-after(parent::node()/@draw:transform,'translate'),'('),')'),' '),'cm')" />
          </xsl:variable>
          <xsl:variable name ="y2">
            <xsl:value-of select="substring-before(substring-after(substring-before(substring-after(substring-after(parent::node()/@draw:transform,'translate'),'('),')'),' '),'cm')" />
          </xsl:variable>
          <xsl:if test="parent::node()/@draw:transform">
            <xsl:attribute name ="rot">
              <xsl:value-of select ="concat('draw-transform:ROT:',substring-before(parent::node()/@svg:width,'cm'), ':',
																   substring-before(parent::node()/@svg:height,'cm'), ':', 
																   $x2, ':',  $y2, ':', $angle)"/>
            </xsl:attribute>
          </xsl:if>
          <a:off>
            <xsl:if test="not(parent::node()/@draw:transform)">
              <xsl:attribute name ="x">
                <!--<xsl:value-of select ="@svg:x"/>-->
                <xsl:call-template name ="convertToPoints">
                  <xsl:with-param name ="unit" select ="'cm'"/>
                  <xsl:with-param name ="length" select ="parent::node()/@svg:x"/>
                </xsl:call-template>
              </xsl:attribute>
              <xsl:attribute name ="y">
                <!--<xsl:value-of select ="@svg:y"/>-->
                <xsl:call-template name ="convertToPoints">
                  <xsl:with-param name ="unit" select ="'cm'"/>
                  <xsl:with-param name ="length" select ="parent::node()/@svg:y"/>
                </xsl:call-template>
              </xsl:attribute>
            </xsl:if >
            <xsl:if test="parent::node()/@draw:transform">
              <xsl:attribute name ="x">
                <xsl:value-of select ="concat('draw-transform:X:',substring-before(parent::node()/@svg:width,'cm'), ':',
																   substring-before(parent::node()/@svg:height,'cm'), ':', 
																   $x2, ':', $y2, ':', $angle)"/>
              </xsl:attribute>
              <xsl:attribute name ="y">
                <xsl:value-of select ="concat('draw-transform:Y:',substring-before(parent::node()/@svg:width,'cm'), ':',
																   substring-before(parent::node()/@svg:height,'cm'), ':', 
																   $x2, ':',$y2, ':', $angle)"/>
              </xsl:attribute>
            </xsl:if>
          </a:off >
          <a:ext>
            <xsl:attribute name ="cx">
              <!--<xsl:value-of select ="@svg:width"/>-->
              <xsl:call-template name ="convertToPoints">
                <xsl:with-param name ="unit" select ="'cm'"/>
                <xsl:with-param name ="length" select ="parent::node()/@svg:width"/>
              </xsl:call-template>
            </xsl:attribute>
            <xsl:attribute name ="cy">
              <!--<xsl:value-of select ="@svg:height"/>-->
              <xsl:call-template name ="convertToPoints">
                <xsl:with-param name ="unit" select ="'cm'"/>
                <xsl:with-param name ="length" select ="parent::node()/@svg:height"/>
              </xsl:call-template>
            </xsl:attribute>
          </a:ext >
        </a:xfrm>        
        <!--Sateesh-->
        <a:prstGeom prst="rect">
          <a:avLst/>
        </a:prstGeom>        
        <xsl:variable name="varFileName">
          <xsl:choose>
            <xsl:when test="$master">
              <xsl:value-of select="'styles.xml'"/>
            </xsl:when>
            <xsl:otherwise>
              <xsl:value-of select="'content.xml'"/>
            </xsl:otherwise>
          </xsl:choose>
        </xsl:variable>
        <xsl:call-template name ="pictureBorderLine">
          <xsl:with-param name="FileName" select="$varFileName"/>
        </xsl:call-template>
      </p:spPr>
    </p:pic>
  </xsl:template>
  <xsl:template name="InsertAudio">
    <xsl:param name ="imageNo" />
    <xsl:param name ="AudNo" />
    <xsl:param name ="NvPrId" />
    <xsl:variable name="PostionCount">
      <xsl:value-of select="$AudNo"/>
    </xsl:variable>
    <xsl:variable name ="audioSerialNo">
      <xsl:value-of select ="concat('sl',$imageNo,'Au' ,$AudNo)"/>
    </xsl:variable>
    <xsl:variable name ="imageSerialNo">
      <xsl:value-of select ="concat('sl',$imageNo,'Im' ,$AudNo)"/>
    </xsl:variable>
    <pzip:copy pzip:source="{'Thumbnails/thumbnail.png'}"
				   pzip:target="{concat('ppt/media/thumbnail.png','')}"/>
    <xsl:if test="@xlink:href">
      <pzip:copy pzip:source="{@xlink:href}"
       pzip:target="{concat('ppt/media/',substring-after(@xlink:href,'/'))}"/>
    </xsl:if>
    <xsl:if test="./draw:image/@xlink:href">
      <pzip:copy pzip:source="{./draw:image/@xlink:href}"
         pzip:target="{concat('ppt/media/',substring-after(./draw:image/@xlink:href,'/'))}"/>
    </xsl:if>
    <xsl:variable name="varMediaFilePath">
      <xsl:if test="@xlink:href [ contains(.,'../')]">
        <xsl:value-of select="@xlink:href" />
      </xsl:if>
      <xsl:if test="not(@xlink:href [ contains(.,'../')])">
        <xsl:value-of select="substring-after(@xlink:href,'/')" />
      </xsl:if>
    </xsl:variable>
    <xsl:if test="@xlink:href[ contains(.,'wav')] or @xlink:href[ contains(.,'WAV')]">
      <pzip:import pzip:source="{$varMediaFilePath}" pzip:target="{concat('ppt/media/',$audioSerialNo,'.wav')}" />
    </xsl:if>
    <p:pic>
      <p:nvPicPr>
        <p:cNvPr id="4" name="sound1">
          <a:hlinkClick r:id="" action="ppaction://media"/>
        </p:cNvPr>
        <p:cNvPicPr>
          <a:picLocks noRot="1" noChangeAspect="1"/>
        </p:cNvPicPr>
        <p:nvPr>
          <!--@xlink:href[ contains(.,'asf')] i am using this extension in Video file bcoz it is avilable in movies and sounds-->
          <xsl:if test="@xlink:href[ contains(.,'mp3')] or @xlink:href[ contains(.,'m3u')] or @xlink:href[ contains(.,'wma')] or 
          @xlink:href[ contains(.,'wax')] or @xlink:href[ contains(.,'aif')] or @xlink:href[ contains(.,'aifc')] or
          @xlink:href[ contains(.,'aiff')] or @xlink:href[ contains(.,'au')] or @xlink:href[ contains(.,'snd')] or
          @xlink:href[ contains(.,'mid')] or @xlink:href[ contains(.,'midi')] or @xlink:href[ contains(.,'rmi')]">
            <a:audioFile>
              <xsl:attribute name ="r:link">
                <xsl:value-of select ="$audioSerialNo"/>
              </xsl:attribute>
            </a:audioFile>
          </xsl:if>
          <xsl:if test="not(@xlink:href[ contains(.,'mp3')] or @xlink:href[ contains(.,'m3u')] or @xlink:href[ contains(.,'wma')] or 
          @xlink:href[ contains(.,'wax')] or @xlink:href[ contains(.,'aif')] or @xlink:href[ contains(.,'aifc')] or
          @xlink:href[ contains(.,'aiff')] or @xlink:href[ contains(.,'au')] or @xlink:href[ contains(.,'snd')] or
          @xlink:href[ contains(.,'mid')] or @xlink:href[ contains(.,'midi')] or @xlink:href[ contains(.,'rmi')])">
            <xsl:if test="not(@xlink:href[ contains(.,'wav')] or @xlink:href[ contains(.,'WAV')])">
              <a:videoFile>
                <xsl:attribute name ="r:link">
                  <xsl:value-of select ="$audioSerialNo"/>
                </xsl:attribute>
              </a:videoFile>
            </xsl:if>
          </xsl:if>
          <xsl:if test="@xlink:href[ contains(.,'wav')] or @xlink:href[ contains(.,'WAV')]">
            <a:wavAudioFile>
              <xsl:attribute name ="r:embed">
                <xsl:value-of select ="$audioSerialNo"/>
              </xsl:attribute>
              <xsl:attribute name ="name">
                <xsl:value-of select ="'sound1'"/>
              </xsl:attribute>
            </a:wavAudioFile>
          </xsl:if>
        </p:nvPr>
      </p:nvPicPr>
      <p:blipFill>
        <a:blip>
          <xsl:attribute name="r:embed">
            <xsl:value-of select="$imageSerialNo"/>
          </xsl:attribute>
        </a:blip>
        <a:stretch>
          <a:fillRect/>
        </a:stretch>
      </p:blipFill>
      <p:spPr>
        <a:xfrm>
          <xsl:variable name ="angle">
            <xsl:value-of select="substring-after(substring-before(substring-before(parent::node()/@draw:transform,'translate'),')'),'(')" />
          </xsl:variable>
          <xsl:variable name ="x2">
            <xsl:value-of select="substring-before(substring-before(substring-before(substring-after(substring-after(parent::node()/@draw:transform,'translate'),'('),')'),' '),'cm')" />
          </xsl:variable>
          <xsl:variable name ="y2">
            <xsl:value-of select="substring-before(substring-after(substring-before(substring-after(substring-after(parent::node()/@draw:transform,'translate'),'('),')'),' '),'cm')" />
          </xsl:variable>
          <xsl:if test="@draw:transform">
            <xsl:attribute name ="rot">
              <xsl:value-of select ="concat('draw-transform:ROT:',substring-before(parent::node()/@svg:width,'cm'), ':',
																   substring-before(parent::node()/@svg:height,'cm'), ':', 
																   $x2, ':', 
                                   $y2, ':', 
																   $angle)"/>
            </xsl:attribute>
          </xsl:if>
          <a:off >
            <xsl:if test="not(@draw:transform)">
              <xsl:attribute name ="x">
                <!--<xsl:value-of select ="@svg:x"/>-->
                <xsl:call-template name ="convertToPoints">
                  <xsl:with-param name ="unit" select ="'cm'"/>
                  <xsl:with-param name ="length" select ="parent::node()/@svg:x"/>
                </xsl:call-template>
              </xsl:attribute>
              <xsl:attribute name ="y">
                <!--<xsl:value-of select ="@svg:y"/>-->
                <xsl:call-template name ="convertToPoints">
                  <xsl:with-param name ="unit" select ="'cm'"/>
                  <xsl:with-param name ="length" select ="parent::node()/@svg:y"/>
                </xsl:call-template>
              </xsl:attribute>
            </xsl:if >
            <xsl:if test="@draw:transform">
              <xsl:attribute name ="x">
                <xsl:value-of select ="concat('draw-transform:X:',substring-before(parent::node()/@svg:width,'cm'), ':',
																   substring-before(parent::node()/@svg:height,'cm'), ':', 
																   $x2, ':', $y2, ':', $angle)"/>
              </xsl:attribute>
              <xsl:attribute name ="y">
                <xsl:value-of select ="concat('draw-transform:Y:',substring-before(parent::node()/@svg:width,'cm'), ':',
																   substring-before(parent::node()/@svg:height,'cm'), ':', 
																   $x2, ':',$y2, ':', $angle)"/>
              </xsl:attribute>
            </xsl:if>
          </a:off>
          <a:ext>
            <xsl:attribute name ="cx">
              <!--<xsl:value-of select ="@svg:width"/>-->
              <xsl:call-template name ="convertToPoints">
                <xsl:with-param name ="unit" select ="'cm'"/>
                <xsl:with-param name ="length" select ="parent::node()/@svg:width"/>
              </xsl:call-template>
            </xsl:attribute>
            <xsl:attribute name ="cy">
              <!--<xsl:value-of select ="@svg:height"/>-->
              <xsl:call-template name ="convertToPoints">
                <xsl:with-param name ="unit" select ="'cm'"/>
                <xsl:with-param name ="length" select ="parent::node()/@svg:height"/>
              </xsl:call-template>
            </xsl:attribute>
          </a:ext >
        </a:xfrm>
        <a:prstGeom prst="rect">
          <a:avLst />
        </a:prstGeom>
      </p:spPr>
    </p:pic>
  </xsl:template>
  <xsl:template name ="pictureBorderLine">
    <xsl:param name="FileName"/>
    <xsl:param name ="grProp">
      <xsl:value-of select ="parent::node()/@draw:style-name"/>
    </xsl:param>
    <xsl:for-each select ="document($FileName)//style:style[@style:name=$grProp]/style:graphic-properties">
        <xsl:call-template name ="getLineStyle"/>
    </xsl:for-each>    

  </xsl:template>
  <xsl:template name ="tmpInsertBackImage">
    <xsl:param name ="FileName" />
    <xsl:param name ="imageName" />
    <xsl:variable name ="imageSerialNo">
      <xsl:value-of select ="concat($FileName,'BackImg')"/>
    </xsl:variable>
    <xsl:for-each select="document('styles.xml')/office:document-styles/office:styles/draw:fill-image[@draw:name=$imageName]">
      <xsl:variable name="Source">
        <xsl:choose>
          <xsl:when test="contains(@xlink:href,'#')">
            <xsl:value-of select="substring-after(@xlink:href,'#')"/>
          </xsl:when>
          <xsl:otherwise>
            <xsl:value-of select="@xlink:href"/>
          </xsl:otherwise>
        </xsl:choose>

      </xsl:variable>
      <pzip:copy pzip:source="{$Source}"
          pzip:target="{concat('ppt/media/',substring-after(@xlink:href,'/'))}"/>
    </xsl:for-each>

    <a:blip>
      <xsl:attribute name ="r:embed">
        <xsl:value-of select ="$imageSerialNo"/>
      </xsl:attribute>
      <a:lum/>
    </a:blip >

  </xsl:template>

</xsl:stylesheet>