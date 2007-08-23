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
<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
  xmlns:v="urn:schemas-microsoft-com:vml" xmlns:xlink="http://www.w3.org/1999/xlink"
  xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0"
  xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0"
  xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
  xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0"
  xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"
  xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0"
  xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
  xmlns:e="http://schemas.openxmlformats.org/spreadsheetml/2006/main" exclude-result-prefixes="e r">

  <xsl:template name="InsertText">
    <xsl:param name="position"/>
    <xsl:param name="colNum"/>
    <xsl:param name="rowNum"/>
    <xsl:param name="sheetNr"/>

    <xsl:choose>
      <xsl:when test="@t='s' ">
        <xsl:attribute name="office:value-type">
          <xsl:text>string</xsl:text>
        </xsl:attribute>
        <xsl:variable name="id">
          <xsl:value-of select="e:v"/>
        </xsl:variable>
        <text:p>
          <xsl:choose>
            <xsl:when test="key('ref',@r)">
              <text:a>
                <xsl:attribute name="xlink:href">
                  <xsl:variable name="target">
                    <!-- path to sheet file from xl/ catalog (i.e. $sheet = worksheets/sheet1.xml) -->
                    <xsl:for-each select="key('ref',@r)">
                      <xsl:call-template name="GetTarget">
                        <xsl:with-param name="id">
                          <xsl:value-of select="@r:id"/>
                        </xsl:with-param>
                        <xsl:with-param name="document">
                          <xsl:value-of select="concat('xl/worksheets/sheet', $sheetNr, '.xml')"/>
                        </xsl:with-param>
                      </xsl:call-template>
                    </xsl:for-each>
                  </xsl:variable>
                  <xsl:choose>
                    <!-- when hyperlink leads to a file in network -->
                    <xsl:when test="starts-with($target,'file:///\\')">
                      <xsl:value-of select="translate(substring-after($target,'file:///'),'\','/')"
                      />
                    </xsl:when>
                    <!--when hyperlink leads to www or mailto -->
                    <xsl:when test="contains($target,':')">
                      <xsl:value-of select="$target"/>
                    </xsl:when>
                    <!-- when hyperlink leads to another place in workbook -->
                    <xsl:when test="key('ref',@r)/@location != '' ">
                      <xsl:for-each select="key('ref',@r)">

                        <xsl:variable name="apos">
                          <xsl:text>&apos;</xsl:text>
                        </xsl:variable>

                        <xsl:variable name="sheetName">
                          <xsl:choose>
                            <xsl:when test="starts-with(@location,$apos)">
                              <xsl:value-of select="$apos"/>
                              <xsl:value-of
                                select="substring-before(substring-after(@location,$apos),$apos)"/>
                              <xsl:value-of select="$apos"/>
                            </xsl:when>
                            <xsl:otherwise>
                              <xsl:value-of select="substring-before(@location,'!')"/>
                            </xsl:otherwise>
                          </xsl:choose>
                        </xsl:variable>

                        <xsl:variable name="invalidChars">
                          <xsl:text>&apos;!,.+$-()</xsl:text>
                        </xsl:variable>

                        <xsl:variable name="checkedName">
                          <xsl:for-each
                            select="document('xl/workbook.xml')/e:workbook/e:sheets/e:sheet[@name = translate($sheetName,$apos,'')]">
                            <xsl:call-template name="CheckSheetName">
                              <xsl:with-param name="sheetNumber">
                                <xsl:for-each
                                  select="document('xl/workbook.xml')/e:workbook/e:sheets/e:sheet[@name = translate($sheetName,$apos,'')]">
                                  <xsl:value-of select="count(preceding-sibling::e:sheet) + 1"/>
                                </xsl:for-each>
                              </xsl:with-param>
                              <xsl:with-param name="name">
                                <xsl:value-of select="translate($sheetName,$invalidChars,'')"/>
                              </xsl:with-param>
                            </xsl:call-template>
                          </xsl:for-each>
                        </xsl:variable>

                        <xsl:text>#</xsl:text>
                        <xsl:value-of select="$checkedName"/>
                        <xsl:text>.</xsl:text>
                        <xsl:value-of select="substring-after(@location,concat($sheetName,'!'))"/>

                      </xsl:for-each>
                    </xsl:when>
                    <!--when hyperlink leads to a document -->
                    <xsl:otherwise>
                      <xsl:call-template name="Change20PercentToSpace">
                        <xsl:with-param name="string" select="$target"/>
                      </xsl:call-template>
                    </xsl:otherwise>
                  </xsl:choose>
                </xsl:attribute>

                <!-- a postprocessor puts here strings from sharedstrings -->
                <pxsi:v xmlns:pxsi="urn:cleverage:xmlns:post-processings:shared-strings">
                  <xsl:value-of select="e:v"/>
                </pxsi:v>

              </text:a>
            </xsl:when>

            <xsl:otherwise>
              <!-- a postprocessor puts here strings from sharedstrings -->
              <pxsi:v xmlns:pxsi="urn:cleverage:xmlns:post-processings:shared-strings">
                <xsl:value-of select="e:v"/>
              </pxsi:v>
            </xsl:otherwise>
          </xsl:choose>
        </text:p>
      </xsl:when>
      <xsl:when test="@t = 'e' ">
        <xsl:attribute name="office:value-type">
          <xsl:text>string</xsl:text>
        </xsl:attribute>
        <text:p>
          <xsl:value-of select="e:v"/>
        </text:p>
      </xsl:when>
      <xsl:when test="@t = 'str' ">
        <xsl:attribute name="office:value-type">
          <xsl:choose>
            <xsl:when test="number(e:v)">
              <xsl:text>float</xsl:text>
            </xsl:when>
            <xsl:otherwise>
              <xsl:text>string</xsl:text>
            </xsl:otherwise>
          </xsl:choose>
        </xsl:attribute>
        <text:p>
          <xsl:choose>
            <xsl:when test="number(e:v)">
              <xsl:call-template name="FormatNumber">
                <xsl:with-param name="value">
                  <xsl:value-of select="e:v"/>
                </xsl:with-param>
                <xsl:with-param name="numStyle">
                  <xsl:for-each select="document('xl/styles.xml')">
                    <xsl:value-of
                      select="key('numFmtId',key('Xf','')[position()=$position]/@numFmtId)/@formatCode"
                    />
                  </xsl:for-each>
                </xsl:with-param>
              </xsl:call-template>
            </xsl:when>
            <xsl:otherwise>
              <xsl:value-of select="e:v"/>
            </xsl:otherwise>
          </xsl:choose>
        </text:p>
      </xsl:when>
      <xsl:when test="@t = 'n'">
        <xsl:attribute name="office:value-type">
          <xsl:text>float</xsl:text>
        </xsl:attribute>
        <xsl:attribute name="office:value">
          <xsl:value-of select="e:v"/>
        </xsl:attribute>
        <text:p>
          <xsl:call-template name="FormatNumber">
            <xsl:with-param name="value">
              <xsl:value-of select="e:v"/>
            </xsl:with-param>
            <xsl:with-param name="numStyle">
              <xsl:for-each select="document('xl/styles.xml')">
                <xsl:value-of
                  select="key('numFmtId',key('Xf','')[position()=$position]/@numFmtId)/@formatCode"
                />
              </xsl:for-each>
            </xsl:with-param>
          </xsl:call-template>
        </text:p>
      </xsl:when>
      <xsl:otherwise>
        <xsl:variable name="numStyle">
          <xsl:for-each select="document('xl/styles.xml')">
            <xsl:value-of
              select="key('numFmtId',key('Xf','')[position()=$position]/@numFmtId)/@formatCode"/>
          </xsl:for-each>
        </xsl:variable>
        <xsl:variable name="numId">
          <xsl:for-each select="document('xl/styles.xml')">
            <xsl:value-of select="key('Xf','')[position()=$position]/@numFmtId"/>
          </xsl:for-each>
        </xsl:variable>
        <xsl:attribute name="office:value-type">
          <xsl:choose>
            <xsl:when
              test="contains($numStyle,'%') or ((not($numStyle) or $numStyle = '')  and ($numId = 9 or $numId = 10))">
              <xsl:text>percentage</xsl:text>
            </xsl:when>
            <xsl:when
              test="(contains($numStyle,'y') or (contains($numStyle,'m') and not(contains($numStyle,'h') or contains($numStyle,'s'))) or (contains($numStyle,'d') and not(contains($numStyle,'Red'))) or ($numId &gt; 13 and $numId &lt; 18) or $numId = 22)">
              <xsl:text>date</xsl:text>
            </xsl:when>

            <!--'and' at the end is for Latvian currency -->
            <xsl:when
              test="contains($numStyle,'h') or contains($numStyle,'s') and not(contains($numStyle,'[$Ls-426]'))">
              <xsl:text>time</xsl:text>
            </xsl:when>

            <xsl:when
              test="contains($numStyle,'zł') or contains($numStyle,'$') or contains($numStyle,'£') or contains($numStyle,'€')">
              <xsl:text>currency</xsl:text>
            </xsl:when>
            <xsl:when test="$numId = 18">
              <xsl:text>time</xsl:text>
            </xsl:when>
            <xsl:when test="$numId = 49">
              <xsl:text>string</xsl:text>
            </xsl:when>
            <xsl:otherwise>
              <xsl:text>float</xsl:text>
            </xsl:otherwise>
          </xsl:choose>
        </xsl:attribute>
        <xsl:choose>

          <xsl:when
            test="(contains($numStyle,'y') or (contains($numStyle,'m') and not(contains($numStyle,'h') or contains($numStyle,'s'))) or (contains($numStyle,'d') and not(contains($numStyle,'Red'))) or ($numId &gt; 13 and $numId &lt; 18) or $numId = 22)">
            <xsl:attribute name="office:date-value">
              <xsl:call-template name="NumberToDate">
                <xsl:with-param name="value">
                  <xsl:choose>
                    <xsl:when test="document('xl/workbook.xml')/e:workbook/e:workbookPr/@date1904 =1 ">
                      <xsl:value-of select="(e:v) + (1462)"/>
                    </xsl:when>
                    <xsl:otherwise>
                      <xsl:value-of select="e:v"/>
                    </xsl:otherwise>
                  </xsl:choose>
                </xsl:with-param>
              </xsl:call-template>
            </xsl:attribute>
          </xsl:when>

          <!--'and' at the end is for Latvian currency -->
          <xsl:when
            test="(contains($numStyle,'h') or contains($numStyle,'s') or $numId = 18 and not(contains($numStyle,'[$Ls-426]')))">
            <xsl:attribute name="office:time-value">
              <xsl:call-template name="NumberToTime">
                <xsl:with-param name="value">
                  <xsl:value-of select="e:v"/>
                </xsl:with-param>
              </xsl:call-template>
            </xsl:attribute>
          </xsl:when>
          <xsl:otherwise>
            <xsl:attribute name="office:value">
              <xsl:value-of select="e:v"/>
            </xsl:attribute>
          </xsl:otherwise>
        </xsl:choose>
        <text:p>
          <xsl:choose>

            <xsl:when
              test="(contains($numStyle,'y') or (contains($numStyle,'m') and not(contains($numStyle,'h') or contains($numStyle,'s'))) or (contains($numStyle,'d') and not(contains($numStyle,'Red'))) or ($numId &gt; 13 and $numId &lt; 18) or $numId = 22)">
              <xsl:call-template name="FormatDate">
                <xsl:with-param name="value">
                  <xsl:call-template name="NumberToDate">
                    <xsl:with-param name="value">
                      <xsl:value-of select="e:v"/>
                    </xsl:with-param>
                  </xsl:call-template>
                </xsl:with-param>
                <xsl:with-param name="format">
                  <xsl:choose>
                    <xsl:when test="contains($numStyle,']')">
                      <xsl:value-of select="substring-after($numStyle,']')"/>
                    </xsl:when>
                    <xsl:otherwise>
                      <xsl:value-of select="$numStyle"/>
                    </xsl:otherwise>
                  </xsl:choose>
                </xsl:with-param>
                <xsl:with-param name="numId">
                  <xsl:value-of select="$numId"/>
                </xsl:with-param>
                <xsl:with-param name="processedFormat">
                  <xsl:choose>
                    <xsl:when test="contains($numStyle,']')">
                      <xsl:value-of select="substring-after($numStyle,']')"/>
                    </xsl:when>
                    <xsl:otherwise>
                      <xsl:value-of select="$numStyle"/>
                    </xsl:otherwise>
                  </xsl:choose>
                </xsl:with-param>
                <xsl:with-param name="numValue">
                  <xsl:value-of select="e:v"/>
                </xsl:with-param>
              </xsl:call-template>
            </xsl:when>

            <!--'and' at the end is for Latvian currency -->
            <xsl:when
              test="contains($numStyle,'h') or contains($numStyle,'s') and not(contains($numStyle,'[$Ls-426]'))">
              <xsl:call-template name="FormatTime">
                <xsl:with-param name="value">
                  <xsl:call-template name="NumberToTime">
                    <xsl:with-param name="value">
                      <xsl:value-of select="e:v"/>
                    </xsl:with-param>
                  </xsl:call-template>
                </xsl:with-param>
                <xsl:with-param name="format">
                  <xsl:choose>
                    <xsl:when test="contains($numStyle,']') and not(contains($numStyle,'[h'))">
                      <xsl:value-of select="substring-after($numStyle,']')"/>
                    </xsl:when>
                    <xsl:otherwise>
                      <xsl:value-of select="$numStyle"/>
                    </xsl:otherwise>
                  </xsl:choose>
                </xsl:with-param>
                <xsl:with-param name="numId">
                  <xsl:value-of select="$numId"/>
                </xsl:with-param>
                <xsl:with-param name="processedFormat">
                  <xsl:choose>
                    <xsl:when test="contains($numStyle,']') and not(contains($numStyle,'[h'))">
                      <xsl:value-of select="substring-after($numStyle,']')"/>
                    </xsl:when>
                    <xsl:otherwise>
                      <xsl:value-of select="$numStyle"/>
                    </xsl:otherwise>
                  </xsl:choose>
                </xsl:with-param>
                <xsl:with-param name="numValue">
                  <xsl:value-of select="e:v"/>
                </xsl:with-param>
              </xsl:call-template>
            </xsl:when>

            <xsl:otherwise>
              <xsl:call-template name="FormatNumber">
                <xsl:with-param name="value">
                  <xsl:value-of select="e:v"/>
                </xsl:with-param>
                <xsl:with-param name="numStyle">
                  <xsl:value-of select="$numStyle"/>
                </xsl:with-param>
                <xsl:with-param name="numId">
                  <xsl:value-of select="$numId"/>
                </xsl:with-param>
              </xsl:call-template>
            </xsl:otherwise>
          </xsl:choose>
        </text:p>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>

  <!-- change  '%20' to space  after conversion-->
  <xsl:template name="Change20PercentToSpace">
    <xsl:param name="string"/>

    <xsl:choose>
      <xsl:when test="contains($string,'%20')">
        <xsl:choose>
          <xsl:when test="substring-before($string,'%20') =''">
            <xsl:call-template name="Change20PercentToSpace">
              <xsl:with-param name="string" select="concat(' ',substring-after($string,'%20'))"/>
            </xsl:call-template>
          </xsl:when>
          <xsl:when test="substring-before($string,'%20') !=''">
            <xsl:call-template name="Change20PercentToSpace">
              <xsl:with-param name="string"
                select="concat(substring-before($string,'%20'),' ',substring-after($string,'%20'))"
              />
            </xsl:call-template>
          </xsl:when>
        </xsl:choose>
      </xsl:when>

      <!--xsl:when test="contains($slash,'..\..\..\')">
          <xsl:choose>
          <xsl:when test="substring-before($slash,'..\..\..\') =''">
          <xsl:call-template name="Change20PercentToSpace">
          <xsl:with-param name="slash"
          select="concat('../../../',substring-after($slash,'..\..\..\'))"/>
          </xsl:call-template>
          </xsl:when>
          </xsl:choose>
          </xsl:when-->

      <xsl:when test="contains($string,'..\..\')">
        <xsl:choose>
          <xsl:when test="substring-before($string,'..\..\') =''">

            <xsl:call-template name="Change20PercentToSpace">
              <xsl:with-param name="string"
                select="concat('../../../',substring-after($string,'..\..\'))"/>
            </xsl:call-template>
          </xsl:when>
        </xsl:choose>
      </xsl:when>

      <xsl:when test="contains($string,'..\')">
        <xsl:choose>
          <xsl:when test="substring-before($string,'..\') =''">

            <xsl:call-template name="Change20PercentToSpace">
              <xsl:with-param name="string" select="concat('../../',substring-after($string,'..\'))"
              />
            </xsl:call-template>
          </xsl:when>
        </xsl:choose>
      </xsl:when>


      <xsl:when test="not(contains($string,'../')) and not(contains($string,'..\..\'))">

        <xsl:value-of select="concat('../',$string)">
          <!--xsl:call-template name="Change20PercentToSpace">
              <xsl:with-param name="slash" select="concat('..\',substring-after($string,''))"/>
              </xsl:call-template-->
        </xsl:value-of>

      </xsl:when>

      <xsl:otherwise>
        <xsl:value-of select="$string"/>
        <!--xsl:value-of select="translate($string,'\','/')"/-->
      </xsl:otherwise>
    </xsl:choose>

  </xsl:template>

  <!--  convert multiple white spaces  -->
  <xsl:template name="InsertWhiteSpaces">
    <xsl:param name="string" select="."/>
    <xsl:param name="length" select="string-length(.)"/>
    <!-- string which doesn't contain whitespaces-->
    <xsl:choose>
      <xsl:when test="not(contains($string,' '))">
        <xsl:value-of select="$string"/>
      </xsl:when>
      <!-- convert white spaces  -->
      <xsl:otherwise>
        <xsl:variable name="before">
          <xsl:value-of select="substring-before($string,' ')"/>
        </xsl:variable>
        <xsl:variable name="after">
          <xsl:call-template name="CutStartSpaces">
            <xsl:with-param name="cuted">
              <xsl:value-of select="substring-after($string,' ')"/>
            </xsl:with-param>
          </xsl:call-template>
        </xsl:variable>
        <xsl:if test="$before != '' ">
          <xsl:value-of select="concat($before,' ')"/>
        </xsl:if>
        <!--add remaining whitespaces as text:s if there are any-->
        <xsl:if test="string-length(concat($before,' ', $after)) &lt; $length ">
          <xsl:choose>
            <xsl:when test="($length - string-length(concat($before, $after))) = 1">
              <text:s/>
            </xsl:when>
            <xsl:otherwise>
              <text:s>
                <xsl:attribute name="text:c">
                  <xsl:choose>
                    <xsl:when test="$before = ''">
                      <xsl:value-of select="$length - string-length($after)"/>
                    </xsl:when>
                    <xsl:otherwise>
                      <xsl:value-of select="$length - string-length(concat($before,' ', $after))"/>
                    </xsl:otherwise>
                  </xsl:choose>
                </xsl:attribute>
              </text:s>
            </xsl:otherwise>
          </xsl:choose>
        </xsl:if>
        <!--repeat it for substring which has whitespaces-->
        <xsl:if test="contains($string,' ') and $length &gt; 0">
          <xsl:call-template name="InsertWhiteSpaces">
            <xsl:with-param name="string">
              <xsl:value-of select="$after"/>
            </xsl:with-param>
            <xsl:with-param name="length">
              <xsl:value-of select="string-length($after)"/>
            </xsl:with-param>
          </xsl:call-template>
        </xsl:if>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>

  <!--  cut start spaces -->
  <xsl:template name="CutStartSpaces">
    <xsl:param name="cuted"/>
    <xsl:choose>
      <xsl:when test="starts-with($cuted,' ')">
        <xsl:call-template name="CutStartSpaces">
          <xsl:with-param name="cuted">
            <xsl:value-of select="substring-after($cuted,' ')"/>
          </xsl:with-param>
        </xsl:call-template>
      </xsl:when>
      <xsl:otherwise>
        <xsl:value-of select="$cuted"/>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>

  <xsl:template name="CountContinuous">
    <xsl:param name="count" select="0"/>

    <xsl:variable name="carryOn">
      <xsl:choose>
        <xsl:when test="following-sibling::e:c[1]/@s and not(following-sibling::e:c[1]/e:v)">
          <xsl:variable name="position">
            <xsl:value-of select="following-sibling::e:c[1]/@s + 1"/>
          </xsl:variable>
          <xsl:variable name="horizontal">
            <xsl:for-each select="document('xl/styles.xml')">
              <xsl:value-of select="key('Xf', '')[position() = $position]/e:alignment/@horizontal"/>
            </xsl:for-each>
          </xsl:variable>

          <xsl:choose>
            <xsl:when test="$horizontal = 'centerContinuous' ">
              <xsl:text>true</xsl:text>
            </xsl:when>
            <xsl:otherwise>
              <xsl:text>false</xsl:text>
            </xsl:otherwise>
          </xsl:choose>

        </xsl:when>
        <xsl:otherwise>
          <xsl:text>false</xsl:text>
        </xsl:otherwise>
      </xsl:choose>
    </xsl:variable>

    <xsl:choose>
      <xsl:when test="$carryOn = 'true'">
        <xsl:for-each select="following-sibling::e:c[1]">
          <xsl:call-template name="CountContinuous">
            <xsl:with-param name="count" select="$count + 1"/>
          </xsl:call-template>
        </xsl:for-each>
      </xsl:when>
      <xsl:otherwise>
        <!-- number of following cells plus the starting one -->
        <xsl:value-of select="$count + 1"/>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>

</xsl:stylesheet>
