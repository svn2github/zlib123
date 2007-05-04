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
  xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
  xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
  xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0"
  xmlns:number="urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0"
  xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0"
  exclude-result-prefixes="number style fo">
  
  <!-- template to insert number formats -->
  
  <xsl:template match="number:number-style" mode="numFormat">
    <xsl:param name="numId"/>
    <numFmt numFmtId="{$numId}">
      <xsl:attribute name="formatCode">
        <xsl:choose>
          
          <!-- when negative number is red, positive and negative number are formatted separately --> 
          <xsl:when test="style:text-properties/@fo:color">
            <xsl:variable name="formatPositive">
              <xsl:call-template name="GetFormatCode">
                <xsl:with-param name="sign">positive</xsl:with-param>
              </xsl:call-template>
            </xsl:variable>
            <xsl:variable name="formatNegative">
              <xsl:call-template name="GetFormatCode">
                <xsl:with-param name="sign">negative</xsl:with-param>
              </xsl:call-template>
            </xsl:variable>
            <xsl:value-of select="concat($formatPositive,concat(';',$formatNegative))"/>
          </xsl:when>
          
          <xsl:otherwise>
            <xsl:call-template name="GetFormatCode"/>
          </xsl:otherwise>
        </xsl:choose>
      </xsl:attribute>
    </numFmt>
    <xsl:choose>
      <xsl:when test="following-sibling::number:number-style">
        <xsl:apply-templates select="following-sibling::number:number-style[1]" mode="numFormat">
          <xsl:with-param name="numId">
            <xsl:value-of select="$numId + 1"/>
          </xsl:with-param>
        </xsl:apply-templates>
      </xsl:when>
    </xsl:choose>
  </xsl:template>
  
  <!-- template to insert percentage formats -->
  
  <xsl:template match="number:percentage-style" mode="numFormat">
    <xsl:param name="numId"/>
    <numFmt numFmtId="{$numId}">
      <xsl:attribute name="formatCode">
        <xsl:choose>
          
          <!-- when negative number is red, positive and negative number are formatted separately --> 
          <xsl:when test="style:text-properties/@fo:color">
            <xsl:variable name="formatPositive">
              <xsl:call-template name="GetFormatCode">
                <xsl:with-param name="sign">positive</xsl:with-param>
              </xsl:call-template>
            </xsl:variable>
            <xsl:variable name="formatNegative">
              <xsl:call-template name="GetFormatCode">
                <xsl:with-param name="sign">negative</xsl:with-param>
              </xsl:call-template>
            </xsl:variable>
            <xsl:value-of select="concat(concat($formatPositive,'%'),concat(concat(';',$formatNegative),'%'))"/>
          </xsl:when>
          
          <xsl:otherwise>
            <xsl:variable name="format">
            <xsl:call-template name="GetFormatCode"/>
            </xsl:variable>
            <xsl:value-of select="concat($format,'%')"/>
          </xsl:otherwise>
        </xsl:choose>
      </xsl:attribute>
    </numFmt>
    <xsl:choose>
      <xsl:when test="following-sibling::number:percentage-style">
        <xsl:apply-templates select="following-sibling::number:percentage-style[1]" mode="numFormat">
          <xsl:with-param name="numId">
            <xsl:value-of select="$numId + 1"/>
          </xsl:with-param>
        </xsl:apply-templates>
      </xsl:when>
    </xsl:choose>
  </xsl:template>
  
  <!-- template to insert currency formats -->
  
  <xsl:template match="number:currency-style" mode="numFormat">
    <xsl:param name="numId"/>
    
    <xsl:variable name="currencySymbol">
      <xsl:call-template name="ConvertValueSymbol">
        <xsl:with-param name="symbol">
          <xsl:value-of select="number:currency-symbol"/>
        </xsl:with-param>
      </xsl:call-template>
    </xsl:variable>
    
    <numFmt numFmtId="{$numId}">
      <xsl:attribute name="formatCode">
        <xsl:choose>
          
          <!-- when negative number is red, positive and negative number are formatted separately --> 
          <xsl:when test="style:text-properties/@fo:color">
            <xsl:variable name="formatPositive">
              <xsl:call-template name="GetFormatCode">
                <xsl:with-param name="sign">positive</xsl:with-param>
                <xsl:with-param name="currencySymbol" select="$currencySymbol"/>
              </xsl:call-template>
            </xsl:variable>
            <xsl:variable name="formatNegative">
              <xsl:call-template name="GetFormatCode">
                <xsl:with-param name="sign">negative</xsl:with-param>
                <xsl:with-param name="currencySymbol" select="$currencySymbol"/>
              </xsl:call-template>
            </xsl:variable>
            <xsl:value-of select="concat($formatPositive,concat(';',$formatNegative))"/>
          </xsl:when>
          
          <xsl:otherwise>
            <xsl:call-template name="GetFormatCode">
              <xsl:with-param name="currencySymbol" select="$currencySymbol"/>
            </xsl:call-template>
          </xsl:otherwise>
        </xsl:choose>
      </xsl:attribute>
    </numFmt>
    <xsl:choose>
      <xsl:when test="following-sibling::number:currency-style">
        <xsl:apply-templates select="following-sibling::number:currency-style[1]" mode="numFormat">
          <xsl:with-param name="numId">
            <xsl:value-of select="$numId + 1"/>
          </xsl:with-param>
          <xsl:with-param name="currencySymbol" select="$currencySymbol"/>
        </xsl:apply-templates>
      </xsl:when>
    </xsl:choose>
  </xsl:template>
  
  <!-- get number code format -->
  <xsl:template name="GetFormatCode">
    <xsl:param name="sign"/>
    <xsl:param name="currencySymbol"/>
    <xsl:variable name="value">
      <xsl:choose>
        
        <!-- add leading zeros if min-integer-digits > 0 -->
        <xsl:when test="number:number/@number:min-integer-digits and number:number/@number:min-integer-digits &gt; 0">
          <xsl:call-template name="AddLeadingZeros">
            <xsl:with-param name="num">
              <xsl:value-of select="number:number/@number:min-integer-digits"/>
            </xsl:with-param>
            <xsl:with-param name="val"/>
          </xsl:call-template>
        </xsl:when>
        
        <xsl:otherwise>#</xsl:otherwise>
      </xsl:choose>
    </xsl:variable>
    <xsl:variable name="endValue">
      <xsl:choose>
        
        <!-- add decimal places -->
        <xsl:when test="number:number/@number:decimal-places &gt; 0">
          <xsl:call-template name="AddDecimalPlaces">
            <xsl:with-param name="value">
              <xsl:choose>
                
                <!-- add grouping -->
                <xsl:when test="number:number/@number:grouping and number:number/@number:grouping = 'true'">
                  <xsl:call-template name="AddGrouping">
                    <xsl:with-param name="value">
                      <xsl:value-of select="concat($value,'.')"/>
                    </xsl:with-param>
                    <xsl:with-param name="numDigits">
                      <xsl:choose>
                        <xsl:when test="number:number/@number:min-integer-digits">
                          <xsl:value-of select="3 - number:number/@number:min-integer-digits"/>
                        </xsl:when>
                        <xsl:otherwise>2</xsl:otherwise>
                      </xsl:choose>
                    </xsl:with-param>
                  </xsl:call-template>
                </xsl:when>
                
                <xsl:otherwise>
                  <xsl:value-of select="concat($value,'.')"/>
                </xsl:otherwise>
              </xsl:choose>
            </xsl:with-param> 
            <xsl:with-param name="num"> 
              <xsl:value-of select="number:number/@number:decimal-places"/>   
            </xsl:with-param>
          </xsl:call-template>
        </xsl:when>
        
        <xsl:otherwise>
          <xsl:choose>
            
            <!-- add grouping -->
            <xsl:when test="number:number/@number:grouping and number:number/@number:grouping = 'true'">
              <xsl:call-template name="AddGrouping">
                <xsl:with-param name="value">
                  <xsl:value-of select="$value"/>
                </xsl:with-param>
                <xsl:with-param name="numDigits">
                  <xsl:choose>
                    <xsl:when test="number:number/@number:min-integer-digits">
                      <xsl:value-of select="3 - number:number/@number:min-integer-digits"/>
                    </xsl:when>
                    <xsl:otherwise>2</xsl:otherwise>
                  </xsl:choose>
                </xsl:with-param>
              </xsl:call-template>
            </xsl:when>
            
            <xsl:otherwise>
              <xsl:choose>
                <xsl:when test="number:number/@number:display-factor and number:number/@number:display-factor!=''">
                  <xsl:call-template name="UseThousandDisplayFactor">
                    <xsl:with-param name="displayFactor">
                      <xsl:value-of select="number:number/@number:display-factor"/>
                    </xsl:with-param>
                    <xsl:with-param name="value">
                      <xsl:value-of select="$value"/>
                    </xsl:with-param>
                  </xsl:call-template>
                </xsl:when>
                <xsl:otherwise>
                  <xsl:value-of select="$value"/>
                </xsl:otherwise>
              </xsl:choose>
            </xsl:otherwise>
          </xsl:choose>
        </xsl:otherwise>
      </xsl:choose>
    </xsl:variable>
    
    <!-- add color to negative number formatting when necessary -->
    <xsl:variable name="finalValue">
      <xsl:choose>
        <xsl:when test="$sign = 'negative' and style:text-properties/@fo:color">
          <xsl:value-of select="concat('[Red]-',$endValue)"/>
        </xsl:when>
        <xsl:otherwise>
          <xsl:value-of select="$endValue"/>
        </xsl:otherwise>
      </xsl:choose>
    </xsl:variable>
    
    <!-- add currency symbol -->
    <xsl:variable name="valueWithCurrency">
    <xsl:choose>
      <xsl:when test="$currencySymbol and $currencySymbol!='' and number:number/following-sibling::number:currency-symbol">
        
        <!-- add space before currency symbol -->
        <xsl:variable name="currencyFormat">
          <xsl:choose>
            <xsl:when test="number:number/following-sibling::number:text = ' '">
              <xsl:value-of select="concat('\ ',$currencySymbol)"/>
            </xsl:when>
            <xsl:otherwise>
              <xsl:value-of select="$currencySymbol"/>
            </xsl:otherwise>
          </xsl:choose>
        </xsl:variable>
        
        <xsl:value-of select="concat($finalValue,$currencyFormat)"/>
      </xsl:when>
      <xsl:when test="$currencySymbol and $currencySymbol!='' and number:number/preceding-sibling::number:currency-symbol">
        
        <!-- add space after currency symbol -->
        <xsl:variable name="currencyFormat2">
          <xsl:choose>
            <xsl:when test="number:number/preceding-sibling::number:text = ' '">
              <xsl:value-of select="concat($currencySymbol,'\ ')"/>
            </xsl:when>
            <xsl:otherwise>
              <xsl:value-of select="$currencySymbol"/>
            </xsl:otherwise>
          </xsl:choose>
        </xsl:variable>
        <xsl:value-of select="concat($currencyFormat2,$finalValue)"/>
      </xsl:when>
      <xsl:otherwise>
        <xsl:value-of select="$finalValue"/>
      </xsl:otherwise>
    </xsl:choose>
    </xsl:variable>
    <xsl:variable name="text">
      <xsl:value-of select="number:text[not(preceding-sibling::node())]"/>
    </xsl:variable>
    <xsl:choose>
      
      <!-- handle text in number format -->
      <xsl:when test="$text and $text!='' and $text != ' ' and $text != '-'">
        <xsl:value-of select="concat(concat(concat('&quot;',$text),'&quot;'),$valueWithCurrency)"/>
      </xsl:when>
      
      <xsl:otherwise>
        <xsl:value-of select="$valueWithCurrency"/>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>
  
  <!-- template to add leading zeros -->
  
  <xsl:template name="AddLeadingZeros">
    <xsl:param name="num"/>
    <xsl:param name="val"/>
    <xsl:choose>
      <xsl:when test="$num &gt; 0">
        <xsl:call-template name="AddLeadingZeros">
          <xsl:with-param name="num">
            <xsl:value-of select="$num - 1"/>
          </xsl:with-param>
          <xsl:with-param name="val">
            <xsl:value-of select="concat('0',$val)"/>
          </xsl:with-param>
        </xsl:call-template>
      </xsl:when>
      <xsl:otherwise>
        <xsl:value-of select="$val"/>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>
  
  <!-- template to add grouping -->
  
  <xsl:template name="AddGrouping">
    <xsl:param name="numDigits"/>
    <xsl:param name="value"/>
    <xsl:choose>
      <xsl:when test="$numDigits &gt; 0">
        <xsl:call-template name="AddGrouping">
          <xsl:with-param name="numDigits">
            <xsl:value-of select="$numDigits -1"/>
          </xsl:with-param>
          <xsl:with-param name="value">
            <xsl:value-of select="concat('#',$value)"/>
          </xsl:with-param>
        </xsl:call-template>
      </xsl:when>
      <xsl:otherwise>
        <xsl:value-of select="concat('#,',$value)"/>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>
  
  <!-- template to add decimal places -->
  
  <xsl:template name="AddDecimalPlaces">
    <xsl:param name="value"/>
    <xsl:param name="num"/>
    <xsl:choose>
      <xsl:when test="$num &gt; 0">
        <xsl:call-template name="AddDecimalPlaces">
          <xsl:with-param name="value">
            <xsl:value-of select="concat($value,'0')"/>
          </xsl:with-param>
          <xsl:with-param name="num">
            <xsl:value-of select="$num - 1"/>
          </xsl:with-param>
        </xsl:call-template>
      </xsl:when>
      <xsl:otherwise>
        <xsl:value-of select="$value"/>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>
  
  <!-- template to divide number by display factor -->
  
  <xsl:template name="UseThousandDisplayFactor">
    <xsl:param name="displayFactor"/>
    <xsl:param name="value"/>
    <xsl:choose>
      <xsl:when test="string-length($displayFactor) &gt; 1">
        <xsl:call-template name="UseThousandDisplayFactor">
          <xsl:with-param name="displayFactor">
            <xsl:value-of select="number($displayFactor) div 1000"/>
          </xsl:with-param>
          <xsl:with-param name="value">
            <xsl:value-of select="concat($value,',')"/>
          </xsl:with-param>
        </xsl:call-template>
      </xsl:when>
      <xsl:otherwise>
        <xsl:value-of select="$value"/>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>
  
  <!-- template to get number format id -->
  
  <xsl:template name="GetNumFmtId">
    <xsl:param name="numStyle"/>
    <xsl:param name="numStyleCount"/>
    <xsl:param name="styleNumStyleCount"/>
    <xsl:param name="percentStyleCount"/>
    <xsl:param name="stylePercentStyleCount"/>
    <xsl:param name="currencyStyleCount"/>
    <xsl:param name="styleCurrencyStyleCount"/>
    <xsl:param name="dateStyleCount"/>
    <xsl:choose>
      <xsl:when test="key('number',$numStyle)">
        <xsl:for-each select="key('number',$numStyle)">
          <xsl:value-of select="count(preceding-sibling::number:number-style)+1"/>
        </xsl:for-each>
      </xsl:when>
      <xsl:when test="document('styles.xml')/office:document-styles/office:styles/number:number-style[@style:name=$numStyle]">
        <xsl:for-each select="document('styles.xml')/office:document-styles/office:styles/number:number-style[@style:name=$numStyle]">
          <xsl:value-of select="count(preceding-sibling::number:number-style)+1+$numStyleCount"/>
        </xsl:for-each>
      </xsl:when>
      <xsl:when test="key('percentage',$numStyle)">
        <xsl:for-each select="key('percentage',$numStyle)">
          <xsl:value-of select="count(preceding-sibling::number:percentage-style)+1+$numStyleCount+$styleNumStyleCount"/>
        </xsl:for-each>
      </xsl:when>
      <xsl:when test="document('styles.xml')/office:document-styles/office:styles/number:percentage-style[@style:name=$numStyle]">
        <xsl:for-each select="document('styles.xml')/office:document-styles/office:styles/number:percentage-style[@style:name=$numStyle]">
          <xsl:value-of select="count(preceding-sibling::number:percentage-style)+1+$numStyleCount+$styleNumStyleCount+$percentStyleCount"/>
        </xsl:for-each>
      </xsl:when>
      <xsl:when test="key('currency',$numStyle)">
        <xsl:for-each select="key('currency',$numStyle)">
          <xsl:value-of select="count(preceding-sibling::number:currency-style)+1+$numStyleCount+$styleNumStyleCount+$percentStyleCount+$stylePercentStyleCount"/>
        </xsl:for-each>
      </xsl:when>
      <xsl:when test="document('styles.xml')/office:document-styles/office:styles/number:currency-style[@style:name=$numStyle]">
        <xsl:for-each select="document('styles.xml')/office:document-styles/office:styles/number:currency-style[@style:name=$numStyle]">
          <xsl:value-of select="count(preceding-sibling::number:currency-style)+1+$numStyleCount+$styleNumStyleCount+$percentStyleCount+$stylePercentStyleCount+$currencyStyleCount"/>
        </xsl:for-each>
      </xsl:when>
      <xsl:when test="key('date',$numStyle)">
        <xsl:for-each select="key('date',$numStyle)">
          <xsl:value-of select="count(preceding-sibling::number:date-style)+1+$numStyleCount+$styleNumStyleCount+$percentStyleCount+$stylePercentStyleCount+$currencyStyleCount+$styleCurrencyStyleCount"/>
        </xsl:for-each>
      </xsl:when>
      <xsl:when test="document('styles.xml')/office:document-styles/office:styles/number:date-style[@style:name=$numStyle]">
        <xsl:for-each select="document('styles.xml')/office:document-styles/office:styles/number:date-style[@style:name=$numStyle]">
          <xsl:value-of select="count(preceding-sibling::number:date-style)+1+$numStyleCount+$styleNumStyleCount+$percentStyleCount+$stylePercentStyleCount+$currencyStyleCount+$styleCurrencyStyleCount+$dateStyleCount"/>
        </xsl:for-each>
        </xsl:when>
      <xsl:otherwise>0</xsl:otherwise>
    </xsl:choose>
  </xsl:template>
  
  <xsl:template name="ConvertValueSymbol">
    <xsl:param name="symbol"/>
    <xsl:choose>
      <xsl:when test="$symbol = '$'">[$$-409]</xsl:when>
      <xsl:when test="$symbol = 'USD'">[$USD]</xsl:when>
      <xsl:when test="$symbol = '£'">[$£-809]</xsl:when>
      <xsl:when test="$symbol = 'GBP'">[$GBP]</xsl:when>
      <xsl:when test="$symbol = 'zł'">&quot;zł&quot;</xsl:when>
      <xsl:when test="$symbol = 'PLN'">[$PLN]</xsl:when>
      <xsl:when test="$symbol = '€'">
        <xsl:choose>
          <xsl:when test="following-sibling::number:number">[$€-2]</xsl:when>
          <xsl:otherwise>[$€-1]</xsl:otherwise>
        </xsl:choose>
      </xsl:when>
      <xsl:when test="$symbol = 'EUR'">[$EUR]</xsl:when>
    </xsl:choose>
  </xsl:template>
  
</xsl:stylesheet>
