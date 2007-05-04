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
  xmlns:e="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
  xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0"
  xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0"
  xmlns:number="urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0"
  exclude-result-prefixes="e">
  
  <!-- insert  number format style -->
  
  <xsl:template match="e:numFmt" mode="automaticstyles">
    
    <xsl:choose>
      
      <!-- date style -->
      <xsl:when test="contains(@formatCode,'y') or contains(@formatCode,'m') or (contains(@formatCode,'d') and not(contains(@formatCode,'Red'))) or contains(@formatCode,'h') or contains(@formatCode,'s')">
        <number:date-style style:name="{generate-id(.)}">
        <xsl:call-template name="ProcessFormat">
          <xsl:with-param name="format">
            <xsl:choose>
              <xsl:when test="contains(@formatCode,']')">
                <xsl:value-of select="substring-after(@formatCode,']')"/>
              </xsl:when>
              <xsl:otherwise>
                <xsl:value-of select="@formatCode"/>
              </xsl:otherwise>
            </xsl:choose>
          </xsl:with-param>
          <xsl:with-param name="processedFormat">
            <xsl:choose>
              <xsl:when test="contains(@formatCode,']')">
                <xsl:value-of select="substring-after(@formatCode,']')"/>
              </xsl:when>
              <xsl:otherwise>
                <xsl:value-of select="@formatCode"/>
              </xsl:otherwise>
            </xsl:choose>
          </xsl:with-param>
        </xsl:call-template>
        </number:date-style>
      </xsl:when>
      
      <!-- when there are different formats for positive and negative numbers -->
      <xsl:when test="contains(@formatCode,';') and not(contains(substring-after(@formatCode,';'),';'))">
        <xsl:choose>
          
          <!-- currency style -->
          <xsl:when test="contains(substring-before(@formatCode,';'),'$') or contains(substring-before(@formatCode,';'),'zł') or contains(substring-before(@formatCode,';'),'€') or contains(substring-before(@formatCode,';'),'£')">
            <number:currency-style style:name="{concat(generate-id(.),'P0')}">
              <xsl:call-template name="InsertNumberFormatting">
                <xsl:with-param name="formatCode">
                  <xsl:value-of select="substring-before(@formatCode,';')"/>
                </xsl:with-param>
              </xsl:call-template>
            </number:currency-style>
            <number:currency-style style:name="{generate-id(.)}">
              <xsl:call-template name="InsertNumberFormatting">
                <xsl:with-param name="formatCode">
                  <xsl:value-of select="substring-after(@formatCode,';')"/>
                </xsl:with-param>
              </xsl:call-template>
              <style:map style:condition="value()&gt;=0" style:apply-style-name="{concat(generate-id(.),'P0')}"/>
            </number:currency-style>
          </xsl:when>
          
          <!--percentage style -->
          <xsl:when test="contains(substring-before(@formatCode,';'),'%')">
            <number:percentage-style style:name="{concat(generate-id(.),'P0')}">
              <xsl:call-template name="InsertNumberFormatting">
                <xsl:with-param name="formatCode">
                  <xsl:value-of select="substring-before(@formatCode,';')"/>
                </xsl:with-param>
              </xsl:call-template>
              <number:text>%</number:text>
            </number:percentage-style>
            <number:percentage-style style:name="{generate-id(.)}">
              <xsl:call-template name="InsertNumberFormatting">
                <xsl:with-param name="formatCode">
                  <xsl:value-of select="substring-after(@formatCode,';')"/>
                </xsl:with-param>
              </xsl:call-template>
              <number:text>%</number:text>
              <style:map style:condition="value()&gt;=0" style:apply-style-name="{concat(generate-id(.),'P0')}"/>
            </number:percentage-style>
          </xsl:when>
          
          <!-- number style -->
          <xsl:otherwise>
            <number:number-style style:name="{concat(generate-id(.),'P0')}">
              <xsl:call-template name="InsertNumberFormatting">
                <xsl:with-param name="formatCode">
                  <xsl:value-of select="substring-before(@formatCode,';')"/>
                </xsl:with-param>
              </xsl:call-template>
            </number:number-style>
            <number:number-style style:name="{generate-id(.)}">
              <xsl:call-template name="InsertNumberFormatting">
                <xsl:with-param name="formatCode">
                  <xsl:value-of select="substring-after(@formatCode,';')"/>
                </xsl:with-param>
              </xsl:call-template>
              <style:map style:condition="value()&gt;=0" style:apply-style-name="{concat(generate-id(.),'P0')}"/>
            </number:number-style>
          </xsl:otherwise>
          
        </xsl:choose>
      </xsl:when>
      
      <!-- when there are separate formats for positive numbers, negative numbers and zeros -->
      <xsl:when test="contains(@formatCode,';') and contains(substring-after(@formatCode,';'),';')">
        <xsl:choose>
          
          <!-- currency style -->
          <xsl:when test="contains(substring-before(@formatCode,';'),'$') or contains(substring-before(@formatCode,';'),'zł') or contains(substring-before(@formatCode,';'),'€') or contains(substring-before(@formatCode,';'),'£')">
            <number:currency-style style:name="{concat(generate-id(.),'P0')}">
              <xsl:call-template name="InsertNumberFormatting">
                <xsl:with-param name="formatCode">
                  <xsl:value-of select="substring-before(@formatCode,';')"/>
                </xsl:with-param>
              </xsl:call-template>
            </number:currency-style>
            <number:currency-style style:name="{concat(generate-id(.),'P1')}">
              <xsl:call-template name="InsertNumberFormatting">
                <xsl:with-param name="formatCode">
                  <xsl:value-of select="substring-before(substring-after(@formatCode,';'),';')"/>
                </xsl:with-param>
              </xsl:call-template>
            </number:currency-style>
            
            <xsl:choose>
              <xsl:when test="contains(substring-after(substring-after(@formatCode,';'),';'),';')">
                <number:currency-style style:name="{concat(generate-id(.),'P2')}">
                  <xsl:call-template name="InsertNumberFormatting">
                    <xsl:with-param name="formatCode">
                      <xsl:value-of select="substring-before(substring-after(substring-after(@formatCode,';'),';'),';')"/>
                    </xsl:with-param>
                  </xsl:call-template>
                </number:currency-style>
                <number:text-style style:name="{generate-id(.)}">
                  <xsl:variable name="text">
                    <xsl:value-of select="substring-after(substring-after(substring-after(@formatCode,';'),';'),';')"/>
                  </xsl:variable>
                  <xsl:choose>
                    
                    <!-- text content -->
                    <xsl:when test="contains($text,'@')">
                      <number:text>
                       <xsl:value-of select="translate(substring-before($text,'@'),'_-',' ')"/>
                      </number:text>
                      <number:text-content/>
                      <number:text>
                        <xsl:value-of select="translate(substring-after($text,'@'),'_-',' ')"/>
                      </number:text>
                    </xsl:when>
                    <xsl:otherwise>
                      <xsl:value-of select="translate($text,'_-',' ')"/>
                    </xsl:otherwise>
                    </xsl:choose>
                  <style:map style:condition="value()&gt;0" style:apply-style-name="{concat(generate-id(.),'P0')}"/>
                  <style:map style:condition="value()&lt;0" style:apply-style-name="{concat(generate-id(.),'P1')}"/>
                  <style:map style:condition="value()=0" style:apply-style-name="{concat(generate-id(.),'P2')}"/>
                </number:text-style>
              </xsl:when>
              <xsl:otherwise>
                <number:currency-style style:name="{generate-id(.)}">
                  <xsl:call-template name="InsertNumberFormatting">
                    <xsl:with-param name="formatCode">
                      <xsl:value-of select="substring-after(substring-after(@formatCode,';'),';')"/>
                    </xsl:with-param>
                  </xsl:call-template>
                  <style:map style:condition="value()&gt;0" style:apply-style-name="{concat(generate-id(.),'P0')}"/>
                  <style:map style:condition="value()&lt;0" style:apply-style-name="{concat(generate-id(.),'P1')}"/>
                </number:currency-style>
              </xsl:otherwise>
            </xsl:choose>
          </xsl:when>
          
          <!-- percentage style -->
          <xsl:when test="contains(substring-before(@formatCode,';'),'%')">
            <number:percentage-style style:name="{concat(generate-id(.),'P0')}">
              <xsl:call-template name="InsertNumberFormatting">
                <xsl:with-param name="formatCode">
                  <xsl:value-of select="substring-before(@formatCode,';')"/>
                </xsl:with-param>
              </xsl:call-template>
            </number:percentage-style>
            <number:percentage-style style:name="{concat(generate-id(.),'P1')}">
              <xsl:call-template name="InsertNumberFormatting">
                <xsl:with-param name="formatCode">
                  <xsl:value-of select="substring-before(substring-after(@formatCode,';'),';')"/>
                </xsl:with-param>
              </xsl:call-template>
            </number:percentage-style>
            
            <xsl:choose>
              <xsl:when test="contains(substring-after(substring-after(@formatCode,';'),';'),';')">
                <number:percentage-style style:name="{concat(generate-id(.),'P2')}">
                  <xsl:call-template name="InsertNumberFormatting">
                    <xsl:with-param name="formatCode">
                      <xsl:value-of select="substring-before(substring-after(substring-after(@formatCode,';'),';'),';')"/>
                    </xsl:with-param>
                  </xsl:call-template>
                </number:percentage-style>
                <number:text-style style:name="{generate-id(.)}">
                  <xsl:variable name="text">
                    <xsl:value-of select="substring-after(substring-after(substring-after(@formatCode,';'),';'),';')"/>
                  </xsl:variable>
                  <xsl:choose>
                    
                    <!-- text content -->
                    <xsl:when test="contains($text,'@')">
                      <number:text>
                        <xsl:value-of select="translate(substring-before($text,'@'),'_-',' ')"/>
                      </number:text>
                      <number:text-content/>
                      <number:text>
                        <xsl:value-of select="translate(substring-after($text,'@'),'_-',' ')"/>
                      </number:text>
                    </xsl:when>
                    <xsl:otherwise>
                      <xsl:value-of select="translate($text,'_-',' ')"/>
                    </xsl:otherwise>
                  </xsl:choose>
                  <style:map style:condition="value()&gt;0" style:apply-style-name="{concat(generate-id(.),'P0')}"/>
                  <style:map style:condition="value()&lt;0" style:apply-style-name="{concat(generate-id(.),'P1')}"/>
                  <style:map style:condition="value()=0" style:apply-style-name="{concat(generate-id(.),'P2')}"/>
                </number:text-style>
              </xsl:when>
              <xsl:otherwise>
                <number:percentage-style style:name="{generate-id(.)}">
                  <xsl:call-template name="InsertNumberFormatting">
                    <xsl:with-param name="formatCode">
                      <xsl:value-of select="substring-after(substring-after(@formatCode,';'),';')"/>
                    </xsl:with-param>
                  </xsl:call-template>
                  <style:map style:condition="value()&gt;0" style:apply-style-name="{concat(generate-id(.),'P0')}"/>
                  <style:map style:condition="value()&lt;0" style:apply-style-name="{concat(generate-id(.),'P1')}"/>
                </number:percentage-style>
              </xsl:otherwise>
            </xsl:choose>
          </xsl:when>
          
          <!-- number style -->
          <xsl:otherwise>
            <number:number-style style:name="{concat(generate-id(.),'P0')}">
              <xsl:call-template name="InsertNumberFormatting">
                <xsl:with-param name="formatCode">
                  <xsl:value-of select="substring-before(@formatCode,';')"/>
                </xsl:with-param>
              </xsl:call-template>
            </number:number-style>
            <number:number-style style:name="{concat(generate-id(.),'P1')}">
              <xsl:call-template name="InsertNumberFormatting">
                <xsl:with-param name="formatCode">
                  <xsl:value-of select="substring-before(substring-after(@formatCode,';'),';')"/>
                </xsl:with-param>
              </xsl:call-template>
            </number:number-style>
            
            <xsl:choose>
              <xsl:when test="contains(substring-after(substring-after(@formatCode,';'),';'),';')">
                <number:number-style style:name="{concat(generate-id(.),'P2')}">
                  <xsl:call-template name="InsertNumberFormatting">
                    <xsl:with-param name="formatCode">
                      <xsl:value-of select="substring-before(substring-after(substring-after(@formatCode,';'),';'),';')"/>
                    </xsl:with-param>
                  </xsl:call-template>
                </number:number-style>
                <number:text-style style:name="{generate-id(.)}">
                  <xsl:variable name="text">
                    <xsl:value-of select="substring-after(substring-after(substring-after(@formatCode,';'),';'),';')"/>
                  </xsl:variable>
                  <xsl:choose>
                    
                    <!-- text content -->
                    <xsl:when test="contains($text,'@')">
                      <number:text>
                        <xsl:value-of select="translate(substring-before($text,'@'),'_-',' ')"/>
                      </number:text>
                      <number:text-content/>
                      <number:text>
                        <xsl:value-of select="translate(substring-after($text,'@'),'_-',' ')"/>
                      </number:text>
                    </xsl:when>
                    <xsl:otherwise>
                      <xsl:value-of select="translate($text,'_-',' ')"/>
                    </xsl:otherwise>
                  </xsl:choose>
                  <style:map style:condition="value()&gt;0" style:apply-style-name="{concat(generate-id(.),'P0')}"/>
                  <style:map style:condition="value()&lt;0" style:apply-style-name="{concat(generate-id(.),'P1')}"/>
                  <style:map style:condition="value()=0" style:apply-style-name="{concat(generate-id(.),'P2')}"/>
                </number:text-style>
              </xsl:when>
              <xsl:otherwise>
                <number:number-style style:name="{generate-id(.)}">
                  <xsl:call-template name="InsertNumberFormatting">
                    <xsl:with-param name="formatCode">
                      <xsl:value-of select="substring-after(substring-after(@formatCode,';'),';')"/>
                    </xsl:with-param>
                  </xsl:call-template>
                  <style:map style:condition="value()&gt;0" style:apply-style-name="{concat(generate-id(.),'P0')}"/>
                  <style:map style:condition="value()&lt;0" style:apply-style-name="{concat(generate-id(.),'P1')}"/>
                </number:number-style>
              </xsl:otherwise>
            </xsl:choose>
          </xsl:otherwise>
        </xsl:choose>
        
      </xsl:when>
      
      <xsl:otherwise>
        <xsl:choose>
          
          <!-- currency style -->
          <xsl:when test="contains(@formatCode,'$') or contains(@formatCode,'zł') or contains(@formatCode,'€') or contains(@formatCode,'£')">
            <number:currency-style style:name="{generate-id(.)}">
              <xsl:call-template name="InsertNumberFormatting">
                <xsl:with-param name="formatCode">
                  <xsl:value-of select="@formatCode"/>
                </xsl:with-param>
              </xsl:call-template>
            </number:currency-style>
          </xsl:when>
          
          <!--percentage style -->
          <xsl:when test="contains(@formatCode,'%')">
            <number:percentage-style style:name="{generate-id(.)}">
              <xsl:call-template name="InsertNumberFormatting">
                <xsl:with-param name="formatCode">
                  <xsl:value-of select="@formatCode"/>
                </xsl:with-param>
              </xsl:call-template>
              <number:text>%</number:text>
            </number:percentage-style>
          </xsl:when>
          
          <!-- number style -->
          <xsl:otherwise>
            <number:number-style style:name="{generate-id(.)}">
              <xsl:call-template name="InsertNumberFormatting">
                <xsl:with-param name="formatCode">
                  <xsl:value-of select="@formatCode"/>
                </xsl:with-param>
              </xsl:call-template>
            </number:number-style>
          </xsl:otherwise>
        </xsl:choose>
      </xsl:otherwise>
      
    </xsl:choose>
    
  </xsl:template>
  
  <!-- template to create number format -->
  
  <xsl:template name="InsertNumberFormatting">
    <xsl:param name="formatCode"/>
    
    <!-- '*' is not converted -->
    <xsl:variable name="realFormatCode">
      <xsl:choose>
        <xsl:when test="contains($formatCode,'*')">
          <xsl:value-of select="substring-after($formatCode,'*')"/>
        </xsl:when>
        <xsl:otherwise>
          <xsl:value-of select="$formatCode"/>
        </xsl:otherwise>
      </xsl:choose>
    </xsl:variable>
    
    <!-- adding text -->
    <xsl:if test="starts-with($realFormatCode,'\') and not(starts-with($realFormatCode,'\ '))">
      <xsl:call-template name="AddNumberText">
        <xsl:with-param name="format">
          <xsl:value-of select="$realFormatCode"/>
        </xsl:with-param>
      </xsl:call-template>
    </xsl:if> 
    
    <!-- handle red negative numbers -->
    <xsl:if test="contains($formatCode,'Red')">
      <style:text-properties fo:color="#ff0000"/>
    </xsl:if>
    
    <xsl:variable name="currencyFormat">
      <xsl:choose>
        <xsl:when test="contains($realFormatCode,'zł')">zł</xsl:when>
        <xsl:when test="contains($realFormatCode,'Red')">
          <xsl:value-of select="substring-after(substring-before(substring-after($realFormatCode,'Red]'),']'),'[')"/>
        </xsl:when>
        <xsl:when test="contains($realFormatCode,'[$')">
          <xsl:value-of select="substring-after(substring-before($realFormatCode,']'),'[')"/>
        </xsl:when>
        <xsl:when test="contains($realFormatCode,'$')">$</xsl:when>
        <xsl:when test="contains($realFormatCode,'€')">€</xsl:when>
        <xsl:when test="contains($realFormatCode,'£')">£</xsl:when>
      </xsl:choose>
    </xsl:variable>
    
    <!-- add space at the beginning -->
    <xsl:if test="starts-with($realFormatCode,'_') and not(contains($realFormatCode,'(-') or contains($realFormatCode,'(#') or contains($realFormatCode,'(0'))">
      <number:text>
        <xsl:value-of xml:space="preserve" select="' '"/>
      </number:text>
    </xsl:if>
    
    <!-- add brackets -->
    <xsl:if test="contains($realFormatCode,'(-') or contains($realFormatCode,'(#') or contains($realFormatCode,'(0')">
      <xsl:choose>
        <xsl:when test="starts-with($formatCode,'_')">
          <number:text>
            <xsl:value-of xml:space="preserve" select="' ('"/>
          </number:text>
        </xsl:when>
        <xsl:otherwise>
          <number:text>(</number:text>
        </xsl:otherwise>
      </xsl:choose>
    </xsl:if>
    
    <!-- add '-' at the beginning -->
    <xsl:if test="contains($realFormatCode,'-') and not($currencyFormat and $currencyFormat!='') and not(contains(substring-after($realFormatCode,'#'),'-') or contains(substring-after($realFormatCode,'0'),'-'))">
      <number:text>-</number:text>
    </xsl:if>
    
    <!-- add currency symbol at the beginning -->
    <xsl:if test="$currencyFormat and $currencyFormat!='' and not(contains(substring-before($realFormatCode,$currencyFormat),'0') or contains(substring-before($realFormatCode,$currencyFormat),'#'))">
      
      <!-- add '-' at the beginning -->
      <xsl:if test="contains(substring-after($realFormatCode,$currencyFormat),'-')">
        <number:text>-</number:text>
      </xsl:if>
      <xsl:call-template name="InsertCurrencySymbol">
        <xsl:with-param name="value" select="$currencyFormat"/>
      </xsl:call-template>
      
      <!-- add space after currency symbol -->
      <xsl:if test="contains(substring-after($realFormatCode,$currencyFormat),'\ ') and (contains(substring-after(substring-after($realFormatCode,$currencyFormat),'\ '),'0') or contains(substring-after(substring-after($realFormatCode,$currencyFormat),'\ '),'#'))">
        <number:text>
          <xsl:value-of xml:space="preserve" select="' '"/>
        </number:text>
      </xsl:if>
      
    </xsl:if>
    
    <!-- add '-' at the beginning -->
    <xsl:if test="$currencyFormat and $currencyFormat!='' and (contains(substring-before($realFormatCode,$currencyFormat),'0') or contains(substring-before($realFormatCode,$currencyFormat),'#')) and contains(substring-before($realFormatCode,$currencyFormat),'-')">
      <number:text>-</number:text>
    </xsl:if>
    
    <number:number>
      
      <xsl:variable name="formatCodeWithoutComma">
        <xsl:choose>
          <xsl:when test="contains($realFormatCode,'.')">
            <xsl:value-of select="substring-before($realFormatCode,'.')"/>
          </xsl:when>
          <xsl:otherwise>
            <xsl:value-of select="$realFormatCode"/>
          </xsl:otherwise>
        </xsl:choose>
      </xsl:variable>
      
      <!-- decimal places -->
      <xsl:attribute name="number:decimal-places">
        <xsl:choose>
          <xsl:when test="contains($realFormatCode,'.')">
            <xsl:call-template name="InsertDecimalPlaces">
              <xsl:with-param name="code">
                <xsl:value-of select="substring-after($realFormatCode,'.')"/>
              </xsl:with-param>
              <xsl:with-param name="value">0</xsl:with-param>
            </xsl:call-template>
          </xsl:when>
          <xsl:otherwise>0</xsl:otherwise>
        </xsl:choose>
      </xsl:attribute>
      
      <!-- min integer digits -->
      
      <xsl:attribute name="number:min-integer-digits">
        <xsl:choose>
          <xsl:when test="contains($formatCodeWithoutComma,'0')">
            <xsl:call-template name="InsertMinIntegerDigits">
              <xsl:with-param name="code">
                <xsl:value-of select="substring-after($formatCodeWithoutComma,'0')"/>
              </xsl:with-param>
              <xsl:with-param name="value">1</xsl:with-param>
            </xsl:call-template>
          </xsl:when>
          <xsl:otherwise>0</xsl:otherwise>
        </xsl:choose>
      </xsl:attribute>
      
      <!-- grouping -->
      <xsl:if test="contains($realFormatCode,',')">
        <xsl:choose>
          <xsl:when test="contains(substring-after($realFormatCode,','),'0') or contains(substring-after($realFormatCode,','),'#')">
            <xsl:attribute name="number:grouping">true</xsl:attribute>
          </xsl:when>
          <xsl:otherwise>
            <xsl:attribute name="number:display-factor">
              <xsl:call-template name="UseDisplayFactor">
                <xsl:with-param name="formatBeforeSeparator">
                  <xsl:value-of select="substring-before($realFormatCode,',')"/>
                </xsl:with-param>
                <xsl:with-param name="formatAfterSeparator">
                  <xsl:value-of select="substring-after($realFormatCode,',')"/>
                </xsl:with-param>
                <xsl:with-param name="value">1000</xsl:with-param>
              </xsl:call-template>
            </xsl:attribute>
          </xsl:otherwise>
        </xsl:choose>
      </xsl:if>
      
      <!-- '-' embedded in number format -->
      <xsl:if test="contains(substring-after(substring-before($formatCode,'.'),'#'),'-') or (not(contains($formatCode,'.')) and contains(substring-after($formatCode,'#'),'-')) or contains(substring-after(substring-before($formatCode,'.'),'0'),'-')  or (not(contains($formatCode,'.')) and contains(substring-after($formatCode,'0'),'-'))">
        <xsl:call-template name="FindTextNumberFormat">
          <xsl:with-param name="format">
            <xsl:choose>
              <xsl:when test="contains(substring-after(substring-before($formatCode,'.'),'#'),'-')">
            <xsl:value-of select="concat('#',substring-after(substring-before($formatCode,'.'),'#'))"/>
              </xsl:when>
              <xsl:when test="contains(substring-after(substring-before($formatCode,'.'),'0'),'-')">
                <xsl:value-of select="concat('0',substring-after(substring-before($formatCode,'.'),'0'))"/>
              </xsl:when>
              <xsl:when test="not(contains($formatCode,'.')) and contains(substring-after($formatCode,'#'),'-')">
                <xsl:value-of select="concat('#',substring-after($formatCode,'#'))"/>
              </xsl:when>
              <xsl:when test="not(contains($formatCode,'.')) and contains(substring-after($formatCode,'0'),'-')">
                <xsl:value-of select="concat('0',substring-after($formatCode,'0'))"/>
              </xsl:when>
            </xsl:choose>
          </xsl:with-param>
          <xsl:with-param name="embeddedText">-</xsl:with-param>
        </xsl:call-template>
      </xsl:if>
      
      <!-- '\ ' embedded in number format -->
      <xsl:if test="contains(substring-after(substring-before($formatCode,'.'),'#'),'\ ') or (not(contains($formatCode,'.')) and contains(substring-after($formatCode,'#'),'\ ')) or contains(substring-after(substring-before($formatCode,'.'),'0'),'\ ')  or (not(contains($formatCode,'.')) and contains(substring-after($formatCode,'0'),'\ '))">
        <xsl:call-template name="FindTextNumberFormat">
          <xsl:with-param name="format">
            <xsl:choose>
              <xsl:when test="contains(substring-after(substring-before($formatCode,'.'),'#'),'\ ')">
                <xsl:value-of select="concat('#',substring-after(substring-before($formatCode,'.'),'#'))"/>
              </xsl:when>
              <xsl:when test="contains(substring-after(substring-before($formatCode,'.'),'0'),'\ ')">
                <xsl:value-of select="concat('0',substring-after(substring-before($formatCode,'.'),'0'))"/>
              </xsl:when>
              <xsl:when test="not(contains($formatCode,'.')) and contains(substring-after($formatCode,'#'),'\ ')">
                <xsl:value-of select="concat('#',substring-after($formatCode,'#'))"/>
              </xsl:when>
              <xsl:when test="not(contains($formatCode,'.')) and contains(substring-after($formatCode,'0'),'\ ')">
                <xsl:value-of select="concat('0',substring-after($formatCode,'0'))"/>
              </xsl:when>
            </xsl:choose>
          </xsl:with-param>
          <xsl:with-param name="embeddedText">
            <xsl:value-of xml:space="preserve" select="'\ '"/>
          </xsl:with-param>
        </xsl:call-template>
      </xsl:if>
      
    </number:number>
    
    <!-- add currency symbol at the end -->
    <xsl:if test="$currencyFormat and $currencyFormat!='' and (contains(substring-before($realFormatCode,$currencyFormat),'0') or contains(substring-before($realFormatCode,$currencyFormat),'#'))">
      
      <!-- add space before currency symbol -->
      <xsl:if test="contains(substring-before($realFormatCode,$currencyFormat),'\ ')">
        <number:text>
          <xsl:value-of xml:space="preserve" select="' '"/>
        </number:text>
      </xsl:if>
        
      <xsl:call-template name="InsertCurrencySymbol">
        <xsl:with-param name="value" select="$currencyFormat"/>
      </xsl:call-template>
      
    </xsl:if>
    
    <!-- add brackets -->
    <xsl:if test="contains($realFormatCode,'(-') or contains($realFormatCode,'(#') or contains($realFormatCode,'(0')">
      <number:text>)</number:text>
    </xsl:if>
    
    <!-- add space at the end -->
    <xsl:if test="(contains($realFormatCode,'\ ') and not($currencyFormat and $currencyFormat!='' and (contains(substring-before($realFormatCode,$currencyFormat),'0') or contains(substring-before($realFormatCode,$currencyFormat),'#'))) and not(contains(substring-after($realFormatCode,'\ '),'0') or contains(substring-after($realFormatCode,'\ '),'#'))) or (contains($realFormatCode,'_') and not($currencyFormat and $currencyFormat!='' and (contains(substring-before($realFormatCode,$currencyFormat),'0') or contains(substring-before($realFormatCode,$currencyFormat),'#'))) and not(contains(substring-after($realFormatCode,'_'),'0') or contains(substring-after($realFormatCode,'_'),'#')))">
      <number:text>
        <xsl:value-of xml:space="preserve" select="' '"/>
      </number:text>
    </xsl:if>
    
  </xsl:template>
  
  <!-- template which inserts display factor -->
  
  <xsl:template name="UseDisplayFactor">
    <xsl:param name="formatBeforeSeparator"/>
    <xsl:param name="formatAfterSeparator"/>
    <xsl:param name="value"/>
    <xsl:choose>
      <xsl:when test="$formatAfterSeparator and $formatAfterSeparator!=''">
        <xsl:call-template name="UseDisplayFactor">
          <xsl:with-param name="formatBeforeSeparator" select="$formatBeforeSeparator"/>
          <xsl:with-param name="formatAfterSeparator">
            <xsl:value-of select="substring($formatAfterSeparator,2)"/>
          </xsl:with-param>
          <xsl:with-param name="value">
            <xsl:value-of select="number($value)*1000"/>
          </xsl:with-param>
        </xsl:call-template>
      </xsl:when>
      <xsl:otherwise>
        <xsl:value-of select="$value"/>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>
  
  <!-- template which inserts min integer digits -->
  
  <xsl:template name="InsertMinIntegerDigits">
    <xsl:param name="code"/>
    <xsl:param name="value"/>
    <xsl:choose>
      <xsl:when test="contains($code,'0')">
        <xsl:call-template name="InsertMinIntegerDigits">
          <xsl:with-param name="code">
            <xsl:value-of select="substring($code,0,string-length($code))"/>
          </xsl:with-param>
          <xsl:with-param name="value">
            <xsl:value-of select="$value+1"/>
          </xsl:with-param>
        </xsl:call-template>
      </xsl:when>
      <xsl:otherwise>
        <xsl:value-of select="$value"/>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>
  
  <!-- template which inserts decimal places -->
  
  <xsl:template name="InsertDecimalPlaces">
    <xsl:param name="code"/>
    <xsl:param name="value"/>
    <xsl:choose>
      <xsl:when test="contains($code,'0') or contains($code,'#')">
        <xsl:call-template name="InsertDecimalPlaces">
          <xsl:with-param name="code">
            <xsl:value-of select="substring($code,2)"/>
          </xsl:with-param>
          <xsl:with-param name="value">
            <xsl:value-of select="$value+1"/>
          </xsl:with-param>
        </xsl:call-template>
      </xsl:when>
      <xsl:otherwise>
        <xsl:value-of select="$value"/>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>
  
  <!-- template which adds number text -->
  
  <xsl:template name="AddNumberText">
    <xsl:param name="format"/>
    <xsl:choose>
    <xsl:when test="starts-with($format,'\') and not(starts-with($format,'\ '))">
      <number:text>
        <xsl:value-of select="substring($format,2,1)"/>
      </number:text>
      <xsl:call-template name="AddNumberText">
        <xsl:with-param name="format">
          <xsl:value-of select="substring($format,3)"/>
        </xsl:with-param>
      </xsl:call-template>
    </xsl:when>
      <xsl:when test="starts-with($format,'\ ')">
        <number:text>
          <xsl:value-of xml:space="preserve" select="' '"/>
        </number:text>
      </xsl:when>
    </xsl:choose>
  </xsl:template>
  
  <!-- adding number style with fixed number format -->
  
  <xsl:template match="e:xf" mode="fixedNumFormat">
    <xsl:if test="@numFmtId and @numFmtId &gt; 0 and not(key('numFmtId',@numFmtId))">
      <xsl:choose>
        
        <!-- date style -->
        <xsl:when test="(@numFmtId &gt; 13 and @numFmtId &lt; 18) or @numFmtId = 22">
          <number:date-style style:name="{concat('N',@numFmtId)}">
            <xsl:call-template name="InsertFixedDateFormat">
              <xsl:with-param name="ID">
                <xsl:value-of select="@numFmtId"/>
              </xsl:with-param>
            </xsl:call-template>
          </number:date-style>
        </xsl:when>
        
        <!-- percentage style -->
        <xsl:when test="@numFmtId = 9 or @numFmtId = 10">
          <number:percentage-style style:name="{concat('N',@numFmtId)}">
            <xsl:call-template name="InsertFixedNumFormat">
              <xsl:with-param name="ID">
                <xsl:value-of select="@numFmtId"/>
              </xsl:with-param>
            </xsl:call-template>
          </number:percentage-style>
        </xsl:when>
        
        <!-- number style -->
        <xsl:otherwise>
          <number:number-style style:name="{concat('N',@numFmtId)}">
            <xsl:call-template name="InsertFixedNumFormat">
              <xsl:with-param name="ID">
                <xsl:value-of select="@numFmtId"/>
              </xsl:with-param>
            </xsl:call-template>
          </number:number-style>
        </xsl:otherwise>
        
      </xsl:choose>
    </xsl:if>
  </xsl:template>
  
  <!-- insert currency symbol element -->
  <xsl:template name="InsertCurrencySymbol">
    <xsl:param name="value"/>
    <xsl:choose>
      <xsl:when test="$value = '$$-409' or $value = '$'">
        <number:currency-symbol number:language="en" number:country="US">$</number:currency-symbol>
      </xsl:when>
      <xsl:when test="$value = '$USD'">
        <number:currency-symbol number:language="en" number:country="US">USD</number:currency-symbol>
      </xsl:when>
      <xsl:when test="$value = '$£-809' or $value = '£'">
        <number:currency-symbol number:language="en" number:country="GB">£</number:currency-symbol>
      </xsl:when>
      <xsl:when test="$value = '$GBP'">
        <number:currency-symbol number:language="en" number:country="GB">GBP</number:currency-symbol>
      </xsl:when>
      <xsl:when test="$value = '$€-1' or $value = '$€-2' or $value = '€'">
        <number:currency-symbol>€</number:currency-symbol>
      </xsl:when>
      <xsl:when test="$value = '$EUR'">
        <number:currency-symbol>EUR</number:currency-symbol>
      </xsl:when>
      <xsl:when test="$value = 'zł'">
        <number:currency-symbol number:language="pl" number:country="PL">zł</number:currency-symbol>
      </xsl:when>
      <xsl:when test="$value = '$PLN'">
        <number:currency-symbol number:language="pl" number:country="PL">PLN</number:currency-symbol>
      </xsl:when>
    </xsl:choose>
  </xsl:template>
  
  <!-- template which inserts fixed number format -->
  
  <xsl:template name="InsertFixedNumFormat">
    <xsl:param name="ID"/>
    <number:number>
      <xsl:attribute name="number:decimal-places">
        <xsl:choose>
          <xsl:when test="$ID = 1 or $ID = 3 or $ID = 9">0</xsl:when>
          <xsl:when test="$ID = 2 or $ID = 4 or $ID = 10">2</xsl:when>
          <xsl:otherwise>0</xsl:otherwise>
        </xsl:choose>
      </xsl:attribute>
      <xsl:attribute name="number:min-integer-digits">
        <xsl:choose>
          <xsl:when test="$ID = 1 or $ID = 2 or $ID = 3 or $ID = 4 or $ID = 9 or $ID = 10">1</xsl:when>
          <xsl:otherwise>0</xsl:otherwise>
        </xsl:choose>
      </xsl:attribute>
      <xsl:if test="$ID = 3 or $ID = 4">
        <xsl:attribute name="number:grouping">true</xsl:attribute>
      </xsl:if>
    </number:number>
    <xsl:if test="$ID = 9 or $ID = 10">
      <number:text>%</number:text>
    </xsl:if>
  </xsl:template>
  
  <xsl:template name="ProcessFormat">
    <xsl:param name="format"/>
    <xsl:param name="processedFormat"/>
    <xsl:choose>
      
      <!-- year -->
      <xsl:when test="starts-with($processedFormat,'y')">
        <xsl:choose>
          <xsl:when test="starts-with(substring-after($processedFormat,'y'),'yyy')">
            <number:year number:style="long"/>
            <xsl:call-template name="ProcessFormat">
              <xsl:with-param name="format" select="$format"/>
              <xsl:with-param name="processedFormat">
                <xsl:value-of select="substring-after($processedFormat,'yyyy')"/>
              </xsl:with-param>
            </xsl:call-template>
          </xsl:when>
          <xsl:otherwise>
            <number:year/>
            <xsl:call-template name="ProcessFormat">
              <xsl:with-param name="format" select="$format"/>
              <xsl:with-param name="processedFormat">
                <xsl:value-of select="substring-after($processedFormat,'yy')"/>
              </xsl:with-param>
            </xsl:call-template>
          </xsl:otherwise>
        </xsl:choose>
      </xsl:when>
      
      <xsl:when test="starts-with($processedFormat,'m')">
        <xsl:choose>
          
          <!-- minutes -->
          <xsl:when test="contains(substring-before($format,'m'),'h:')">
            <xsl:choose>
              <xsl:when test="starts-with(substring-after($processedFormat,'m'),'m')">
                <number:minutes number:style="long"/>
                <xsl:call-template name="ProcessFormat">
                  <xsl:with-param name="format" select="$format"/>
                  <xsl:with-param name="processedFormat">
                    <xsl:value-of select="substring-after($processedFormat,'mm')"/>
                  </xsl:with-param>
                </xsl:call-template>
              </xsl:when>
              <xsl:otherwise>
                <number:minutes/>
                <xsl:call-template name="ProcessFormat">
                  <xsl:with-param name="format" select="$format"/>
                  <xsl:with-param name="processedFormat">
                    <xsl:value-of select="substring-after($processedFormat,'m')"/>
                  </xsl:with-param>
                </xsl:call-template>
              </xsl:otherwise>
            </xsl:choose>
          </xsl:when>
          
          <!-- month -->
          <xsl:otherwise>
            <xsl:choose>
              <xsl:when test="starts-with(substring-after($processedFormat,'m'),'mmm')">
                <number:month number:style="long" number:textual="true"/>
                <xsl:call-template name="ProcessFormat">
                  <xsl:with-param name="format" select="$format"/>
                  <xsl:with-param name="processedFormat">
                    <xsl:value-of select="substring-after($processedFormat,'mmmm')"/>
                  </xsl:with-param>
                </xsl:call-template>
              </xsl:when>
              <xsl:when test="starts-with(substring-after($processedFormat,'m'),'mm')">
                <number:month number:textual="true"/>
                <xsl:call-template name="ProcessFormat">
                  <xsl:with-param name="format" select="$format"/>
                  <xsl:with-param name="processedFormat">
                    <xsl:value-of select="substring-after($processedFormat,'mmm')"/>
                  </xsl:with-param>
                </xsl:call-template>
              </xsl:when>
              <xsl:when test="starts-with(substring-after($processedFormat,'m'),'m')">
                <number:month number:style="long"/>
                <xsl:call-template name="ProcessFormat">
                  <xsl:with-param name="format" select="$format"/>
                  <xsl:with-param name="processedFormat">
                    <xsl:value-of select="substring-after($processedFormat,'mm')"/>
                  </xsl:with-param>
                </xsl:call-template>
              </xsl:when>
              <xsl:otherwise>
                <number:month/>
                <xsl:call-template name="ProcessFormat">
                  <xsl:with-param name="format" select="$format"/>
                  <xsl:with-param name="processedFormat">
                    <xsl:value-of select="substring-after($processedFormat,'m')"/>
                  </xsl:with-param>
                </xsl:call-template>
              </xsl:otherwise>
            </xsl:choose>
          </xsl:otherwise>
        </xsl:choose>
      </xsl:when>
      
      <!-- day -->
      <xsl:when test="starts-with($processedFormat,'d')">
        <xsl:choose>
          <xsl:when test="starts-with(substring-after($processedFormat,'d'),'ddd')">
            <number:day-of-week number:style="long"/>
            <xsl:call-template name="ProcessFormat">
              <xsl:with-param name="format" select="$format"/>
              <xsl:with-param name="processedFormat">
                <xsl:value-of select="substring-after($processedFormat,'dddd')"/>
              </xsl:with-param>
            </xsl:call-template>
          </xsl:when>
          <xsl:when test="starts-with(substring-after($processedFormat,'d'),'dd')">
            <number:day-of-week/>
            <xsl:call-template name="ProcessFormat">
              <xsl:with-param name="format" select="$format"/>
              <xsl:with-param name="processedFormat">
                <xsl:value-of select="substring-after($processedFormat,'ddd')"/>
              </xsl:with-param>
            </xsl:call-template>
          </xsl:when>
          <xsl:when test="starts-with(substring-after($processedFormat,'d'),'d')">
            <number:day number:style="long"/>
            <xsl:call-template name="ProcessFormat">
              <xsl:with-param name="format" select="$format"/>
              <xsl:with-param name="processedFormat">
                <xsl:value-of select="substring-after($processedFormat,'dd')"/>
              </xsl:with-param>
            </xsl:call-template>
          </xsl:when>
          <xsl:otherwise>
            <number:day/>
            <xsl:call-template name="ProcessFormat">
              <xsl:with-param name="format" select="$format"/>
              <xsl:with-param name="processedFormat">
                <xsl:value-of select="substring-after($processedFormat,'d')"/>
              </xsl:with-param>
            </xsl:call-template>
          </xsl:otherwise>
        </xsl:choose>
      </xsl:when>
      
      <!-- hours -->
      <xsl:when test="starts-with($processedFormat,'h')">
        <xsl:choose>
          <xsl:when test="starts-with(substring-after($processedFormat,'h'),'h')">
            <number:hours number:style="long"/>
            <xsl:call-template name="ProcessFormat">
              <xsl:with-param name="format" select="$format"/>
              <xsl:with-param name="processedFormat">
                <xsl:value-of select="substring-after($processedFormat,'hh')"/>
              </xsl:with-param>
            </xsl:call-template>
          </xsl:when>
          <xsl:otherwise>
            <number:hours/>
            <xsl:call-template name="ProcessFormat">
              <xsl:with-param name="format" select="$format"/>
              <xsl:with-param name="processedFormat">
                <xsl:value-of select="substring-after($processedFormat,'h')"/>
              </xsl:with-param>
            </xsl:call-template>
          </xsl:otherwise>
        </xsl:choose>
      </xsl:when>
      
      <!-- seconds -->
      <xsl:when test="starts-with($processedFormat,'s')">
        <xsl:choose>
          <xsl:when test="starts-with(substring-after($processedFormat,'s'),'s')">
            <number:seconds number:style="long"/>
            <xsl:call-template name="ProcessFormat">
              <xsl:with-param name="format" select="$format"/>
              <xsl:with-param name="processedFormat">
                <xsl:value-of select="substring-after($processedFormat,'ss')"/>
              </xsl:with-param>
            </xsl:call-template>
          </xsl:when>
          <xsl:otherwise>
            <number:seconds/>
            <xsl:call-template name="ProcessFormat">
              <xsl:with-param name="format" select="$format"/>
              <xsl:with-param name="processedFormat">
                <xsl:value-of select="substring-after($processedFormat,'s')"/>
              </xsl:with-param>
            </xsl:call-template>
          </xsl:otherwise>
        </xsl:choose>
      </xsl:when>
      
      <xsl:when test="starts-with($processedFormat,'AM/PM')">
        <number:am-pm/>
        <xsl:call-template name="ProcessFormat">
          <xsl:with-param name="format" select="$format"/>
          <xsl:with-param name="processedFormat">
            <xsl:value-of select="substring-after($processedFormat,'AM/PM')"/>
          </xsl:with-param>
        </xsl:call-template>
      </xsl:when>
      <xsl:when test="starts-with($processedFormat,'\') or starts-with($processedFormat,'@') or starts-with($processedFormat,';')">
        <xsl:call-template name="ProcessFormat">
          <xsl:with-param name="format" select="$format"/>
          <xsl:with-param name="processedFormat">
            <xsl:value-of select="substring($processedFormat,2)"/>
          </xsl:with-param>
        </xsl:call-template>
      </xsl:when>
      <xsl:when test="string-length($processedFormat) = 0"/>
      <xsl:otherwise>
        <number:text>
          <xsl:value-of xml:space="preserve" select="substring($processedFormat,0,2)"/>
        </number:text>
        <xsl:call-template name="ProcessFormat">
          <xsl:with-param name="format" select="$format"/>
          <xsl:with-param name="processedFormat">
            <xsl:value-of select="substring($processedFormat,2)"/>
          </xsl:with-param>
        </xsl:call-template>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>
  
  <!-- template which inserts fixed date format -->
  
  <xsl:template name="InsertFixedDateFormat">
    <xsl:param name="ID"/>
    <xsl:choose>
      <xsl:when test="$ID = 14">
        <number:month number:style="long"/>
        <number:text>-</number:text>
        <number:day number:style="long"/>
        <number:text>-</number:text>
        <number:year/>
      </xsl:when>
      <xsl:when test="$ID = 15">
        <number:day/>
        <number:text>-</number:text>
        <number:month number:textual="true"/>
        <number:text>-</number:text>
        <number:year/>
      </xsl:when>
      <xsl:when test="$ID = 16">
        <number:day/>
        <number:text>-</number:text>
        <number:month number:textual="true"/>
      </xsl:when>
      <xsl:when test="$ID = 17">
        <number:month number:textual="true"/>
        <number:text>-</number:text>
        <number:year/>
      </xsl:when>
      <xsl:when test="$ID = 22">
        <number:month/>
        <number:text>/</number:text>
        <number:day/>
        <number:text>/</number:text>
        <number:year/>
      </xsl:when>
    </xsl:choose>
  </xsl:template>
  
  <!-- template which adds embedded text in number format -->
  
  <xsl:template name="FindTextNumberFormat">
    <xsl:param name="format"/>
    <xsl:param name="embeddedText"/>
    <xsl:choose>
      <xsl:when test="string-length($format) &gt; 0">
        <xsl:choose>
          <xsl:when test="starts-with($format,$embeddedText)">
            <number:embedded-text number:position="{string-length(translate(substring($format,1+string-length($embeddedText)),$embeddedText,''))}">
              <xsl:value-of xml:space="preserve" select="translate($embeddedText,'\ ',' ')"/>
            </number:embedded-text>
            <xsl:call-template name="FindTextNumberFormat">
              <xsl:with-param name="format">
                <xsl:value-of select="substring($format,1+string-length($embeddedText))"/>
              </xsl:with-param>
              <xsl:with-param xml:space="preserve" name="embeddedText" select="$embeddedText"/>
            </xsl:call-template>
          </xsl:when>
          <xsl:otherwise>
            <xsl:call-template name="FindTextNumberFormat">
              <xsl:with-param name="format">
                <xsl:value-of select="substring($format,2)"/>
              </xsl:with-param>
              <xsl:with-param xml:space="preserve" name="embeddedText" select="$embeddedText"/>
            </xsl:call-template>
          </xsl:otherwise>
        </xsl:choose>
      </xsl:when>
    </xsl:choose>
  </xsl:template>
  
</xsl:stylesheet>
