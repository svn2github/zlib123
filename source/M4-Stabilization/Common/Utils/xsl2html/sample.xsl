<?xml version="1.0" encoding="UTF-8"?>
<?xml-stylesheet href="xsl2html.xsl" type="text/xsl" media="screen"?>
<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform">

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
  
  <!-- @Filename: sample.xsl -->
  <!-- @Description: This is a sample stylesheet. You may use it as a template for your new
    xsl files. To display the documentation for this document, simply drag and drop this file 
    on your web browser. You can add as many 'tagged' comments as you want. A tagged comment 
    is constructed as follows @Heading: comment.
  -->
  <!-- @Created: dd-mm-yyyy -->
  <!-- @Tag: This is a tagged comment.... -->
  
  <xsl:variable name="global-var" /> <!-- Global-val description -->
  <xsl:param name="global-param"/> <!-- global-param description -->
  <xsl:key name="global-key" match="aPattern" use="@id"/> <!-- global-ke description -->
  
 
  <xsl:template name="myTemplate">
    <!-- @Description: This is a sample template definition. You may use <b>xhtml</b> in the 
      description body... You can add as many 'tagged' comments as you want within the template's
      definition
    -->  
    <!-- @Context: The node set that are in the context before invoking this template -->
    <!-- @Tag: Tagged comment.... -->
   
    <xsl:param name="myParam"/> <!-- (string|number|boolean) parameter's description... -->
    
    <xsl:variable name="myVar"/> <!-- (string|number|boolean) variable's description... -->
    
    <!-- Template's body -->
    
  </xsl:template>

  <xsl:template name="myPrivateTemplate">
    <!-- @Private -->
    <!-- @Description: This is a private template definition since it has got the @Private
      special comment (see above). Private templates will be hidden 
      in the documentation unless you set the 'readmode' variable to true() in xsl2html.xsl.
    -->  
    <!-- @Context: The node set that are in the context before invoking this template -->
    
    <xsl:param name="myParam"/> <!-- (string|number|boolean) parameter's description... -->
    
    <xsl:variable name="myVar"/> <!-- (string|number|boolean) variable's description... -->
    
    <!-- Template's body -->
    
  </xsl:template>
  
  <xsl:template match="comment()">
    <!-- @Description: This is a sample match template -->
    <!-- @Context: see template's pattern definition -->
    
    <xsl:param name="myParam"/> <!-- (string|number|boolean) parameter's description... -->
    
    <xsl:variable name="myVar"/> <!-- (string|number|boolean) variable's description... -->
    
    <!-- Template's body -->    
  </xsl:template>


</xsl:stylesheet>