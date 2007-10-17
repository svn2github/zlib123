<?xml version="1.0" encoding="utf-8"?>
<?xml-stylesheet href="xsl2html.xsl" type="text/xsl" media="screen"?>
<xsl:stylesheet version="1.0"
  xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
  xmlns="http://www.w3.org/1999/xhtml"
  xmlns:h="http://www.w3.org/1999/xhtml"
  exclude-result-prefixes="h"
>
 
 <!-- @Filename: xsl2html.xsl -->
 <!-- @Description:  <p>This XSL stylesheet generates XSLT documentation 
  based on template definitions and comments. This documentation was generated 
  using this stylesheet...</p> 
  The following comments<sup>1/3</sup> will be displayed:
     <ul>
  <li>Comments following the stylesheet definition and starting with a '@'<sup>2/3</sup></li>
  <li>Comments following a template definition and starting with a '@'</li>
  <li>Comments following a variable/parameter/key definition</li> 
     </ul>
  Each template description contains the following information:
  <ul>
  <li>Template comments<sup>2</sup></li>
  <li>Variable names and description</li>
  <li>Parameter names and description</li>
  <li>Template calls</li>
  <li>Calling templates<sup>4</sup></li> 
  </ul>
  <p><b>Notes</b>: 
  <ol>
  <li>Comments can contain xHTML as long as it is well formed. HTML characters (e.g. &lt;) that should not be 
  interpreted have to be escaped using entities.</li>
  <li>Comments starting with '@' should contain a title and a body separated by a column 
      (e.g. '<i><code>&lt;&ndash;&ndash; @Description: This is a comment... &ndash;&ndash;&gt;</code></i>').
      Comments that do not contain a column won't be displayed. 
  </li>
  <li>You can hide some templates by marking them as private. 
  This is achieved by adding the special comment '<i><code>&lt;&ndash;&ndash; @Private &ndash;&ndash;&gt;</code></i>'
    within the template's definition. If you want to display private templates too, you will
    need to set the variable '<b>readmode</b>' (see below) to '<i>private</i>'.
  </li>   
  <li>For match templates, the following calls won't be listed:
  <ul>
  <li>Implicit calls (i.e. no select attributes)</li>
  <li>Calls using an attribute</li>
  <li>Calls using the '.' value</li>
  <li>Calls with predicats</li>
  </ul>
  </li>
  </ol>
 -->
 <!-- @Author: Olivier Durand -->
 <!-- @Created: 08 March 2007 -->
 
 <xsl:output method="html"/>
 <xsl:strip-space elements="*"/>
  
 <xsl:variable name="key-nodes" select="/xsl:stylesheet/xsl:key"/> <!-- A collection of key nodes defined at the root -->
 <xsl:variable name="variable-nodes" select="/xsl:stylesheet/xsl:variable"/> <!-- A collection of variable nodes defined at the root -->
 <xsl:variable name="parameters-nodes" select="/xsl:stylesheet/xsl:param"/> <!-- A collection of parameter nodes defined at the root -->
 <xsl:variable name="readmode" select="'private' "/>  <!-- Determines whether private templates are to be displayed -->
 <xsl:variable name="invalidChars">/@*:|[]()=</xsl:variable> <!-- Invalid URL characters -->
 <xsl:variable name="URLchars">-_x..</xsl:variable> <!-- Characters replacing invalid URL charaters -->  
 

 <xsl:template match="/">
  <!-- @Description: This template lays out the document body and calls other templates -->
  
  <html lang="en">
   <head>
    <link rel="stylesheet" type="text/css" href="styles.css" media="all"/>
    <script type="text/javascript" src="decodeHTML.js"></script>
   </head>
   
   <!--    <body> -->
   <body onload="decodeHTML();">
    <div id="wrapper">
     <h1 id="top">XSLT Documentation</h1>       
     <div class="content">
      <h2>File Information</h2>
      <xsl:call-template name="root-descriptions"/>
      <h2>Template descriptions</h2>
      <h4>Content</h4>
      <ul>
       <xsl:apply-templates select=".//xsl:template[@match]" mode="TOC">
        <xsl:sort select="@match" data-type="text"/>
       </xsl:apply-templates>
       <xsl:apply-templates select=".//xsl:template[@name]" mode="TOC">
        <xsl:sort select="@name" data-type="text"/>
       </xsl:apply-templates>
      </ul>
       
      <xsl:apply-templates select=".//xsl:template[@match]" mode="body">
       <xsl:sort select="@match" data-type="text"/>
      </xsl:apply-templates>
      <xsl:apply-templates select=".//xsl:template[@name]" mode="body">
       <xsl:sort select="@name" data-type="text"/>
      </xsl:apply-templates>
     </div>
    </div>
   </body>
  </html>
 </xsl:template>

 
 <xsl:template match="xsl:template" mode="TOC">
  <!-- @Description: Generates a template entry in the Table of content-->
  
  <xsl:variable name="private" select="./comment()[starts-with(normalize-space(.), '@Private')]"/>
  <!-- Checks if template is private -->
  
  <xsl:if test="($readmode='public' and not($private) ) or $readmode='private' ">
   <li>
    <xsl:call-template name="template-heading">
     <xsl:with-param name="isPrivate" select="boolean($private)"></xsl:with-param>
    </xsl:call-template>
   </li>
  </xsl:if>
 </xsl:template>


 <xsl:template match="xsl:template" mode="body">
  <!-- @Description: Generates a template description. -->

  <xsl:variable name="private" select="./comment()[starts-with(normalize-space(.), '@')][starts-with(normalize-space(.), '@Private')]"/> <!-- Checks if template is private -->    

  <!-- Checks if the description should be displayed if the template is marked as private... -->  
  <xsl:if test="($readmode='public' and not($private) ) or $readmode='private' ">

   <xsl:variable name="mode-name"> <!-- Obtains the mode name (if any) -->
    <xsl:if test="@mode">-<xsl:value-of select="translate(@mode, $invalidChars,$URLchars)"/></xsl:if>
   </xsl:variable>
   
   <h3 id="t-{translate(@name|@match,$invalidChars,$URLchars)}{$mode-name}">
     <xsl:call-template name="template-heading">
      <xsl:with-param name="isPrivate" select="boolean($private)"></xsl:with-param>
     </xsl:call-template>
    </h3>

     <!--Displays the template metadata.-->
   <xsl:apply-templates select="./comment()[starts-with(normalize-space(.), '@')]" mode="metadata"/>

     <!--Display other valuable information.-->
     <xsl:call-template name="template-parameters"/>
     <xsl:call-template name="template-variables"/>
     <xsl:call-template name="template-calls"/>
     <xsl:call-template name="template-calledby"/>

   <div class="nav"><a href="#top">Back to top</a></div>
  </xsl:if>
 </xsl:template>

 
 <xsl:template match="comment()">
  <!-- @Description: Displays an inline comment -->   

    <!--Display the comment-->
    <span class="comment">
     <decodeable>
       : <xsl:value-of select="." disable-output-escaping="yes"/>
     </decodeable>
    </span>

 </xsl:template>
 
 
 <xsl:template match="comment()" mode="metadata">  
  <!-- @Description: Displays heading and body of comments constructed as follows:
   <pre>&lt;&ndash;&ndash; @&lt;Heading&gt;: &lt;Comment body&gt; &ndash;&ndash;&gt;</pre>
   -->
  <!-- @Context: A comment node -->
  
  <h4><xsl:value-of select="substring(normalize-space(substring-before(., ':')), 2)"/></h4>
  <div class="metadata">
   <!-- The decodeable is used in place of the disable-output-escaping option 
    for the browsers not supporting it. It is used by the script disableOutputEscaping.js -->
   <decodeable>
    <xsl:value-of select="normalize-space(substring-after(., ':'))" disable-output-escaping="yes"/>
   </decodeable>  
  </div>
 </xsl:template>


 <xsl:template name="template-parameters">
  <!-- @Description: Displays template's parameters and their description if any. 
   The parameter's first sibling comment is considered to be the description. --> 
  <!-- @Context: A node containing an element &lt;xsd:template&gt;. -->
  
  <xsl:variable name="params" select="xsl:param"/> <!-- The template's parameters -->
  
  <xsl:if test="count($params) &gt; 0">
   <h4>Parameters</h4>
   <ul>
    <xsl:for-each select="$params">
     <xsl:sort select="@name" data-type="text"/>
     <li>
      <dfn>
       $<xsl:value-of select="./@name"/>
       <xsl:if test="./@select">
        = <xsl:value-of select="./@select"/>
       </xsl:if>
      </dfn>
      <xsl:apply-templates select="comment()[1]"/>
      <xsl:apply-templates select="following-sibling::comment()[1]"/>
     </li>
    </xsl:for-each>
   </ul>
  </xsl:if>
 </xsl:template>


 <xsl:template name="template-variables">
  <!-- @Description: Displays template's variables and their description. 
   The variable's first sibling comment is considered to be the description. 
  -->
  <!-- @Context: A node containing a template defition -->
  <xsl:variable name="vars" select=".//xsl:variable"/> <!-- The template's variables -->
  
  <xsl:if test="count($vars) &gt; 0">
   <h4>Variables</h4>
   <ul>
    <xsl:for-each select="$vars">
     <xsl:sort select="@name" data-type="text"/>
     <li>
      <dfn>$<xsl:value-of select="./@name"/></dfn>
       <xsl:apply-templates select="comment()[1]"/>
       <xsl:apply-templates select="following-sibling::comment()[1]"/>
     </li>
    </xsl:for-each>
   </ul>
  </xsl:if>
 </xsl:template>

 
 
 <xsl:template name="template-calls">
  <!-- @Description: Displays template's calls. -->
  <!-- @Context: A node containing a template defition -->
  
  <xsl:variable name="call" select=".//xsl:call-template"/>   <!-- All the call-template  instances -->
  <xsl:variable name="apply" select=".//xsl:apply-templates"/> <!-- All the apply-template instances --> 
  
  <xsl:if test="count($call) + count($apply) &gt; 0">
  <h4>Template calls</h4>
   <ul>
    <xsl:for-each select="$call|$apply">
     <li>
      <xsl:call-template name="template-heading"/>
      <xsl:call-template name="template-paramList">
       <xsl:with-param name="params" select="xsl:with-param"/>
      </xsl:call-template>      
     </li>
    </xsl:for-each>
   </ul>
  </xsl:if>
 </xsl:template>

 <xsl:template name="template-paramList">
  <!-- @Description: Displays an inline list of template parameters, where each 
   parameter is delimited by a comma-->
  <!-- @Context: A node containing a template defition -->
  
  <xsl:param name="params"/>                         <!-- The node set containing the parameters -->
  <xsl:param name="default-values" select="'true'"/> <!-- true: display default values -->
  
  <xsl:if test="count($params) &gt; 0">
   <code>
    <xsl:text>(</xsl:text>
    <xsl:for-each select="$params">
     <xsl:value-of select="@name"/>
     <xsl:if test="$default-values='true' and (.//@select != '' or ./* != '')">= 
      <xsl:value-of select=".//@select"/>
      <xsl:value-of select="./*"/>
     </xsl:if>
     <xsl:if test="position() != last()">, </xsl:if>
    </xsl:for-each>)
   </code>
  </xsl:if>
 </xsl:template>
 
 <xsl:template name="template-heading">
  <!-- @Description: Dislpays a template name and mode if any and creates an href anchor & link.   
  -->
  <!-- @Context: A node containing either of the followings:
   <ul>
     <li>xsl:template</li>
     <li>xsl:call-template</li>
     <li>xsl:aply-template</li>
   </ul>   
  -->
  
  <xsl:param name="isPrivate" select="false()"/> <!-- if true will append '(private)" to the heading -->
  
  <xsl:variable name="private"> <!-- Contains the string to append -->
   <xsl:if test="$isPrivate"> (private) </xsl:if>
  </xsl:variable>

  <xsl:variable name="mode-name"> <!-- Obtains the mode name (if any) -->
   <xsl:if test="@mode">-<xsl:value-of select="translate(@mode, $invalidChars,$URLchars)"/></xsl:if>
  </xsl:variable>
  
  <xsl:choose>
   <xsl:when test="@name">named: 
    <a href="#t-{@name}"><xsl:value-of select="@name"/>
     <xsl:call-template name="template-attributes"/>
     <xsl:value-of select="$private"/>
    </a>
   </xsl:when>
   <xsl:when test="@match">match:    
    
    <a href="#t-{translate(@match,$invalidChars,$URLchars)}{$mode-name}">
     <xsl:value-of select="@match"/>
     <xsl:call-template name="template-attributes"/>
     <xsl:value-of select="$private"/>
    </a> 
   </xsl:when>

   <!-- The template is called using the 'xsl:apply-template' element -->
   <xsl:otherwise>match: <code>
    <xsl:variable name="s" select="@select"/> <!-- The value of the select attribute -->    
    
    <xsl:choose>
     
     <!-- If the template definition is found, a hyperlink is created... -->
     <xsl:when test="//xsl:template[@match=$s]">
      <a href="#t-{translate($s,$invalidChars,$URLchars)}{$mode-name}">
       <xsl:value-of select="$s"/>
       <xsl:call-template name="template-attributes"/>
      </a>
     </xsl:when>
     
     <!-- OTherwise no href... -->
     <xsl:otherwise>
      <xsl:value-of select="$s"/>
      <xsl:call-template name="template-attributes"/>
     </xsl:otherwise>
    </xsl:choose></code>    
    <xsl:value-of select="$private"/>
   </xsl:otherwise>
   
  </xsl:choose>
  
 </xsl:template>

 <xsl:template name="template-attributes">
  <!-- @Description: Displays the mode and priority attributes (if set) -->
  
  <xsl:if test="@mode"> mode='<xsl:value-of select="@mode"/>'</xsl:if>
  <xsl:if test="@priority"> priority='<xsl:value-of select="@priority"/>'</xsl:if>    
 </xsl:template>
 
 
 <xsl:template name="template-calledby">
  <!-- @Description: Displays the calling templates. -->
  <!-- @Context: A xsl:template node  -->
  
  <xsl:variable name="name" select="@name|@match"/> <!-- The template name/pattern -->
  <xsl:variable name="mode" select="@mode"/> <!-- The template's mode -->
  
  <xsl:variable name="parents" select="//xsl:stylesheet/xsl:template[.//xsl:call-template[@name=$name] or 
   ((.//xsl:apply-templates[substring(@select, string-length(@select) - string-length($name)+1) = $name]) and 
   ((boolean($mode) and .//xsl:apply-templates[@mode=$mode]) or not($mode) ))   ]"/>
  
  <xsl:if test="count($parents) &gt; 0">
  <h4>Explicitly called by</h4>
   <ul>
    <xsl:for-each select="$parents">
     <li><xsl:call-template name="template-heading"/></li>
    </xsl:for-each>
   </ul>
  </xsl:if>
 </xsl:template>

 <xsl:template name="root-descriptions">
  <!-- @Description: Generates the XSL description -->
  
  <xsl:apply-templates select="/xsl:stylesheet/comment()[starts-with(normalize-space(.), '@')]" mode="metadata"/>
  <xsl:call-template name="root-externals"/>          
  <xsl:call-template name="root-defs"/>  
 </xsl:template>
 
 <xsl:template name="root-externals">
  <!-- @Description: Lists up linked external XSLT, CSS stylesheets and javascripts. -->
  
  <dl>
   <xsl:call-template name="util-show-refs">
    <xsl:with-param name="heading" select="'Included XSLT'"/>
    <xsl:with-param name="resource" select="//xsl:include/@href"/>
   </xsl:call-template>
   
   <xsl:call-template name="util-show-refs">
    <xsl:with-param name="heading" select="'Imported XSLT'"/>
    <xsl:with-param name="resource" select="//xsl:import/@href"/>
   </xsl:call-template>
   
   <xsl:call-template name="util-show-refs">
    <xsl:with-param name="heading" select="'CSS linked'"/>
    <xsl:with-param name="resource" select="//h:link[@rel='stylesheet']/@href"/>
   </xsl:call-template>
   
   <xsl:call-template name="util-show-refs">
    <xsl:with-param name="heading" select="'External scripts'"/>
    <xsl:with-param name="resource" select="//h:script/@src"/>
   </xsl:call-template>
  </dl>
 </xsl:template>
 
 <xsl:template name="root-defs">
  <!-- @Description: Displays the <i>Global definitions</i> section if and only if either of the followings is declared at the root of the document:
   <ul>
   <li>key</li>
   <li>param</li>
   <li>variable</li>
   </ul>
  --> 
  
  <xsl:if test="(count($key-nodes) &gt; 0) or 
   (count($variable-nodes) &gt; 0)  or 
   (count($parameters-nodes) &gt; 0)">
   <h2>Global definitions</h2>  
   <xsl:call-template name="root-keys"/>
   <xsl:call-template name="root-variables"/>
   <xsl:call-template name="root-parameters"/>
   <div class="nav"><a href="#top">Back to top</a></div>   
  </xsl:if>
 </xsl:template> 

 <xsl:template name="root-keys">
  <!--@Description:  Display the key section if some keys are defined at the root of the document-->
  
    <xsl:if test="count($key-nodes) &gt; 0">
      <h3>Key(s)</h3>
     <ul>
      <xsl:for-each select="$key-nodes">
       <xsl:sort select="@name" data-type="text"/>
       <li>
        <dfn><xsl:value-of select="./@name"/></dfn>
        <code>(match=<xsl:value-of select="./@match"/></code>,
        <code> use=<xsl:value-of select="./@use"/>)</code>
        
        <!-- If the next node is a comment, display it... -->
        <xsl:if test="generate-id(./following-sibling::node()[1]) = 
                      generate-id(./following-sibling::comment()[1]) ">
          <xsl:apply-templates select="./following-sibling::node()[1]"/>
        </xsl:if>
        </li>
      </xsl:for-each>
     </ul>
    </xsl:if>
 </xsl:template>

 <xsl:template name="root-variables">
  <!-- @Description: Display the variable section if some global variables are defined at the root of the document-->
  
    <xsl:if test="count($variable-nodes) &gt; 0">
     <h3>Variable(s)</h3>
     <ul>
      <xsl:for-each select="$variable-nodes">
       <xsl:sort select="@name" data-type="text"/>
       <li>
        <dfn><xsl:value-of select="./@name"/></dfn>
        <!-- If the next node is a comment, display it... -->
        <xsl:if test="generate-id(./following-sibling::node()[1]) = 
                      generate-id(./following-sibling::comment()[1]) ">
          <xsl:apply-templates select="./following-sibling::comment()[1]"/>
        </xsl:if>
        </li>
      </xsl:for-each>
     </ul>
    </xsl:if>
 </xsl:template>

 <xsl:template name="root-parameters">
  <!-- @Description: Display the parameter section if some global parameters 
   are defined at the root of the document
  -->
  
  <xsl:if test="count($parameters-nodes) &gt; 0">
   <h3>Parameter(s)</h3>
   <ul>
    <xsl:for-each select="$parameters-nodes">
     <xsl:sort select="@name" data-type="text"/>
     <li>
      <dfn><xsl:value-of select="./@name"/></dfn>
      <xsl:if test="./@select != '' ">
        <code>(select=<xsl:value-of select="./@select"/>)</code>
      </xsl:if>
      <!-- If the next node is a comment, display it... -->
      <xsl:if test="generate-id(./following-sibling::node()[1]) = 
                    generate-id(./following-sibling::comment()[1]) ">
        <xsl:apply-templates select="./following-sibling::comment()[1]"/>
      </xsl:if>
     </li>
    </xsl:for-each>
   </ul>
  </xsl:if>
 </xsl:template>

 <xsl:template name="util-show-refs">
  <!-- @Description: Displays a list of external resources with href links -->
  <!-- @Context: None -->
  
  <xsl:param name="heading"/><!-- section name-->
  <xsl:param name="resource"/><!-- A collection of nodes that link to external resources-->
  
  <xsl:if test="count($resource) &gt; 0">
   <h4><xsl:value-of select="$heading"/></h4>
   <ul>
   <xsl:for-each select="$resource">
    <li><a href="{.}.html"><xsl:value-of select="."/></a></li>
   </xsl:for-each>
   </ul>
  </xsl:if>
 </xsl:template>
 
</xsl:stylesheet>
