<?xml version="1.0" encoding="UTF-8"?>
<p:library xmlns:p="http://www.w3.org/ns/xproc" xmlns:c="http://www.w3.org/ns/xproc-step" xmlns:xtlxo="http://www.xtpxlib.nl/ns/xoffice"
  xmlns:xtlcon="http://www.xtpxlib.nl/ns/container" xmlns:xtlc="http://www.xtpxlib.nl/ns/common" xmlns:pxp="http://exproc.org/proposed/steps"
  xmlns:local="#local.word-to-xml.mod.xpl" version="1.0" xpath-version="2.0" exclude-inline-prefixes="#all">
  <!-- ================================================================== -->
  <!--* 
    Various conversions for Word (.docx) files	
  -->
  <!-- ================================================================== -->

  <p:declare-step type="xtlxo:extract-docx">

    <p:documentation>
      Extracts the contents of a Word file in a more useable format.
    </p:documentation>

    <!-- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -->
    <!-- SETUP: -->

    <p:option name="docx-href" required="true">
      <p:documentation>Document reference of the docx file to process (must have file:// in front).</p:documentation>
    </p:option>

    <p:output port="result" primary="true" sequence="false">
      <p:documentation>
        The resulting XML representation of the Word file.
      </p:documentation>
    </p:output>

    <p:import href="../../../xtpxlib-container/xplmod/container.mod/container.mod.xpl"/>

    <p:variable name="debug" select="false()"/>

    <!-- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -->

    <!-- Extract all XML content: -->
    <xtlcon:zip-to-container>
      <p:with-option name="href-source-zip" select="$docx-href"/>
      <p:with-option name="add-document-target-paths" select="false()"/>
    </xtlcon:zip-to-container>

    <p:xslt>
      <p:input port="stylesheet">
        <p:document href="xsl/extract-docx-1.xsl"/>
      </p:input>
      <p:with-param name="debug" select="$debug"/>
    </p:xslt>

  </p:declare-step>

  <!-- ================================================================== -->
  <!-- PROCESS XML INTO A WORD .DOCX FILE: -->

  <p:declare-step type="xtlxo:create-docx">

    <p:documentation>
      Turns the Word XML (back) into a Word .docx file, using a template file. This must be in the format the extract-docx pipeline creates.
    </p:documentation>

    <!-- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -->
    <!-- SETUP: -->

    <p:input port="source" primary="true" sequence="false">
      <p:documentation>
        The Word XML that must be turned into the .docx file.
      </p:documentation>
    </p:input>

    <p:option name="template-docx-href" required="true">
      <p:documentation>Document reference of the template docx file to use (must have file:// in front).</p:documentation>
    </p:option>

    <p:option name="result-docx-href" required="true">
      <p:documentation>Document reference where to write the resulting .docx file.</p:documentation>
    </p:option>

    <p:output port="result" primary="true" sequence="false">
      <p:documentation>
        The xtlcon:document-container as written to the final Word file.
      </p:documentation>
    </p:output>

    <p:import href="../../../xtpxlib-container/xplmod/container.mod/container.mod.xpl"/>

    <!-- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -->

    <p:identity name="input-source"/>
    <p:sink/>

    <!-- Get the template document in a container: -->
    <xtlcon:zip-to-container>
      <p:with-option name="href-source-zip" select="$template-docx-href"/>
      <p:with-option name="add-document-target-paths" select="true()"/>
    </xtlcon:zip-to-container>

    <!-- Add the Word XML to the container (as a normal container document): -->
    <p:insert match="/*" position="first-child">
      <p:input port="insertion">
        <p:inline>
          <xtlcon:document document-type="xtpxlib-word-xml">
            <WORDXMLHERE/>
          </xtlcon:document>
        </p:inline>
      </p:input>
    </p:insert>
    <p:viewport match="WORDXMLHERE">
      <p:identity>
        <p:input port="source">
          <p:pipe port="result" step="input-source"/>
        </p:input>
      </p:identity>
    </p:viewport>

    <!-- Merge the Word XML into the real .docx document XML: -->
    <p:xslt>
      <p:input port="stylesheet">
        <p:document href="xsl/create-docx-1.xsl"/>
      </p:input>
      <p:with-param name="null" select="()"/>
    </p:xslt>

    <!-- Write the container back to disk: -->
    <xtlcon:container-to-zip>
      <p:with-option name="href-target-zip" select="$result-docx-href"/>
    </xtlcon:container-to-zip>

  </p:declare-step>

</p:library>
