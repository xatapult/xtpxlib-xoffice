<?xml version="1.0" encoding="UTF-8"?>
<p:library xmlns:p="http://www.w3.org/ns/xproc" xmlns:c="http://www.w3.org/ns/xproc-step" xmlns:map="http://www.w3.org/2005/xpath-functions/map"
  xmlns:array="http://www.w3.org/2005/xpath-functions/array" xmlns:xs="http://www.w3.org/2001/XMLSchema" xmlns:local="#local.ldq_m1p_dyb"
  xmlns:xtlxo="http://www.xtpxlib.nl/ns/xoffice" xmlns:xtlcon="http://www.xtpxlib.nl/ns/container" version="3.0" exclude-inline-prefixes="#all">

  <p:documentation>
    Library with some support/shared code for the xtpxlib-xoffice component.
  </p:documentation>

  <!-- ================================================================== -->

  <p:declare-step type="xtlxo:get-xo-container">

    <p:documentation>Returns an xtpxlib-container structure, suitable for processing Microsoft Open Office documents.</p:documentation>

    <!-- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -->

    <p:import href="../../xtpxlib-container/xpl3mod/zip-to-container/zip-to-container.xpl"/>

    <p:output port="result" primary="true" sequence="false" content-types="xml">
      <p:documentation>The resulting xtpxlib-container structure.</p:documentation>
    </p:output>

    <p:option name="href" as="xs:string" required="true" >
      <p:documentation>Document reference of the Microsoft Open Office file to process.</p:documentation>
    </p:option>

    <!-- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -->

    <!-- Extract the Open Office document, which is actually a zip file, into a container structure. 
         Don't forget to treat these weird .rels files as XML also. -->
    <xtlcon:zip-to-container href-source-zip="{$href}">
      <p:with-option name="override-content-types" select="[ ['\.rels$', 'text/xml'] ]"/>
    </xtlcon:zip-to-container>

  </p:declare-step>

  <!-- ======================================================================= -->

  <p:declare-step type="xtlxo:xlsx-xo-container-to-xml">

    <p:documentation>Turns an xo container (as created by xtlxo:get-xo-container) of an xlsx (Excel) file into 
      xlsx-extract format (see ../xsd/xlsx-extract.xsd).</p:documentation>

    <!-- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -->

    <p:input port="source" primary="true" sequence="false" content-types="xml">
      <p:documentation>The .xlsx file xo container</p:documentation>
    </p:input>

    <p:output port="result" primary="true" sequence="false" content-types="xml" serialization="map{'method': 'xml', 'indent': true()}">
      <p:documentation>The .xlsx file in xlsx-extract XML format.</p:documentation>
    </p:output>

    <!-- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -->

    <!-- Transform the contents into something manageable: -->
    <p:xslt>
      <p:with-input port="stylesheet" href="xsl-xoffice.mod/extract-xlsx-1.xsl"/>
    </p:xslt>

    <!-- Remove any empty rows and cells: -->
    <p:xslt>
      <p:with-input port="stylesheet" href="xsl-xoffice.mod/extract-xlsx-2.xsl"/>
    </p:xslt>

  </p:declare-step>

</p:library>
