<?xml version="1.0" encoding="UTF-8"?>
<p:declare-step xmlns:p="http://www.w3.org/ns/xproc" xmlns:c="http://www.w3.org/ns/xproc-step" xmlns:map="http://www.w3.org/2005/xpath-functions/map"
  xmlns:array="http://www.w3.org/2005/xpath-functions/array" xmlns:xs="http://www.w3.org/2001/XMLSchema" xmlns:local="#local.vsb_j3x_byb"
  xmlns:xwebdoc="http://www.xtpxlib.nl/ns/webdoc" xmlns:xdoc="http://www.xtpxlib.nl/ns/xdoc" version="3.0" exclude-inline-prefixes="#all" name="this">

  <p:documentation>
    Pipeline to create the documentation for this component.
  </p:documentation>

  <!-- ======================================================================= -->
  <!-- IMPORTS: -->

  <p:import href="../../../xtpxlib-webdoc/xpl3/xdoc-to-componentdoc-website.xpl"/>

  <!-- ======================================================================= -->
  <!-- DEVELOPMENT SETTINGS: -->

  <p:option name="develop" as="xs:boolean" static="true" select="false()"/>

  <!-- ======================================================================= -->
  <!-- PORTS: -->
  <p:output port="result" primary="true" sequence="false" content-types="xml" serialization="map{'method': 'xml', 'indent': true()}">
    <p:documentation>A report XML about the component documentation generation</p:documentation>
  </p:output>
 
  <!-- ================================================================== -->
  <!-- MAIN: -->

  <p:variable name="href-parameters" select="resolve-uri('../data/xtpxlib-componentdoc-parameters.xml', static-base-uri())"/>
  <p:variable name="output-directory" select="resolve-uri('../../docs/', static-base-uri())"/>
  <p:variable name="href-readme" select="resolve-uri('../../README.md', static-base-uri())"/>

  <!-- Generate the website: -->
  <xwebdoc:xdoc-to-componentdoc-website>
    <p:with-input port="source" href="../source/xtpxlib-xoffice-componentdoc.xml"/>
    <p:with-option name="component-name" select="(doc('../../component-info.xml')/*/@name, '?COMPONENTNAME?')[1]"/>
    <p:with-option name="href-parameters" select="$href-parameters"/>
    <p:with-option name="output-directory" select="$output-directory"/>
    <p:with-option name="href-readme" select="$href-readme"/>
  </xwebdoc:xdoc-to-componentdoc-website>
  
</p:declare-step>
