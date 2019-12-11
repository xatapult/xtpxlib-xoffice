<?xml version="1.0" encoding="UTF-8"?>
<p:declare-step xmlns:p="http://www.w3.org/ns/xproc" xmlns:xwebdoc="http://www.xtpxlib.nl/ns/webdoc" version="1.0" xpath-version="2.0"
  exclude-inline-prefixes="#all">

  <p:documentation>
    Generates the documentation for this component.
  </p:documentation>

  <!-- ================================================================== -->
  <!-- SETUP: -->

  <p:output port="result" primary="true" sequence="false"/>
  <p:serialization port="result" method="xml" encoding="UTF-8" indent="true" omit-xml-declaration="false"/>

  <p:import href="../../../xtpxlib-webdoc/xpl/xdoc-to-componentdoc-website.xpl"/>

  <!-- ================================================================== -->

  <p:variable name="href-parameters" select="resolve-uri('../data/xtpxlib-componentdoc-parameters.xml', static-base-uri())"/>
  <p:variable name="output-directory" select="resolve-uri('../../docs/', static-base-uri())"/>
  <p:variable name="href-readme" select="resolve-uri('../../README.md', static-base-uri())"/>

  <!-- Generate the website: -->
  <xwebdoc:xdoc-to-componentdoc-website>
    <p:input port="source">
      <p:document href="../source/xtpxlib-xoffice-componentdoc.xml"/>
    </p:input>
    <p:with-option name="component-name" select="(doc('../../component-info.xml')/*/@name, '?COMPONENTNAME?')[1]"/>
    <p:with-option name="href-parameters" select="$href-parameters"/>
    <p:with-option name="output-directory" select="$output-directory"/>
    <p:with-option name="href-readme" select="$href-readme"/>
  </xwebdoc:xdoc-to-componentdoc-website>

</p:declare-step>
