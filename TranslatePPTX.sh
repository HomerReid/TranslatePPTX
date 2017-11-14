#!/bin/bash

export JAVA=/usr/lib/jvm/java-8-openjdk-amd64/jre/bin/java 

if [ "x${POIHOME}" == "x" ]
then
  echo "error: must set POIHOME environment variable to head of Apache POI installation tree"
  exit
fi

export POIBUILD=${POIHOME}/build

export CLASSPATH=""
export CLASSPATH="${CLASSPATH}:${POIBUILD}/classes"
export CLASSPATH="${CLASSPATH}:${POIBUILD}/examples-classes"
export CLASSPATH="${CLASSPATH}:${POIBUILD}/excelant-classes"
export CLASSPATH="${CLASSPATH}:${POIBUILD}/excelant-test-classes"
export CLASSPATH="${CLASSPATH}:${POIBUILD}/integration-test-classes"
export CLASSPATH="${CLASSPATH}:${POIBUILD}/ooxml-classes"
export CLASSPATH="${CLASSPATH}:${POIBUILD}/ooxml-lite-merged"
export CLASSPATH="${CLASSPATH}:${POIBUILD}/ooxml-test-classes"
export CLASSPATH="${CLASSPATH}:${POIBUILD}/scratchpad-classes"
export CLASSPATH="${CLASSPATH}:${POIBUILD}/scratchpad-test-classes"
export CLASSPATH="${CLASSPATH}:${POIBUILD}/test-classes"
export CLASSPATH="${CLASSPATH}:${POIBUILD}/../ooxml-lib/xmlbeans-2.6.0.jar"
export CLASSPATH="${CLASSPATH}:${POIBUILD}/OtherStuffIDownloaded/openxml4j-1.0-beta.jar"
export CLASSPATH="${CLASSPATH}:${POIBUILD}/OtherStuffIDownloaded/ooxml-schemas-1.3.jar"
export CLASSPATH="${CLASSPATH}:${POIBUILD}/OtherStuffIDownloaded/poi-ooxml-3.17.jar"
export CLASSPATH="${CLASSPATH}:${POIBUILD}/OtherStuffIDownloaded/poi-ooxml-schemas-3.17.jar"

${JAVA} org.apache.poi.xslf.extractor.TranslatePPTX $@
