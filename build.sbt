
name := "scala-excel"

version := "0.1-SNAPSHOT"

scalaVersion := "2.11.1"


libraryDependencies ++= {
  Seq(
    "org.apache.poi" % "poi" % "3.9",
    "org.apache.poi" % "poi-ooxml" % "3.9",
    "info.folone" %% "poi-scala" % "0.14",
    "com.google.guava" % "guava" % "12.0",
    "org.scalatest" % "scalatest_2.11" % "2.2.1" % "test"
  )
}