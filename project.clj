(defproject malcolmx "0.1.0"
  :description "FIXME: write description"
  :url "http://example.com/FIXME"
  :license {:name "Eclipse Public License"
            :url "http://www.eclipse.org/legal/epl-v10.html"}
  :dependencies [[org.clojure/clojure "1.6.0"]
                 [com.betinvest/poi "3.11-20150430"]
                 [org.apache.poi/poi-ooxml "3.11" :exclusions [org.apache.poi/poi]]
                 [org.apache.poi/poi-ooxml-schemas "3.11" :exclusions [org.apache.poi/poi]]
                 [commons-codec/commons-codec "1.9"]
                 [org.apache.xmlbeans/xmlbeans "2.6.0"]
                 [dom4j/dom4j "1.6.1"]
                 [org.apache.commons/commons-math3 "3.1.1"]
                 [org.slf4j/slf4j-api "1.7.7"]
                 [org.clojure/tools.logging "0.3.1"]
                 [org.flatland/protobuf "0.8.1"]
                 [com.google.protobuf/protobuf-java "2.5.0"]]

  :repositories ^:replace [["snapshots" {:url "http://nassau.favorit/repository/snapshots"
                                         :username :env
                                         :password :env}]
                           ["releases" {:url "http://nassau.favorit/repository/internal"
                                        :username :env
                                        :password :env}]]
  :plugins [[lein-ring "0.8.2"]
            [lein-protobuf "0.4.1"]]

  :profiles {:dev  {:source-paths ["dev"]
                    :global-vars {*warn-on-reflection* true}
                    :dependencies [[ns-tracker "0.2.2"]
                                   [aprint "0.1.0"]
                                   [http-kit.fake "0.2.1"]
                                   [http-kit "2.1.16"]
                                   [criterium "0.4.3"]
                                   [im.chit/vinyasa "0.3.4"]
                                   [org.clojure/tools.trace "0.7.8"]
                                   [ch.qos.logback/logback-core "1.1.2"]
                                   [ch.qos.logback/logback-classic "1.1.2"]]

                    :injections [(require '[vinyasa.inject :as inject])
                                 (require 'aprint.core)
                                 (require 'clojure.pprint)
                                 (require 'clojure.tools.trace)
                                 (require 'criterium.core)
                                 (inject/in clojure.core >
                                            [aprint.core aprint]
                                            [clojure.pprint pprint]
                                            [clojure.tools.trace trace]
                                            [criterium.core bench])]}})
