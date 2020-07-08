(defproject clj-xlsxio "0.6.5"
  :description "xlsxio for clojure"
  :url "https://github.com/lsevero/clj-xlsxio"
  :license {:name "EPL-2.0 OR GPL-2.0-or-later WITH Classpath-exception-2.0"
            :url "https://www.eclipse.org/legal/epl-2.0/"}
  :dependencies [[net.java.dev.jna/jna "5.5.0"]
                 [org.clojure/clojure "1.10.0"]
                 [joda-time "1.6"]]
  :profiles {:dev {:plugins [[cider/cider-nrepl "0.22.3"]
                             [lein-cloverage "1.1.1"]]
                   :global-vars {*warn-on-reflection* true}
                   :main main
                   :repl-options {:init-ns clj-xlsxio.read}
                   :source-paths ["src" "test" "examples"]}}
  :repositories [["releases" {:url "https://repo.clojars.org/"
                              :username :env/clojars_auth
                              :password :env/clojars_token}]]
  :source-paths ["src"]
  :java-source-paths ["java"])
