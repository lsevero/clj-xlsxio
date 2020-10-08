(ns clj-xlsxio.read-test
  (:require [clj-xlsxio.read :as sut]
            [clojure.test :refer [deftest is testing]]
            [clojure.java.io :as io]
            [clojure.string :as st]))

(deftest read-xlsx
  (testing "basic reading from string filename"
    (let [res (sut/read-xlsx "examples/test.xlsx")]
      (is (= (first res)
             ["column1" "column2" "column3"]))
      (is (= (last res)
             ["43766" "" "←"]))))

  (testing "basic reading convert to enumerated maps by column number."
    (is (= (-> "examples/test.xlsx" sut/read-xlsx sut/xlsx->enumerated-maps)
           (list {0 "column1", 1 "column2", 2 "column3"}
                 {0 "34262", 1 "", 2 "3"}
                 {0 "", 1 "b", 2 "c"}
                 {0 "", 1 "", 2 ""}
                 {0 "z", 1 "x", 2 "c"}
                 {0 "", 1 "", 2 ""}
                 {0 "43766", 1 "", 2 "←"}))))

  (testing "convert to enumerated maps using excel column name as keys"
    (is (= (-> "examples/test.xlsx" sut/read-xlsx sut/xlsx->excel-enumerated-maps)
           (list {:A "column1", :B "column2", :C "column3"}
                 {:A "34262", :B "", :C "3"}
                 {:A "", :B "b", :C "c"}
                 {:A "", :B "", :C ""}
                 {:A "z", :B "x", :C "c"}
                 {:A "", :B "", :C ""}
                 {:A "43766", :B "", :C "←"}))))

  (testing "convert to maps using the content of first rows as keys"
    (is (= (-> "examples/test.xlsx" sut/read-xlsx sut/xlsx->column-title-maps)
           (list {:column1 "34262", :column2 "", :column3 "3"}
                 {:column1 "", :column2 "b", :column3 "c"}
                 {:column1 "", :column2 "", :column3 ""}
                 {:column1 "z", :column2 "x", :column3 "c"}
                 {:column1 "", :column2 "", :column3 ""}
                 {:column1 "43766", :column2 "", :column3 "←"}))))

  (testing "pass a function to modify the formating of the column names"
    (is (= (-> "examples/test.xlsx" sut/read-xlsx (sut/xlsx->column-title-maps :column-fn st/upper-case))
           (list {:COLUMN1 "34262", :COLUMN2 "", :COLUMN3 "3"}
                 {:COLUMN1 "", :COLUMN2 "b", :COLUMN3 "c"}
                 {:COLUMN1 "", :COLUMN2 "", :COLUMN3 ""}
                 {:COLUMN1 "z", :COLUMN2 "x", :COLUMN3 "c"}
                 {:COLUMN1 "", :COLUMN2 "", :COLUMN3 ""}
                 {:COLUMN1 "43766", :COLUMN2 "", :COLUMN3 "←"}))))

  (testing "enable the keys of the maps to be strings."
    (is (= (-> "examples/test.xlsx" sut/read-xlsx (sut/xlsx->column-title-maps :str-keys true))
           (list {"column1" "34262", "column2" "", "column3" "3"}
                 {"column1" "", "column2" "b", "column3" "c"}
                 {"column1" "", "column2" "", "column3" ""}
                 {"column1" "z", "column2" "x", "column3" "c"}
                 {"column1" "", "column2" "", "column3" ""}
                 {"column1" "43766", "column2" "", "column3" "←"}))))

  (testing "also possible to compose both previous two options and modify the string keys in the maps"
    (is (= (-> "examples/test.xlsx" sut/read-xlsx (sut/xlsx->column-title-maps :str-keys true
                                                                               :column-fn st/upper-case))
           (list {"COLUMN1" "34262", "COLUMN2" "", "COLUMN3" "3"}
                 {"COLUMN1" "", "COLUMN2" "b", "COLUMN3" "c"}
                 {"COLUMN1" "", "COLUMN2" "", "COLUMN3" ""}
                 {"COLUMN1" "z", "COLUMN2" "x", "COLUMN3" "c"}
                 {"COLUMN1" "", "COLUMN2" "", "COLUMN3" ""}
                 {"COLUMN1" "43766", "COLUMN2" "", "COLUMN3" "←"}))))

  (testing "basic reading from java.io.File"
    (let [res (sut/read-xlsx (io/file "examples/test.xlsx"))]
      (is (= res
             (list ["column1" "column2" "column3"]
                   ["34262" "" "3"]
                   ["" "b" "c"]
                   ["" "" ""]
                   ["z" "x" "c"]
                   ["" "" ""]
                   ["43766" "" "←"])))))

  (testing "basic reading from java.io.BufferedInputStream"
    (let [res (sut/read-xlsx (io/input-stream "examples/test.xlsx"))]
      (is (= res
             (list ["column1" "column2" "column3"]
                   ["34262" "" "3"]
                   ["" "b" "c"]
                   ["" "" ""]
                   ["z" "x" "c"]
                   ["" "" ""]
                   ["43766" "" "←"]))))))


(deftest read-xlsx-coerce
  (testing "coerce values by passing generic functions to be executed in each cell."
    (let [res (-> "examples/coerce_test.xlsx"
                  sut/read-xlsx
                  (sut/coerce [(comp inc #(Long/parseLong %))
                               st/upper-case
                               sut/excel-date->java-date]))
          test-collumn-pred (fn [col pred]
                              (apply (every-pred pred)
                                     (map #(nth % col) res)))]
      (is (test-collumn-pred 0 #(instance? Long %)))
      (is (test-collumn-pred 1 #(every? (fn [^Character x] (Character/isUpperCase x)) %)))
      (is (test-collumn-pred 2 #(instance? java.util.Date %))))))


(deftest list-sheets
  (testing "list sheets inside the xlsx"
    (is (= (sut/list-sheets "examples/test.xlsx")
           ["Sheet1"])))

  (testing "list sheets reading from java.io.File"
    (is (= (sut/list-sheets (io/file "examples/test.xlsx"))
           ["Sheet1"])))

  (testing "list sheets reading from java.io.BufferedInputStream"
    (is (= (sut/list-sheets (io/input-stream "examples/test.xlsx"))
           ["Sheet1"]))))
