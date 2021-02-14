(ns fxl-example.core
  (:require
    [cheshire.core :as cheshire]
    [clojure.java.io :as io]
    [clj-time.core :as t]
    [clj-time.format :as f]
    [zero-one.fxl.core :as fxl]))

;; ---------------------------------------------------------------------------
;; Data
;; ---------------------------------------------------------------------------

(def inventory
  (cheshire/parse-stream (io/reader "data/inventory.json")))

(def raw-materials
  (sort (distinct (mapcat keys (mapcat vals (vals inventory))))))

(defn timestamp []
  (f/unparse (f/formatter "yyyy/MM/dd HH:mm") (t/now)))

(def form-data
  {:doc-id "F-ABC-123"
   :revision-number 2
   :site "Jakarta"
   :timestamp (timestamp)})

;; ---------------------------------------------------------------------------
;; Styles
;; ---------------------------------------------------------------------------

(defn fill-border [cell]
  (-> cell
      (assoc-in [:style :bottom-border] {:style :thin :colour :black})
      (assoc-in [:style :left-border] {:style :thin :colour :black})
      (assoc-in [:style :right-border] {:style :thin :colour :black})
      (assoc-in [:style :top-border] {:style :thin :colour :black})))

(defn bold [cell]
  (assoc-in cell [:style :bold] true))

(defn align-right [cell]
  (assoc-in cell [:style :horizontal] :right))

(defn align-center [cell]
  (assoc-in cell [:style :horizontal] :center))

(defn grey-bg [cell]
  (assoc-in cell [:style :background-colour] :grey_25_percent))

(defn num-data-format [cell]
  (assoc-in cell [:style :data-format] "#,##0"))

(defn auto-col-size [cell]
  (assoc-in cell [:style :col-size] :auto))

(defn highlight-shortage [cell]
  (let [qty (:value cell)]
    (cond
      (< 100 qty) cell
      (< 0 qty) (assoc-in cell [:style :background-colour] :yellow)
      :else (-> cell
                (assoc-in [:style :background-colour] :red)
                (assoc-in [:style :font-colour] :white)))))

(defn style-form-cell [cell]
  (let [cell (fill-border cell)]
    (if (= 0 (-> cell :coord :col))
      (-> cell bold align-right grey-bg)
      (-> cell align-center))))

(defn style-inventory-cell [cell]
  (let [row  (-> cell :coord :row)
        col  (-> cell :coord :col)]
    (-> cell
        align-center
        fill-border
        (cond->
          (and (not= row 0) (not= col 0)) num-data-format)
        (cond->
          (and (not= row 0) (not= row 1) (not= col 0)) highlight-shortage)
        (cond->
          (or (= row 0) (= row 1) (= col 0)) bold)
        (cond->
          (or (= row 0) (= row 1) (= col 0)) grey-bg)
        (cond->
          (= col 0) align-right))))

;; ---------------------------------------------------------------------------
;; Cells
;; ---------------------------------------------------------------------------
(def form-header-cells
  (fxl/table->cells
    [["Document ID:"     (:doc-id form-data)]
     ["Revision Number:" (:revision-number form-data)]
     ["Site:"            (:site form-data)]
     ["Timestamp:"       (:timestamp form-data)]]))

(def create-footer-cells
  (fxl/table->cells [["Created By:" nil] ["Date:" nil]]))

(def check-footer-cells
  (fxl/table->cells [["Checked By:" nil] ["Date:" nil]]))


(defn inventory-table-header-cells [quarter months]
  (fxl/concat-below
    (fxl/row->cells [quarter
                     nil (first months) nil
                     nil (second months) nil
                     nil (last months) nil])
    (fxl/row->cells ["Raw Material"
                     "Opening" "Inflows" "Outflows"
                     "Opening" "Inflows" "Outflows"
                     "Opening" "Inflows" "Outflows"])))

(def raw-material-column-cells
  (fxl/table->cells (map #(vector %) raw-materials)))

(defn single-month-inventory-cells [month]
  (let [months-inventory (inventory month)]
    (fxl/table->cells
      (map
        #(vector
           ((months-inventory "opening") %)
           ((months-inventory "inflows") %)
           ((months-inventory "outflows") %))
        raw-materials))))

(defn single-quarter-inventory-cells [quarter months]
  (map
    style-inventory-cell
    (fxl/concat-below
      (inventory-table-header-cells quarter months)
      (fxl/concat-right
        raw-material-column-cells
        (single-month-inventory-cells (first months))
        (single-month-inventory-cells (second months))
        (single-month-inventory-cells (last months))))))

(def inventory-cells
  (fxl/concat-below
    (fxl/pad-below 1 (single-quarter-inventory-cells "Q1" ["Jan" "Feb" "Mar"]))
    (fxl/pad-below 1 (single-quarter-inventory-cells "Q2" ["Apr" "May" "Jun"]))
    (fxl/pad-below 1 (single-quarter-inventory-cells "Q3" ["Jul" "Aug" "Sep"]))
    (single-quarter-inventory-cells "Q4" ["Oct" "Nov" "Dec"])))

;; ---------------------------------------------------------------------------
;; Spreadsheet
;; ---------------------------------------------------------------------------

(fxl/write-xlsx!
  (map auto-col-size
    (fxl/concat-below
      (fxl/pad-below 2 (map style-form-cell form-header-cells))
      (fxl/pad-below 2 inventory-cells)
      (fxl/concat-right
        (fxl/pad-right 2 (map style-form-cell create-footer-cells))
        (map style-form-cell check-footer-cells))))
  "data/inventory.xlsx")

