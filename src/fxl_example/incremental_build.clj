(ns fxl-example.incremental-build
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
;; Component 1: form header
;; ---------------------------------------------------------------------------

(def form-header-cells
  (fxl/table->cells
    [["Document ID:"     (:doc-id form-data)]
     ["Revision Number:" (:revision-number form-data)]
     ["Site:"            (:site form-data)]
     ["Timestamp:"       (:timestamp form-data)]]))

(fxl/write-xlsx!
  (->> form-header-cells
       (map #(fxl/shift-right 1 %))
       (map #(fxl/shift-down 1 %)))
  "data/form-header.xlsx")


;; ---------------------------------------------------------------------------
;; Component 2: form footers
;; ---------------------------------------------------------------------------

(def create-footer-cells
  (fxl/table->cells [["Created By:" nil] ["Date:" nil]]))

(def check-footer-cells
  (fxl/table->cells [["Checked By:" nil] ["Date:" nil]]))

(def footer-cells-unpadded
  (fxl/concat-right
    create-footer-cells
    check-footer-cells))

(fxl/write-xlsx!
  (->> footer-cells-unpadded
       (map #(fxl/shift-right 1 %))
       (map #(fxl/shift-down 1 %)))
  "data/footer-unpadded.xlsx")

;; ---------------------------------------------------------------------------
;; Component 2: form footers padded
;; ---------------------------------------------------------------------------

(def footer-cells
  (fxl/concat-right
    (fxl/pad-right 2 create-footer-cells)
    check-footer-cells))

(fxl/write-xlsx!
  (->> footer-cells
       (map #(fxl/shift-right 1 %))
       (map #(fxl/shift-down 1 %)))
  "data/footer.xlsx")

;; ---------------------------------------------------------------------------
;; Component 3: inventory tables
;; Sub-Component 1: header
;; ---------------------------------------------------------------------------

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

(def q1-table-header
  (inventory-table-header-cells "Q1" ["Jan" "Feb" "Mar"]))

(fxl/write-xlsx!
  (->> q1-table-header
       (map #(fxl/shift-right 1 %))
       (map #(fxl/shift-down 1 %)))
  "data/inventory-table-header.xlsx")

;; ---------------------------------------------------------------------------
;; Component 3: inventory tables
;; Sub-Component 2: raw-material column
;; ---------------------------------------------------------------------------

(def raw-material-column-cells
  (fxl/table->cells
    (map #(vector %) raw-materials)))

(fxl/write-xlsx!
  (->> raw-material-column-cells
       (map #(fxl/shift-right 1 %))
       (map #(fxl/shift-down 1 %)))
  "data/raw-material-column.xlsx")

;; ---------------------------------------------------------------------------
;; Component 3: inventory tables
;; Sub-Component 3: single-month inventory
;; ---------------------------------------------------------------------------

(defn single-month-inventory-cells [month]
  (let [months-inventory (inventory month)]
    (fxl/table->cells
      (map
        #(vector
           ((months-inventory "opening") %)
           ((months-inventory "inflows") %)
           ((months-inventory "outflows") %))
        raw-materials))))

(def jan-inventory-cells
  (single-month-inventory-cells "Jan"))

(fxl/write-xlsx!
  (->> jan-inventory-cells
       (map #(fxl/shift-right 1 %))
       (map #(fxl/shift-down 1 %)))
  "data/single-month-inventory.xlsx")

;; ---------------------------------------------------------------------------
;; Component 3: inventory tables
;; Putting together single quarter inventory
;; ---------------------------------------------------------------------------

(defn single-quarter-inventory-cells [quarter months]
  (fxl/concat-below
    (inventory-table-header-cells quarter months)
    (fxl/concat-right
      raw-material-column-cells
      (single-month-inventory-cells (first months))
      (single-month-inventory-cells (second months))
      (single-month-inventory-cells (last months)))))

(def q1-inventory-cells
  (single-quarter-inventory-cells "Q1" ["Jan" "Feb" "Mar"]))

(fxl/write-xlsx!
  (->> q1-inventory-cells
       (map #(fxl/shift-right 1 %))
       (map #(fxl/shift-down 1 %)))
  "data/single-quarter-inventory.xlsx")


;; ---------------------------------------------------------------------------
;; Component 3: inventory tables
;; Putting together the four quarters
;; ---------------------------------------------------------------------------

(def inventory-cells
  (fxl/concat-below
    (fxl/pad-below
      (single-quarter-inventory-cells
        "Q1" ["Jan" "Feb" "Mar"]))
    (fxl/pad-below
      (single-quarter-inventory-cells
        "Q2" ["Apr" "May" "Jun"]))
    (fxl/pad-below
      (single-quarter-inventory-cells
        "Q3" ["Jul" "Aug" "Sep"]))
    (single-quarter-inventory-cells
      "Q4" ["Oct" "Nov" "Dec"])))

(fxl/write-xlsx!
  (->> inventory-cells
       (map #(fxl/shift-right 1 %))
       (map #(fxl/shift-down 1 %)))
  "data/full-inventory.xlsx")

;; ---------------------------------------------------------------------------
;; Putting together the components
;; ---------------------------------------------------------------------------

(def all-cells
  (fxl/concat-below
    (fxl/pad-below 2 form-header-cells)
    (fxl/pad-below 2 inventory-cells)
    (fxl/concat-right
      (fxl/pad-right 2 create-footer-cells)
      check-footer-cells)))

(fxl/write-xlsx!
  (->> all-cells
       (map #(fxl/shift-right 1 %))
       (map #(fxl/shift-down 1 %)))
  "data/unstyled-inventory.xlsx")

;; ---------------------------------------------------------------------------
;; Modular style functions (1/2)
;; ---------------------------------------------------------------------------

(defn fill-border [cell]
  (-> cell
      (assoc-in [:style :bottom-border]
                {:style :thin :colour :black})
      (assoc-in [:style :left-border]
                {:style :thin :colour :black})
      (assoc-in [:style :right-border]
                {:style :thin :colour :black})
      (assoc-in [:style :top-border]
                {:style :thin :colour :black})))

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

;; ---------------------------------------------------------------------------
;; Modular style functions (2/2)
;; ---------------------------------------------------------------------------

(defn highlight-shortage [cell]
  (let [qty (:value cell)]
    (cond
      (< 100 qty)
      cell
      (< 0 qty)
      (assoc-in cell [:style :background-colour] :yellow)
      :else
      (-> cell
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
          (and (not= row 0) (not= col 0))
          num-data-format)
        (cond->
          (and (not= row 0) (not= row 1) (not= col 0))
          highlight-shortage)
        (cond->
          (or (= row 0) (= row 1) (= col 0))
          bold)
        (cond->
          (or (= row 0) (= row 1) (= col 0))
          grey-bg)
        (cond->
          (= col 0) align-right))))

;; ---------------------------------------------------------------------------
;; values ⊥ coordinates ⊥ styles
;; ---------------------------------------------------------------------------

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

(fxl/write-xlsx!
  (map auto-col-size
    (fxl/concat-below
      (fxl/pad-below 2 (map style-form-cell form-header-cells))
      (fxl/pad-below 2 inventory-cells)
      (fxl/concat-right
        (fxl/pad-right 2 (map style-form-cell create-footer-cells))
        (map style-form-cell check-footer-cells))))
  "data/inventory.xlsx")
