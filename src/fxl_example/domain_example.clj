(require '[zero-one.fxl.core :as fxl])

(def cells
  #{{:value "Inv.No"    :coord {:row 0 :col 0}
     :style {:background-colour :grey_25_percent}}
    {:value "Category"  :coord {:row 0 :col 1}
     :style {:background-colour :grey_25_percent}}
    {:value 44          :coord {:row 1 :col 0}}
    {:value "Furniture" :coord {:row 1 :col 1}}
    {:value 23          :coord {:row 2 :col 0}}
    {:value "Grocery"   :coord {:row 2 :col 1}}})

(fxl/write-xlsx!
  cells
  "data/domain_example.xlsx")

;(fxl/write-xlsx! cells xlsx-path)

;(fxl/write-sheets!
  ;cells
  ;{:spreadsheet-id spreadsheet-id
   ;:credentials    credentials})


;(def header-cells
  ;#{...})

;(def left-body-cells
  ;#{...})

;(def right-body-cells
  ;#{...})

;(def footer-cells
  ;#{...})


;(def body-cells
  ;(fxl/concat-right
    ;(fxl/pad-right left-body-cells)
    ;right-body-cells))

;(def spreadsheet-cells
  ;(fxl/concat-below
    ;(fxl/pad-below header-cells)
    ;(fxl/pad-below body-cells)
    ;footer-cells))

;(fxl/write-xlsx!
  ;spreadsheet-cells
  ;xlsx-path)

;(def form-data
  ;{:doc-id "F-ABC-123",
   ;:revision-number 2,
   ;:site "Jakarta",
   ;:timestamp "2021/02/21 04:34"})

;(def inventory
  ;{"Jan" {"opening"  {"Arrowroot" 8429 ...}
          ;"inflows"  {"Arrowroot" 1645 ...}
          ;"outflows" {"Arrowroot" 1000 ...}
          ;"closing"  {"Arrowroot" 9074 ...}}
   ;"Feb" ...
   ;"Mar" ...
   ;...})
