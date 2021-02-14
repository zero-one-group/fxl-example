(ns fxl-example.simulate-data
  (:require
    [cheshire.core :as cheshire]
    [clojure.java.io :as io]))

(def raw-material-names
  ["Arrowroot" "Caffeine" "Calciferol" "Calcium Bromate" "Casein" "Chlorine"
   "Chlorine Dioxide" "Corn Syrup" "Dipotassium Phosphate" "Disodium Phosphate"
   "Edible Bone phosphate" "Extenders" "Fructose" "Gelatine"
   "H. Veg. protein" "Invert Sugar" "Iodine" "Lactose" "Niacin/Nicotinic Acid"
   "Polysorbate 60" "Potassium Bromate" "Sodium Chloride/Salt" "Sucrose" "Thiamine"
   "Vanillin" "Yellow 2G"])

(def months
  ["Jan" "Feb" "Mar" "Apr" "May" "Jun" "Jul" "Aug" "Sep" "Oct" "Nov" "Dec"])

(defn random-quantities-per-raw-material [low high]
  (->> raw-material-names
       (map #(vector % (+ low (rand-int (- high low)))))
       (into {})))

(def opening
  (random-quantities-per-raw-material 1000 10000))

(def inflows
  (->> months
       (map #(vector % (random-quantities-per-raw-material 0 2000)))
       (into {})))

(def outflows
  (->> months
       (map #(vector % (random-quantities-per-raw-material 500 2000)))
       (into {})))


(defn closing [starting in out]
  (into {}
    (for [[raw-material init-qty] starting]
      [raw-material (-> init-qty (+ (in raw-material)) (- (out raw-material)))])))


(defn simulate-inventory [opening inflows outflows months]
  (loop [acc         {}
         latest-qtys opening
         next-months months]
    (if (empty? next-months)
      acc
      (let [month  (first next-months)
            in     (inflows month)
            out    (outflows month)
            new-qtys (closing latest-qtys in out)]
       (recur (assoc acc month {:opening latest-qtys
                                :inflows in
                                :outflows out
                                :closing new-qtys})
              new-qtys
              (rest next-months))))))

(let [inventory (simulate-inventory opening inflows outflows months)]
  (cheshire/generate-stream
    inventory
    (io/writer "data/inventory.json")
    {:pretty true}))
