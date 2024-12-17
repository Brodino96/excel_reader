import readXlsxFile from "read-excel-file/node";
import fs from "fs";

const Files: string[] = fs.readdirSync("data/")
const Coords = {
    "A": 0, "B": 1, "C": 2, "D": 3, "E": 4, "F": 5, "G": 6, "H": 7, "I": 8, "J": 9, "K": 10, "L": 11, "M": 12, "N":13,
    "O": 14, "P": 15, "Q": 16, "R": 17, "S": 18, "T": 19, "U": 20, "V": 21, "W": 22, "X": 23, "Y": 24, "Z": 25
}

async function init() {

    Files.forEach(filename => {
        readXlsxFile(`data/${filename}`, { sheet: "Calcoli su AS400"}).then(function (rows) {
            let client = rows[1][Coords["A"]]
            let name = rows[0][Coords["D"]]
            let job = rows[6][Coords["B"]]
            let total = rows[57][Coords["E"]]
            let cableCost = rows[52][Coords["G"]]
            let materialCost = rows[52][Coords["L"]]
            let panelSale = rows[52][Coords["N"]]
            let value = rows[52][Coords["P"]]

            let result = {
                "Cliente": client,
                "Nome": name,
                "Lavoro": job,
                "Totale": total,
                "Costo cavi": cableCost,
                "Costo materiale": materialCost,
                "Vendita quadro": panelSale,
                "Valore": value
            }

            try {
                let dataArray: Record<string, unknown>[] = [];
                
                if (fs.existsSync("src/result.json")) {
                    const fileData = fs.readFileSync("src/result.json", "utf8");
                    if (fileData) {
                        dataArray = JSON.parse(fileData);
                    }
                }
                //dataArray.push(result);
                dataArray[dataArray.length] = result
        
                fs.writeFileSync("src/result.json", JSON.stringify(dataArray, null, 2), 'utf8');
        
                console.log(`Oggetto aggiunto correttamente al file ${"src/result.json"}`);
            } catch (error) {
                console.error(`Errore durante la scrittura del file JSON: ${error}`);
            }

        })
    });
};

init()