import express from 'express';
import XlsxPopulate from 'xlsx-populate';

const app = express();
const PORT = 9090;

async function getValuesFromExcel() {
    try {
        const workbook = await XlsxPopulate.fromFileAsync("./ganadores-funciones.xlsx");
        const sheet = workbook.sheet(0);  

        const rows = sheet.usedRange().value();

        const values = [];
        for (let i = 1; i < rows.length; i++) {
            const row = rows[i];
            const cellValueCedula = row[1]; 
            const cellValueNombre = row[2];   
            const cellValueFuncione = row[4]; 
            
            if (cellValueCedula !== undefined && cellValueCedula !== null &&
                cellValueNombre !== undefined && cellValueNombre !== null &&
                cellValueFuncione !== undefined && cellValueFuncione !== null 
            ) {
                values.push({ nombre: cellValueNombre, cedula: cellValueCedula, funciones: cellValueFuncione });
            }
        }

        return values;
    } catch (error) {
        console.error("Error leyendo el archivo Excel:", error);
        throw error;
    }
}

app.get('/sorteo/:numGanadores', async (req, res) => {
    const numGanadores = parseInt(req.params.numGanadores, 10)
    try {
        const values = await getValuesFromExcel();
        
        const shuffled = values.sort(() => 0.5 - Math.random());
        const ganadores = shuffled.slice(0, numGanadores);  

        return res.json(ganadores);
    } catch (error) {
        console.error("Error en la solicitud:", error);
        return res.status(500).json({ error: 'Error al procesar la solicitud' });
    }
});

app.listen(PORT, () => {
    console.log(`Servidor corriendo en el puerto ${PORT}`);
});
