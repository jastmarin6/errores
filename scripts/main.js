
function uploadFiles() {
    var file1 = document.getElementById('file1').files[0];
    var file2 = document.getElementById('file2').files[0];
    var file3 = document.getElementById('file3').files[0];

    var data1, data2, data3;

    Papa.parse(file1, {
        header: true,
        complete: function(results) {
            data1 = results.data;
            Papa.parse(file2, {
                header: true,
                complete: function(results) {
                    data2 = results.data;
                    Papa.parse(file3, {
                        header: true,
                        complete: function(results) {
                            data3 = results.data;
                            processFiles(data1, data2, data3);
                        }
                    });
                }
            });
        }
    });
}

function processFiles(data1, data2, data3) {
    const combinedData = data3.map(row => {
        const combinedRow = { ...row };
        const rowContrato = combinedRow.CONTRATO ? combinedRow.CONTRATO.trim() : null;

        // Funciones auxiliares
        const addObservation = (condition, message) => {
            if (condition) combinedRow.OBSERVACION_ERROR += message;
        };

        const allNo = fields => fields.every(field => combinedRow[field] === "NO");
        const anyYes = fields => fields.some(field => combinedRow[field] === "SI");

        // Búsqueda y cruce con data1 (Estado 7)
        const file1Match = data1.find(d1 => d1.CONTRATO && rowContrato && d1.CONTRATO.trim() === rowContrato);
        if (file1Match) {
            Object.assign(combinedRow, {
                DESCRIPCION_ESTADO_PRODUCTO: file1Match.DESCRIPCION_ESTADO_PRODUCTO || '',
                ORDEN_TRABAJO: file1Match.ORDEN_TRABAJO || '',
                CODIGO_TIPO_TRABAJO: file1Match.CODIGO_TIPO_TRABAJO || '',
                FECHA_ASIGNACION: file1Match.FECHA_ASIGNACION || ''
            });
        } else {
            Object.assign(combinedRow, {
                DESCRIPCION_ESTADO_PRODUCTO: '',
                ORDEN_TRABAJO: '',
                CODIGO_TIPO_TRABAJO: '',
                FECHA_ASIGNACION: ''
            });
        }

        // Búsqueda y cruce con data2 (Estado 8)
        const file2Match = data2.find(d2 => d2.CONTRATO && rowContrato && d2.CONTRATO.trim() === rowContrato);
        if (file2Match) {
            Object.assign(combinedRow, {
                DESCCAUSAL: file2Match.DESCCAUSAL || '',
                ID_TIPO_TRABAJO: file2Match.ID_TIPO_TRABAJO || '',
                DESC_ESTADO_PROD: file2Match.DESC_ESTADO_PROD || ''
            });
        } else {
            Object.assign(combinedRow, {
                DESCCAUSAL: '',
                ID_TIPO_TRABAJO: '',
                DESC_ESTADO_PROD: ''
            });
        }

        // Normalizar texto
        const normalizeText = text => text ? text.normalize("NFD").replace(/[\u0300-\u036f]/g, "").toLowerCase() : '';

        // Renombrar DESCCAUSAL
        const descausalMap = {
            "defecto critico servicio suspendido acepta trabajos": "DEFECTO",
            "defecto critico y no acepta reparaciones": "DEFECTO",
            "ot. cumplida instalacion con defecto": "DEFECTO",
            "pendiente por certificar acepta trabajos": "DEFECTO",
            "pendiente por certificar no acepta trabajos": "DEFECTO",
            "ot. cumplida": "DEFECTO",
            "trabajos realizados - pendiente inspeccion y/o certificacion": "DEFECTO",
            "trabajos realizados por terceros": "DEFECTO",
            "certificada": "CERTIFICADA",
            "instalacion certificada": "CERTIFICADA"
        };

        const normalizedDescausal = normalizeText(combinedRow.DESCCAUSAL);
        combinedRow.DESCCAUSAL = descausalMap[normalizedDescausal] || combinedRow.DESCCAUSAL || '';

        // Crear columna de efectiva
        const efectivaConditions = [
            "CERTIFICADA", "CERTIFICADO", "CERTIFICADA CON NOVEDAD", "CERTIFICADO CON NOVEDAD",
            "INSPECCIONADA CON DEFECTO CRITICO", "INSPECCIONADA CON DEFECTO NO CRITICO"
        ];
        combinedRow.Efectiva = efectivaConditions.includes(combinedRow.Resultado_de_la_Inspeccion) ? "Efectiva" : "No efectiva";

        // Evaluar la columna "accion"
        const accionRules = [
            { condition: file1Match && !file2Match && combinedRow.Efectiva === "Efectiva", action: "Legalizar" },
            { condition: !file1Match && !file2Match && combinedRow.Efectiva === "Efectiva", action: "Solicitar" },
            { condition: ["certificada", "certificada con novedad"].includes(normalizeText(combinedRow.Resultado_de_la_Inspeccion)) && combinedRow.DESCCAUSAL === "DEFECTO", action: "Legalizar" },
            { condition: ["inspeccionada con defecto critico", "inspeccionada con defecto no critico"].includes(normalizeText(combinedRow.Resultado_de_la_Inspeccion)) && ["DEFECTO", "CERTIFICADA"].includes(combinedRow.DESCCAUSAL), action: "No legalizar" },
            { condition: ["certificada", "certificada con novedad", "certificado", "certificado con novedad"].includes(normalizeText(combinedRow.Resultado_de_la_Inspeccion)) && combinedRow.DESCCAUSAL === "CERTIFICADA", action: "No legalizar" }
        ];

        combinedRow.accion = accionRules.find(rule => rule.condition)?.action || "No Action";

        // Evaluar el diligenciamiento del acta y agregar columna OBSERVACION_ERROR
        combinedRow.OBSERVACION_ERROR = "";

        const resultadoMap = {
            "CERTIFICADA": "certificada",
            "CERTIFICADO": "certificada",
            "CERTIFICADA CON NOVEDAD": "certificada",
            "CERTIFICADO CON NOVEDAD": "certificada",
            "INSPECCIONADA CON DEFECTO CRITICO": "critico",
            "INSPECCIONADA CON DEFECTO NO CRITICO": "no critico"
        };
        const resultado = resultadoMap[combinedRow.Resultado_de_la_Inspeccion] || "";

        // Reglas de validación
        const validationRules = [
            { condition: resultado === "certificada" && combinedRow.Sin_defectos !== "SI", 
                message: "Error de diligenciamiento en resultado de inspeccion/ " },
            { condition: resultado === "certificada" && (combinedRow.Con_defectos_criticos !== "" || combinedRow.Con_defectos_no_criticos !== ""), 
                message: "error de cierre/ marca defectos/ " },
            { condition: resultado === "critico" && (combinedRow.Con_defectos_criticos !== "SI" || combinedRow.Con_defectos_no_criticos === "NO"), 
                message: "Error de diligenciamiento en resultado de inspeccion/ " },
            { condition: resultado === "no critico" && combinedRow.Con_defectos_no_criticos !== "SI", 
                message: "Error de diligenciamiento en resultado de inspeccion/ " },
            { condition: resultado === "no critico" && combinedRow.Con_defectos_no_criticos === "" && combinedRow.Con_defectos_criticos !== "", 
                message: "error de cierre- marca solo criticos en resultado de inspeccion/ " },
            { condition: combinedRow.Efectiva === "No efectiva" && combinedRow.Sin_defectos !== "" && combinedRow.Con_defectos_criticos !== "" && combinedRow.Con_defectos_no_criticos !== "", 
                message: "error de cierre/ " },
            { condition: resultado === "certificada" && (combinedRow.Continua_en_Servicio === "NO" || combinedRow.cumple_requisitos_y_criterios_verificados_resoluciones_90902_y_41385 === "NO"), 
                message: "Error de diligenciamiento en continua en servicio o cumple 90902/ " },
            { condition: (resultado === "critico" || resultado === "no critico") && combinedRow.cumple_requisitos_y_criterios_verificados_resoluciones_90902_y_41385 === "SI", 
                message: "Error de diligenciamiento en cumple 90902/ " },
            { condition: resultado === "critico" && combinedRow.Continua_en_Servicio === "SI", 
                message: "Error de diligenciamiento en continua en servicio/ " },
            { condition: (combinedRow.Resultado_de_la_Inspeccion === "CERTIFICADA" || combinedRow.Resultado_de_la_Inspeccion === "CERTIFICADO") 
                && (anyYes(["medidor_no_ventilado", "fuga_en_cm", "cm_aislamiento_eléctrico", "cm_bajo_terreno", "art_de_3_famila_en_semisotanos", "distancia_a_elementos_de_combustion", "secadora_sin_ducto", "conector_flex_en_superf_calientes", "espacio_entre_cielo_falso", "metodo_acoplamiento"]) 
                || combinedRow.NOVEDADES_FI22 !== ""), 
                message: "Error de cierre diligencia informe novedades FI-22/ "
            },
            { condition: allNo(["Embebida", "A_la_vista", "Enterrada", "Conducto", "OTRO_TRAZADO"]), 
                message: "No diligencia trazado/ " },
            { condition: allNo(["Acero", "Cobre", "PEALPE", "PEMD", "CSST", "Otro_Material"]), 
                message: "No diligencia material/ " },
            { condition: allNo(["Roscada", "Soldada", "abocinada", "anillada"]), 
                message: "No diligencia tipo de union/ " },
            { condition: combinedRow.Mecanismo_de_control_de_sobrepresion_del_regulador_descarga_al_interior === "N/A" && combinedRow.Efectiva === "Efectiva", 
                message: "No evalua mecanismo de sobrepresion/ " },
            { condition: combinedRow.Correccion === "NO" && combinedRow.Lectura === "" && combinedRow.Efectiva === "Efectiva", 
                message: "No diligencia 12161/ " },
            { condition: combinedRow.Lectura !== "" && combinedRow.Efectiva === "Efectiva" && combinedRow.Correccion === "NO" && combinedRow.Artefacto_1 === "" && combinedRow["Paso/Acomulacion"] === "Sin Responder", 
                message: "No diligencia artefactos en 12161/ " },
            { condition: (resultado === "critico" || resultado === "no critico") && combinedRow.DESCRIPCION_DE_LOS_DEFECTOS === "", 
                message: "No diligencia descripcion de defectos en 12161/ " },
            { condition: combinedRow.Efectiva === "Efectiva" && combinedRow.Embebida === "NO" && combinedRow.A_la_vista === "SI" && combinedRow.Enterrada === "NO" && combinedRow.Conducto === "NO" && combinedRow.OTRO_TRAZADO === "NO" && combinedRow.Potencia_instalada_supera_la_considerada_en_el_disenio ==="N/A", 
                message:"No evalua potencia de disenio/ " },
            ];
// Después de definir validationRules, agrega este código:

        validationRules.forEach(rule => {
            if (rule.condition) {
                combinedRow.OBSERVACION_ERROR += rule.message;
            }
        });
        /*        
        if (Resultado === "certificada" && 
             (combinedRow.Continua_en_Servicio === "NO" || combinedRow.cumple_requisitos_y_criterios_verificados_resoluciones_90902_y_41385 === "NO"))
             {
             combinedRow.OBSERVACION_ERROR += "Error de diligenciamiento en continua en servicio o cumple 90902/ ";                      
             }
        else if((Resultado === "critico" || Resultado === "no critico")  && 
                combinedRow.cumple_requisitos_y_criterios_verificados_resoluciones_90902_y_41385 === "SI")
                {
                combinedRow.OBSERVACION_ERROR += "Error de diligenciamiento en cumple 90902/ ";                      
                }
        else if (Resultado === "critico" && combinedRow.Continua_en_Servicio === "SI")
        {
             combinedRow.OBSERVACION_ERROR += "Error de diligenciamiento en continua en servicio/ ";                      
         }
         if ((combinedRow.medidor_no_ventilado === "SI" ||
            combinedRow.fuga_en_cm === "SI" ||
            combinedRow.cm_aislamiento_eléctrico === "SI" ||
            combinedRow.cm_bajo_terreno === "SI" ||
            combinedRow.art_de_3_famila_en_semisotanos === "SI" ||
            combinedRow.distancia_a_elementos_de_combustion === "SI" ||
            combinedRow.secadora_sin_ducto === "SI" ||
            combinedRow.conector_flex_en_superf_calientes === "SI" ||
            combinedRow.espacio_entre_cielo_falso === "SI" ||
            combinedRow.metodo_acoplamiento === "SI" ||
            combinedRow.NOVEDADES_FI22 !== "") && (combinedRow.Resultado_de_la_Inspeccion === "CERTIFICADA" || combinedRow.Resultado_de_la_Inspeccion === "CERTIFICADO"))
            { 
                combinedRow.OBSERVACION_ERROR +=  "Error de cierre diligencia informe novedades FI-22/ ";
        }
        
          
        if (combinedRow.Resultado_de_la_Inspeccion === "CERTIFICADA" && combinedRow.NOVEDADES_FI_02 === "SI"
        ){
            combinedRow.OBSERVACION_ERROR += "Error de cierre marca novedades en FI-02/ ";
         }
        if (combinedRow.Embebida === "NO" && combinedRow.A_la_vista === "NO" && combinedRow.Enterrada === "NO" && combinedRow.Conducto === "NO" && combinedRow.OTRO_TRAZADO=== "NO" )
            {
            combinedRow.OBSERVACION_ERROR +="No diligencia trazado/ " ;   
            }
        if (combinedRow.Acero === "NO" && combinedRow.Cobre === "NO" && combinedRow.PEALPE === "NO" && combinedRow.PEMD === "NO" && combinedRow.CSST === "NO" && combinedRow.Otro_Material=== "NO" ) 
            {
            combinedRow.OBSERVACION_ERROR += "No diligencia material/ ";
            }
        if (combinedRow.Roscada === "NO" && combinedRow.Soldada === "NO" && combinedRow.abocinada === "NO" && combinedRow.anillada === "NO" 
        ) {
            combinedRow.OBSERVACION_ERROR += "No diligencia tipo de union/ ";
          }
        if (combinedRow.Mecanismo_de_control_de_sobrepresion_del_regulador_descarga_al_interior === "N/A" && combinedRow.Efectiva === "Efectiva"){
            combinedRow.OBSERVACION_ERROR += "No evalua mecanismo de sobrepresion/ ";
        }
        if (combinedRow.Correccion === "NO"){
            if (combinedRow.Lectura === "" && combinedRow.Efectiva === "Efectiva"){
                combinedRow.OBSERVACION_ERROR += "No diligencia 12161/ ";
            }
        }
        else {combinedData +=""}

        if (combinedRow.Lectura !== "" && 
            combinedRow.Efectiva === "Efectiva" && 
            combinedRow.Correccion === "NO" &&
            combinedRow.Artefacto_1 === "" && 
            combinedRow.Paso/Acomulacion ==="Sin Responder")
            {
            combinedRow.OBSERVACION_ERROR += "No diligencia artefactos en 12161/ ";
            }
        if (Resultado === "critico" || Resultado === "no critico" && combinedRow.DESCRIPCION_DE_LOS_DEFECTOS === "")
            {
                combinedRow.OBSERVACION_ERROR += "No diligencia descripcion de defectos en 12161/ ";
            }
        if  (combinedRow.Efectiva === "Efectiva" && combinedRow.Embebida ==="NO" && 
            combinedRow.A_la_vista === "SI" && combinedRow.Enterrada ==="NO" &&
            combinedRow.Conducto === "NO" && combinedRow.OTRO_TRAZADO === "NO" && 
            combinedRow.Potencia_instalada_supera_la_considerada_en_el_disenio ==="N/A")
            {
                combinedRow.OBSERVACION_ERROR += "no evalua potencia de disenio/ ";
            }
            
*/
        
        
        
        //

        // Add null checks for Medidor_OSF and Medidor_FI02, only for effective inspections
        if (combinedRow.Efectiva === "Efectiva") {
            if (combinedRow.Medidor_OSF && combinedRow.Medidor_FI02) {
                if (combinedRow.Medidor_OSF.includes(combinedRow.Medidor_FI02)) {
                    // No need to add anything to OBSERVACION_ERROR
                } else {
                    combinedRow.OBSERVACION_ERROR += "medidor errado/ ";
                }
            } else {
                combinedRow.OBSERVACION_ERROR += "";
            }
        }
        //
        


        if (combinedRow.OBSERVACION_ERROR === "") {
            combinedRow.OBSERVACION_ERROR = "Ok";
        }

        
        
        //agrega columna  ERROR
        combinedRow.ERROR ="";

        if (combinedRow.OBSERVACION_ERROR === "Ok") {
            combinedRow.ERROR ="Sin error"; 
        } else{
            combinedRow.ERROR="Error";
        }


        return combinedRow;
    });

    // Convertir los datos combinados de vuelta a CSV//-
    var csv = Papa.unparse(combinedData);//-
    var blob = new Blob([csv], { type: 'text/csv' });//-
    // Convertir los datos combinados a XLSX//+
    var wb = XLSX.utils.book_new();//+
    var ws = XLSX.utils.json_to_sheet(combinedData);//+
    XLSX.utils.book_append_sheet(wb, ws, "Sheet1");//+
//+
    // Generar el archivo XLSX//+
    var wbout = XLSX.write(wb, {bookType:'xlsx', type:'binary'});//+
//+
    // Convertir el archivo a Blob//+
    function s2ab(s) {//+
        var buf = new ArrayBuffer(s.length);//+
        var view = new Uint8Array(buf);//+
        for (var i=0; i<s.length; i++) view[i] = s.charCodeAt(i) & 0xFF;//+
        return buf;//+
    }//+
//+
    var blob = new Blob([s2ab(wbout)], {type:"application/octet-stream"});//+
    var url = URL.createObjectURL(blob);

    // Actualizar el enlace de descarga
    var downloadLink = document.getElementById('downloadLink');
    downloadLink.href = url;
    downloadLink.download = 'output.xlsx'; // Set the filename//+
    downloadLink.style.display = 'block';
}

