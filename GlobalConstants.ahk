#Requires AutoHotkey v2.0
 
    global G_CONSTANTS := {
        ; Definizione dei percorsi dei file per il controllo delle FL
		file_Country: A_ScriptDir . "\Config\country.csv",
		file_Tech: A_ScriptDir . "\Config\Technology.csv",
		file_Rules: A_ScriptDir . "\Config\Rules.csv",
		file_Mask: A_ScriptDir . "\Config\Mask_FL.csv",
        ; Definizione dei percorsi dei file Guideline per il controllo delle FL
		file_FL_Wind: A_ScriptDir . "\Config\Wind_FL_GuideLine.csv",
		file_FL_Bess: A_ScriptDir . "\Config\Bess_FL_GuideLine.csv",
		file_FL_Solar_Common: A_ScriptDir . "\Config\Solar_FL_Common_GuideLine.csv",
		file_FL_W_SubStation: A_ScriptDir . "\Config\Wind_FL_SubStation_Guideline.csv",
        file_FL_S_SubStation: A_ScriptDir . "\Config\Solar_FL_SubStation_Guideline.csv",
        file_FL_B_SubStation: A_ScriptDir . "\Config\Bess_FL_SubStation_Guideline.csv",
		file_FL_Solar_CentralInv: A_ScriptDir . "\Config\Solar_FL_CentrealInv_GuideLine.csv",
		file_FL_Solar_StringInv: A_ScriptDir . "\Config\Solar_FL_StringInv_GuideLine.csv",
		file_FL_Solar_InvModule: A_ScriptDir . "\Config\Solar_FL_InvModule_GuideLine.csv",
        ; Definizione del percorsi dei file per l'upload
        path_file_UpLoad: A_ScriptDir . "\FileUpLoad\",
        ; Definizione dei files per l'upload delle tabelle globali
		file_FL_2_UpLoad: A_ScriptDir . "\FileUpLoad\FL_2_UpLoad.csv",
		file_FL_n_UpLoad: A_ScriptDir . "\FileUpLoad\FL_n_UpLoad.csv", 
        ; Definizione del file per l'upload delle tabelle Control Asset
		file_ZPMR_CTRL_ASS_UpLoad: A_ScriptDir . "\FileUpLoad\ZPMR_CTRL_ASS.csv",
        ; Definizione del file per l'upload delle tabelle Technical Object
        file_ZPMR_TECH_OBJ_UpLoad: A_ScriptDir . "\FileUpLoad\ZPMR_TECH_OBJ.csv",
        ; Intestazione dei file per upload
        intestazione_FL_2: "TPLKZ;FLTYP;FLLEVEL;LAND1;VALUE`r`n",
        intestazione_FL_n: "TPLKZ;FLTYP;FLLEVEL;VALUE;VALUETX;REFLEVEL`r`n",
        ; Intestazione dei file per upload CTRL_ASS e OBJ
        intestazione_CTRL_ASS: "VALUE;SUB_VALUE;SUB_VALUE2;TPLKZ;FLTYP;FLLEVEL;CODE_SEZ_PM;CODE_SIST;CODE_PARTE;TIPO_ELEM`r`n",
        intestazione_OBJ: "VALUE;SUB_VALUE;SUB_VALUE2;TPLKZ;FLTYP;FLLEVEL;EQART;RBNR;NATURE;DOCUMENTO_PM`r`n",
        ; timeout operazioni in SAP
        timeoutSeconds: 30,
        ; Nomi colonne tabella CTRL_ASS - Control Asset
        CTRL_ASS_Valore_Livello : "Valore Livello",
        CTRL_ASS_Valore_Liv_Superiore_1 : "Valore Liv. Superiore",
        CTRL_ASS_Valore_Liv_Superiore_2 : "Valore Liv. Superiore",
        CTRL_ASS_LivelloSedeTecnica : "Liv.Sede"
        }

        /* 
        Esempio di utilizzo:
        G_CONSTANTS.test
        */