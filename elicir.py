import pandas as pd
import pandas as pd
import matplotlib.pyplot as plt
import numpy as np
import csv
import io
from datetime import datetime
F = pd.read_csv("main.csv")


#--------------eli-cb--------------- 
elicb = pd.DataFrame() 
elicb["Earthing Configuration"]= F["earth_config"]
elicb["device_make"] = F["elicb_device_make"]
elicb["device_type"]= F["elicb_device_type"]
elicb["device_sensitivity"]= F["device_sen"]
elicb["Device Rating (A)"]= F["elicb_device_rating"]
# TMS - TDS	eli_tms          not there
elicb["Type of Circuit Location"]=F["type_cir_loc"]
elicb["No.of Phases"]=F["no_phases"]
elicb["Trip Curve"]=F["b_curve_type"]
elicb["V_LN (V)"]=F["elicb_mvln"]
elicb["V_NE (V)"]=F["elicb_mvne"]
elicb["V_LE (V)"]=F["elicb_mvle"]
elicb["L1-ELI (O)"]=F["elicb_mel1"]
elicb["L2-ELI (O)"]=F["elicb_mel2"]
elicb["L3-ELI (O)"]=F["elicb_mel3"]
elicb["Psc (kA)"]=F["elicb_psc"]
elicb["Suggested Max ELI (O)"]=F["elicb_max_eli"]
elicb_filled = elicb.fillna("")
elicb["Device Rating (A)"] = elicb["Device Rating (A)"].astype(int)
new_column1 = []
result_column1 = []
P = 0
K = 0
TMS = 1
TDS = 1
elicb_filled = elicb.fillna("")
elicb["Device Rating (A)"] =elicb["Device Rating (A)"].astype(int)
Is = elicb["Device Rating (A)"]

gf1 = elicb[
    [
        "Earthing Configuration",
        "Type of Circuit Location",
        "Device Rating (A)",
        "device_make",
        "device_type",
        "device_sensitivity",
        "No.of Phases",
        "Trip Curve",
    ]
]

gf2 = elicb[
    [

        "V_LN (V)",
        "V_NE (V)",
        "V_LE (V)",
        "L1-ELI (O)",
        "L2-ELI (O)",
        "L1-ELI (O)",
        "Psc (kA)",
    ]
]


def eli_test_result1(gf1, gf2):
    I = Is * ((((K * TMS) / Td) + 1) ** (1 / P))
    if gf1["Earthing Configuration"] == "TN":
        IEC_val_TN = gf2["V_LE (V)"] / I
        new_column1.append(round(IEC_val_TN, 4))
        if gf1["No. of Phases"] == 3:
            if (
            gf2["L1-ELI (O)"] <= IEC_val_TN
                and gf2["L2-ELI (O)"] <= IEC_val_TN
                and gf2["L3-ELI (O)"] <= IEC_val_TN
            ):
                result_column1.append("Pass")
            else:
                result_column1.append("Fail")
        elif gf1["No.of Phases"] == 1:
            if gf2["L1-ELI (O)"] <= IEC_val_TN:
                result_column1.append("Pass")
            else:
                result_column1.append("Fail")
        else:
            result_column1.append("N/A")
    elif gf1["Earthing Configuration"] == "TT":
        IEC_val_TT = 50 / I
        new_column1.append(round(IEC_val_TT, 4))
        if gf1["No.of Phases"] == 3:
            if (
                gf2["L1-ELI (O)"] <= IEC_val_TT
                and gf2["L2-ELI (O)"] <= IEC_val_TT
                and gf2["L3-ELI (O)"] <= IEC_val_TT
            ):
                result_column1.append("Pass")
            else:
                result_column1.append("Fail")
        elif gf1["No.of Phases"] == 1:
            if gf2["L1-ELI (O)"] <= IEC_val_TT:
                result_column1.append("Pass")
            else:
                result_column1.append("Fail")
        else:
            result_column1.append("N/A")


def eli_test_result2(gf1, gf2):
    I = Is * (((((A / ((Td / TDS) - B)) + 1)) ** (1 / p)))
    if gf1["Earthing Configuration"] == "TN":
        IEEE_val_TN = gf2["V_LE (V)"] / I
        new_column1.append(round(IEEE_val_TN, 4))
        if gf1["No.of Phases"] == 3:
            if (
                gf2["L1-ELI (O)"] <= IEEE_val_TN
                and gf2["L2-ELI (O)"] <= IEEE_val_TN
                and gf2["L3-ELI (O)I"] <= IEEE_val_TN
            ):
                result_column1.append("Pass")
            else:
                result_column1.append("Fail")
        elif gf1["No.of Phases"] == 1:
            if gf2["L1-ELI (O)"] <= IEEE_val_TN:
                result_column1.append("Pass")
            else:
                result_column1.append("Fail")
        else:
            result_column1.append("N/A")
    elif gf1["Earthing Configuration"] == "TT":
        IEEE_val_TT = 50 / I
        new_column1.append(round(IEEE_val_TT, 4))
        if gf1["No.of Phases"] == 3:
            if (
                gf2["L1-ELI (O)"] <= IEEE_val_TT
                and gf2["L2-ELI (O)"] <= IEEE_val_TT
                and gf2["L3-ELI (O)"] <= IEEE_val_TT
            ):
                result_column1.append("Pass")
            else:
                result_column1.append("Fail")
        elif gf1["No.of Phases"] == 1:
            if gf2["L1-ELI (O)"] <= IEEE_val_TT:
                result_column1.append("Pass")
            else:
                result_column1.append("Fail")
        else:
            result_column1.append("N/A")
for index, row in elicb.iterrows():
    if row["device_type"] == "MCB":
        rating = gf1["Device Rating (A)"]
        trip = gf1.loc[gf1["Device Rating (A)"] == rating, "Trip Curve"].values
        if len(trip) > 0:
            val_MCB = trip[0]
        else:
            val_MCB = 0  # Set a default value when 'Trip Curve' value is not found in sugg-max-eli.csv


        if gf1["No.of Phases"] == 3:
            if gf2["L1-ELI (O)"] <= val_MCB and  gf2["L2-ELI (O)"] <= val_MCB and  gf2["L3-ELI (O)"] <= val_MCB:
                result_column1.append("Pass")
            else:
                result_column1.append("Fail")
        elif gf1["No.of Phases"]== 1:
            if  gf2["L1-ELI (O)"] <= val_MCB:
                result_column1.append("Pass")
            else:
                result_column1.append("Fail")
        else:
            result_column1.append("N/A")

    elif gf1["device_type"] in ["RCD", "RCBO", "RCCB"] and gf1["Earthing Configuration"] == "TN":
        rccb_val_TN = ("V_LE (V)"/ gf1["device_sensitivity"]) * 1000
        new_column1.append(round(rccb_val_TN, 4))
        if gf1["No.of Phases"] == 3:
            if (
                gf2["L1-ELI (O)"] <= rccb_val_TN
                and  gf2["L2-ELI (O)"] <= rccb_val_TN
                and  gf2["L3-ELI (O)"]<= rccb_val_TN
            ):
                result_column1.append("Pass")
            else:
                result_column1.append("Fail")
        elif gf1["No.of Phases"] == 1:
            if row["L1-ELI"] <= rccb_val_TN:
                result_column1.append("Pass")
            else:
                result_column1.append("Fail")
        else:
            result_column1.append("N/A")

    elif  gf1["device_type"] in ["RCD", "RCBO", "RCCB"] and  gf1["Earthing Configuration"]== "TT":
        rccb_val_TT = (50 / gf1["device_sensitivity"]) * 1000
        new_column1.append(round(rccb_val_TT, 4))
        if gf1["No.of Phases"] == 3:
            if (
                gf2["L1-ELI (O)"]<= rccb_val_TT
                and gf2["L2-ELI (O)"] <= rccb_val_TT
                and gf2["L3-ELI (O)"] <= rccb_val_TT
            ):
                result_column1.append("Pass")
            else:
                result_column1.append("Fail")
        elif row["No. of Phases"] == 1:
            if gf2["L1-ELI (O)"] <= rccb_val_TT:
                result_column1.append("Pass")
            else:
                result_column1.append("Fail")
        else:
            result_column1.append("N/A")

    elif gf1["device_type"] == "MCCB" or gf1["device_type"] == "ACB":
        if gf1["Type of Circuit Location"] == "Final":
            Td = 0.4
        elif gf1["Type of Circuit Location"] == "Distribution":
            Td = 5
        Is = gf1["Device Rating (A)"]
        if gf1["Trip Curve"] == "IEC Standard Inverse":
            P = 0.02
            K = 0.14
            eli_test_result1(gf1, gf2)
        elif gf1["Trip Curve"] == "IEC Very Inverse":
            P = 1
            K = 13.5
            eli_test_result1(gf1, gf2)
        elif gf1["Trip Curve"] == "IEC Long-Time Inverse":
            P = 1
            K = 120
            eli_test_result1(gf1, gf2)
        elif gf1["Trip Curve"] == "IEC Extremely Inverse":
            P = 2
            K = 80
            eli_test_result1(gf1, gf2)
        elif gf1["Trip Curve"] == "IEC Ultra Inverse":
            P = 2.5
            K = 315.2
            eli_test_result1(gf1, gf2)
        elif gf1["Trip Curve"] == "IEEE Moderately Inverse":
            A = 0.0515
            B = 0.114
            p = 0.02
            eli_test_result2(gf1, gf2)
        elif gf1["Trip Curve"] == "IEEE Very Inverse":
            A = 19.61
            B = 0.491
            p = 2
            eli_test_result2(gf1, gf2)
        elif gf1["Trip Curve"] == "IEEE Extremely Inverse":
            A = 28.2
            B = 0.1217
            p = 2
            eli_test_result2(gf1, gf2)

new_column1 = pd.Series(new_column1[: len(gf2)], name="Suggested Max ELI (Ω)")
gf2["Suggested Max ELI (Ω)"] = new_column1
gf2["Suggested Max ELI (Ω)"] = gf2["Suggested Max ELI (Ω)"].apply(lambda x: "{:.2f}".format(x))
result_column1 = pd.Series(result_column1[: len(gf2)], name="Result")
gf2["Result"] = result_column1
      