import pandas as pd
import csv
from docx import Document
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from docx.shared import Inches, Pt
import pandas as pd
import matplotlib.pyplot as plt
from docx import Document
import numpy as np
from docx.shared import Inches
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT, WD_ALIGN_VERTICAL
from tabulate import tabulate
from docx.shared import RGBColor
from docx.oxml.ns import nsdecls
from docx.oxml import parse_xml

F = pd.read_csv("Fcsv.csv")

F["app_test_vlt"] = pd.to_numeric(F["app_test_vlt"], errors="coerce")
F["meas_out_curr_ma"] = pd.to_numeric(F["meas_out_curr_ma"], errors="coerce")

irdf = pd.DataFrame()  # IR
irdf["No. Poles"] = F["no_poles"]
irdf["SPD Applicable"] = F["spd_applicable"]
irdf["Nominal Circuit Voltage (V)"] = F["nom_cir_vlt"]
irdf["Measurement Terminals"] = F["meas_term"]
irdf["Test Voltage (V)"] = F["test_vlt"]
irdf["Conductor Type"] = F["cond_type"]
irdf["Conductor Size (sq. mm)"] = F["cond_size"]
irdf["Number of Runs"] = F["no_runs"]
irdf["Insulation Type"] = F["ins_type"]
irdf["Leakage Capacitance (nF)"] = F["leak_cap_nf"]
irdf["Insulation Resistance (MO)"] = F["ins_res_mohm"]

fwrdf = pd.DataFrame()  # FWR
# fwrdf["Location Point"] = F["loc_point"]
# fwrdf["Reference Point"] = F["ref_point"]
fwrdf["Distance from previous test location (m)"] = F["distance_prev_loc"]
fwrdf["Nominal Voltage to Earth of System (V)"] = F["nom_vlt_e_v"]
fwrdf["Applied Test Voltage (V)"] = F["app_test_vlt"]
fwrdf["Measured Output Current (mA)"] = F["meas_out_curr_ma"]
fwrdf["Resistance (kO)"] = F["res_kohm"]

rcdf = pd.DataFrame()  # RC
rcdf["Conductor Location - From"] = F["rc_route_ele_1"]
rcdf["Conductor Location - To"] = F["rc_route_ele_2"]
rcdf["No of runs of Conductor"] = F["no_runs"]
rcdf["Conductor Type"] = F["cond_type"]
rcdf["Conductor Size (sq. mm)"] = F["cond_size"]
rcdf["Conductor Length (m)"] = F["cond_len"]
rcdf["Conductor Temperature (°C)"] = F["cond_temp"]
rcdf["Is Continuity found?"] = F["cont_found"]
rcdf["Lead Internal Resistance (O)"] = F["lead_int_res"]
rcdf["Continuity Resistance (O)"] = F["con_res"]
rcdf["Corrected Continuity Resistance (O)"] = F["corr_cont_res"]
rcdf["Specific Conductor Resistance (MO/m) at 30°C"] = F["spec_cond_res"]

opdf = pd.DataFrame()  # OP
opdf["VL1-L2 (V)"] = F["op_l1_l2_v"]
opdf["VL2-L3 (V)"] = F["op_l2_l3_v"]
opdf["VL3-L1 (V)"] = F["op_l3_l1_v"]
opdf["VL1-N (V)"] = F["op_l1_n_v"]
opdf["VL2-N (V)"] = F["op_l2_n_v"]
opdf["VL3-N (V)"] = F["op_l3_n_v"]
opdf["Phase Sequence"] = F["phase_seq"]

vddf = pd.DataFrame()  # VD
vddf["Circuit Route 'From'"] = F["vd_route_ele_1"]
vddf["Circuit Route 'To'"] = F["vd_route_ele_2"]
vddf["Measured Voltage (V, L-N) 'From'"] = F["ele_1_l_n_v"]
vddf["Measured Voltage (V, L-N) 'To'"] = F["ele_2_l_n_v"]
vddf["Nominal Circuit Voltage (V)"] = F["nom_cir_vlt"]
vddf["Type of Installation Supply System"] = F["type_install_supply"]
vddf["Purpose of Supply"] = F["pur_supply"]
vddf["Conductor Type"] = F["cond_type"]
vddf["Insulation Type"] = F["ins_type"]
vddf["Cable Length (m)"] = F["cond_len"]
vddf["Calculated Voltage Drop (V)"] = F["vlt_drop_v"]
vddf["Voltage Drop %"] = F["vlt_drop_perc"]

ptdf = pd.DataFrame()  # PT
ptdf["Device Type"] = F["pt_device_type"]
ptdf["Type of Supply"] = F["type_supply"]
ptdf["Line to Neutral Voltage (V)"] = F["l_n_v"]
ptdf["Polarity Reference"] = F["pol_ref"]
# ptdf["Polarity"] = F["pol"]

rdcdf = pd.DataFrame()  # RD
rdcdf["Type of Voltage Waveform"] = F["type_supply"]
rdcdf["Type of Earthing System"] = F["type_earth_sys"]
rdcdf["Nominal Line to Earth Voltage (V)"] = F["nom_l_e_v"]
rdcdf["Nominal Current Rating(A)"] = F["nom_curr_rating"]
rdcdf["Rated Residual Operating Current,IΔn (mA)"] = F["rated_res_op_curr"]
rdcdf["Application Type"] = F["appl_type"]
rdcdf["Trip curve type"] = F["trip_curve_type"]
rdcdf["No. of Poles"] = F["no_poles"]
rdcdf["Test Current (mA)"] = F["test_curr_ma"]
rdcdf["Trip Current (mA)"] = F["trip_curr_ma"]
rdcdf["Trip Time (ms)"] = F["trip_time"]
rdcdf["Device Tripped"] = F["device_trip"]
# rdcdf["Result"] = F["rcd_result"]

epedf = pd.DataFrame()  # EPE
epedf["No of Parallel Electrodes"] = F["no_eecp"]
epedf["Earthing Application"] = F["ea"]
epedf["Type of Earthing"] = F["type_earthing"]
epedf["Earth Electrode Depth (m)"] = F["depth_iee"]
epedf["Nearest Electrode Distance (m)"] = F["d_nee"]
epedf["Measured Earth Resistance - Individual (O)"] = F["meri"]
epedf["Calculated Earth Resistance - Individual (O)"] = F["ceri"]
epedf["Electrode Distance Ratio"] = F["elec_dis_ratio"]
# epedf["Result"] = F["epe_result"]

tpsdf = pd.DataFrame()  # TPS
tpsdf["Rated Line Voltage (V)"] = F["rated_line_vlt"]
tpsdf["Voltage-L1L2 (V)"] = F["tps_l1_l2_v"]
tpsdf["Voltage-L2L3 (V)"] = F["tps_l2_l3_v"]
tpsdf["Voltage-L3L1 (V)"] = F["tps_l3_l1_v"]
tpsdf["Voltage-L1N (V)"] = F["tps_l1_n_v"]
tpsdf["Voltage-L2N (V)"] = F["tps_l2_n_v"]
tpsdf["Voltage-L3N (V)"] = F["tps_l3_n_v"]
tpsdf["Average Line Voltage (V)"] = F["avg_line_vlt"]
tpsdf["Average Phase Voltage (V)"] = F["avg_ph_vlt"]
tpsdf["Voltage Unbalance %"] = F["vlt_unb"]
# tpsdf["Voltage Result"] = F["tps_vlt_result"]
tpsdf["Rated Phase Current (A)"] = F["rated_ph_curr"]
tpsdf["Current-L1 (A)"] = F["tps_curr_l1"]
tpsdf["Current-L2 (A)"] = F["tps_curr_l2"]
tpsdf["Current-L3 (A)"] = F["tps_curr_l3"]
tpsdf["Average Phase Current (A)"] = F["avg_ph_curr"]
tpsdf["Current Unbalance %"] = F["curr_unb"]
# tpsdf["Current Result"] = F["tps_curr_result"]
tpsdf["Voltage-NE (V)"] = F["vlt_nev"]
# tpsdf["NEV Result"] = F["tps_nev_result"]
tpsdf["Zero Sum Current (mA)"] = F["zero_sum_curr"]
# tpsdf["ZeroSum Result"] = F["tps_zs_result"]
# tpsdf["Remarks"] = F["tps_result_remarks"]

elicbdf = pd.DataFrame()  # ELI CB
elicbdf["Earthing Configuration"] = F["earth_config"]
elicbdf["Device Make"] = F["elicb_device_make"]
elicbdf["Device Type"] = F["elicb_device_type"]
elicbdf["Device Sensitivity"] = F["device_sen"]
elicbdf["Device Rating (A)"] = F["elicb_device_rating"]
elicbdf["TMS - TDS"] = F["eli_tms"]
elicbdf["Type of Circuit Location"] = F["type_cir_loc"]
elicbdf["No.of Phases"] = F["no_phases"]
elicbdf["Trip Curve"] = F["b_curve_type"]
elicbdf["V_LN (V)"] = F["elicb_mvln"]
elicbdf["V_LE (V)"] = F["elicb_mvle"]
elicbdf["V_NE (V)"] = F["elicb_mvne"]
elicbdf["L1-ELI (O)"] = F["elicb_mel1"]
elicbdf["L2-ELI (O)"] = F["elicb_mel2"]
elicbdf["L3-ELI (O)"] = F["elicb_mel3"]
elicbdf["Psc (kA)"] = F["elicb_psc"]
elicbdf["Suggested Max ELI (O)"] = F["elicb_max_eli"]
# elicbdf["Result"] = F["elicb_result"]

elisocdf = pd.DataFrame()  # ELI SOCKET
elisocdf["Earthing Configuration"] = F["earth_config"]
elisocdf["Socket Rating (A)"] = F["soc_ar"]
elisocdf["Socket Type"] = F["soc_type"]
elisocdf["Upstream Breaker Name"] = F["up_b_name"]
elisocdf["Upstream Breaker Make"] = F["up_b_make"]
elisocdf["Upstream Breaker Type"] = F["up_b_type"]
elisocdf["Upstream Breaker Sensitivity"] = F["up_b_sen"]
elisocdf["Upstream Breaker Rating (A)"] = F["up_b_sen"]
elisocdf["Upstream Breaker TMS - TDS"] = F["up_b_tms"]
elisocdf["Type of Circuit Location"] = F["type_cir_loc"]
elisocdf["No.of Phases"] = F["no_phases"]
elisocdf["Trip Curve"] = F["b_curve_type"]
elisocdf["V_LN (V)"] = F["elis_mvln"]
elisocdf["V_LE (V)"] = F["elis_mvle"]
elisocdf["V_NE (V)"] = F["elis_mvne"]
elisocdf["L1-ELI (O)"] = F["elis_mel1"]
elisocdf["L2-ELI (O)"] = F["elis_mel2"]
elisocdf["L3-ELI (O)"] = F["elis_mel3"]
elisocdf["Psc (kA)"] = F["elis_psc"]
elisocdf["Suggested Max ELI (O)"] = F["elis_max_eli"]
# elisocdf["Result"] = F["elis_result"]

fodf = pd.DataFrame()  # FUNC OPS
fodf["Device type"] = F["fo_device_type"]
fodf["Functional Check"] = F["func_check"]
fodf["Interlock check"] = F["inter_check"]
# fodf["Result"] = F["fo_result"]
# fodf["Remarks"] = F["fo_result_remark"]

patdf = pd.DataFrame()  # PAT
patdf["Device ID"] = F["pat_device_id"]
patdf["Device Name"] = F["pat_dev_name"]
patdf["Location"] = F["pat_location"]
patdf["Voltage Rating (V)"] = F["pat_vlt_rating"]
patdf["Fuse Rating (A)"] = F["pat_fuse_rating"]
patdf["Visual Inspection"] = F["pat_vi"]
patdf["Earth Continuity (O)"] = F["pat_ec"]
patdf["Insulation Resistance (MO)"] = F["pat_ir"]
patdf["Polarity Test"] = F["pat_pt"]
patdf["Leakage (mA)"] = F["pat_leak"]
patdf["Functional Check"] = F["pat_func_check"]
# patdf["Overall Result"] = F["result_pat"]


def find_loc_name(loc_id, locations):
    while loc_id in locations:
        loc_data = locations[loc_id]
        loc_name = loc_data["loc_name"]
        granularity = int(loc_data["granularity"])
        loc_type = loc_data["type"]

        if granularity == 1:
            if loc_type != "fac_area":
                parent_id = loc_data["parent_id"]
                parent_loc_name = find_loc_name(parent_id, locations)
                return parent_loc_name if parent_loc_name else "Parent Location"
            else:
                return loc_name
        else:
            loc_id = loc_data["parent_id"]
    return None


def create_dataframe(F):
    with open("your_file.csv", "r") as file:
        reader = csv.DictReader(file)
        locations = {row["loc_id"]: row for row in reader}
    table_data = []
    for index, row in F.iterrows():
        loc_id = str(row["loc_id"])
        result = find_loc_name(loc_id, locations)
        if result:
            loc_name = locations[loc_id]["loc_name"]
            parent_id = locations[loc_id]["parent_id"]
            parent_loc_name = locations[parent_id]["loc_name"]
            table_data.append([loc_name, parent_loc_name if parent_loc_name else "", result])
    df = pd.DataFrame(table_data, columns=["Location", "Parent Location", "Facility Area"])
    return df


table_df = create_dataframe(F)
print(table_df)
print(opdf)